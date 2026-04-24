#!/usr/bin/env node
/**
 * Static module-graph validator for the Thunderbird MCP extension.
 *
 * Walks every .js / .mjs / .sys.mjs / .cjs file under extension/, tokenizes each
 * file (no regex against raw source -- a proper JS lexer that respects strings,
 * template literals, regex literals, line comments, and block comments), then
 * scans the token stream to extract:
 *
 *   - static `import` declarations (default, named, namespace, side-effect)
 *   - dynamic `ChromeUtils.importESModule("...")` calls, including the
 *     destructured `const { foo } = ChromeUtils.importESModule(...)` form
 *     and its multi-line variants
 *   - exports (`export const|let|var|function|class|default`, `export { ... }`,
 *     `export { ... } from "..."`, `export * from "..."`)
 *
 * For every import targeting `resource://thunderbird-mcp/...` it then:
 *
 *   - resolves the URL to a file under extension/ (cachebust query strings stripped)
 *   - verifies the file exists
 *   - if named bindings are imported, verifies the target actually exports them
 *
 * Builds the directed graph of in-repo imports and detects cycles via DFS coloring.
 *
 * Exit code 0 on clean, 1 on any finding.
 *
 * Zero npm dependencies on purpose. The previous regex-based version choked on
 * JSDoc comments and on URL strings containing "//" inside .sys.mjs source --
 * the tokenizer below sidesteps both classes of false negative.
 */

"use strict";

const fs = require("fs");
const path = require("path");

const ROOT = path.resolve(__dirname, "..");
const EXT_DIR = path.join(ROOT, "extension");
const RES_PREFIX = "resource://thunderbird-mcp/";

// ---------------------------------------------------------------------------
// File walker
// ---------------------------------------------------------------------------

function walk(dir, out) {
  out = out || [];
  for (const entry of fs.readdirSync(dir, { withFileTypes: true })) {
    const full = path.join(dir, entry.name);
    if (entry.isDirectory()) {
      walk(full, out);
    } else if (/\.(?:js|mjs|cjs)$/.test(entry.name)) {
      // .sys.mjs matches .mjs above; .cjs/.js/.mjs all included.
      out.push(full);
    }
  }
  return out;
}

function relExt(p) {
  // Always print POSIX-style separators so output is identical on win/mac/linux.
  return path.relative(ROOT, p).split(path.sep).join("/");
}

function resolveResourceUrl(url) {
  if (!url.startsWith(RES_PREFIX)) return null;
  const tail = url.slice(RES_PREFIX.length).split("?")[0].split("#")[0];
  // Use the same separator the OS uses on disk; relExt() normalises for output.
  return path.join(EXT_DIR, ...tail.split("/"));
}

// ---------------------------------------------------------------------------
// Tokenizer
//
// Produces a flat array of tokens. We don't build a full AST -- the import /
// export / ChromeUtils.importESModule patterns are simple enough that a token
// stream walker handles them with high confidence, as long as the lexer is
// honest about strings, regex literals, templates, and comments.
//
// Token shapes:
//   { type: "punct",    value: "(" | ")" | "{" | "}" | "[" | "]" | "," |
//                                ";" | ":" | "." | "=" | "..." | "*" | "=>" }
//   { type: "ident",    value: "<name>" }            // identifiers + keywords
//   { type: "string",   value: "<unquoted contents>" } // single, double, or
//                                                     // backtick-without-${
//   { type: "template", value: "<raw>" }              // template w/ ${...}
//   { type: "regex",    value: "/.../flags" }
//   { type: "number",   value: "..." }
//   { type: "other",    value: "<char>" }             // operators we don't care about
//
// We only need enough to disambiguate `import` / `export` / `ChromeUtils.
// importESModule` and the surrounding curly/bracket/paren structure. Numbers,
// regex literals, and unhandled operators are recorded so `/` after an
// identifier is correctly treated as division (not a regex literal start).
// ---------------------------------------------------------------------------

const KEYWORD_BEFORE_REGEX = new Set([
  "return", "typeof", "instanceof", "in", "of", "new", "delete", "void",
  "throw", "case", "yield", "await", "do", "else",
]);

function isIdentStart(ch) {
  return (ch >= "a" && ch <= "z") || (ch >= "A" && ch <= "Z") ||
         ch === "_" || ch === "$";
}
function isIdentPart(ch) {
  return isIdentStart(ch) || (ch >= "0" && ch <= "9");
}
function isDigit(ch) {
  return ch >= "0" && ch <= "9";
}

function tokenize(src, file) {
  const tokens = [];
  const n = src.length;
  let i = 0;

  // After certain tokens, a `/` begins a regex literal; after others it's
  // division. Track the last "real" token to make the call.
  function regexAllowed() {
    for (let k = tokens.length - 1; k >= 0; k--) {
      const t = tokens[k];
      if (t.type === "punct") {
        // After most punctuation a regex is allowed (e.g. `(`, `,`, `=`, `;`).
        // After `)` or `]` it's NOT (those end an expression).
        if (t.value === ")" || t.value === "]") return false;
        return true;
      }
      if (t.type === "ident") {
        return KEYWORD_BEFORE_REGEX.has(t.value);
      }
      if (t.type === "number" || t.type === "string" ||
          t.type === "template" || t.type === "regex") {
        return false;
      }
      if (t.type === "other") {
        // Most "other" operators allow a regex to follow (e.g. `+`, `-`, `!`, `?`).
        return true;
      }
    }
    return true; // start of file
  }

  while (i < n) {
    const ch = src[i];

    // Whitespace + line terminators
    if (ch === " " || ch === "\t" || ch === "\n" || ch === "\r" ||
        ch === "\v" || ch === "\f") {
      i++;
      continue;
    }

    // Line comment
    if (ch === "/" && src[i + 1] === "/") {
      i += 2;
      while (i < n && src[i] !== "\n") i++;
      continue;
    }

    // Block comment (incl. JSDoc)
    if (ch === "/" && src[i + 1] === "*") {
      i += 2;
      while (i < n && !(src[i] === "*" && src[i + 1] === "/")) i++;
      i += 2; // skip closing */
      continue;
    }

    // Hashbang at file start
    if (ch === "#" && src[i + 1] === "!" && i === 0) {
      while (i < n && src[i] !== "\n") i++;
      continue;
    }

    // String literal (single or double quoted)
    if (ch === '"' || ch === "'") {
      const quote = ch;
      i++;
      let value = "";
      while (i < n && src[i] !== quote) {
        if (src[i] === "\\") {
          // Preserve the escaped char literally; we only need string contents
          // for source URLs and those are plain ASCII paths.
          value += src[i + 1] || "";
          i += 2;
        } else if (src[i] === "\n") {
          // Unterminated string on this line -- bail with original behavior.
          break;
        } else {
          value += src[i++];
        }
      }
      i++; // closing quote
      tokens.push({ type: "string", value: value });
      continue;
    }

    // Template literal. We capture the raw text including ${...} but for our
    // purposes templates are never used as import sources, so we just consume.
    if (ch === "`") {
      i++;
      let value = "";
      let depth = 0;
      while (i < n) {
        if (depth === 0 && src[i] === "`") { i++; break; }
        if (src[i] === "\\") { value += src[i] + (src[i + 1] || ""); i += 2; continue; }
        if (depth === 0 && src[i] === "$" && src[i + 1] === "{") {
          value += "${"; i += 2; depth = 1;
          while (i < n && depth > 0) {
            if (src[i] === "{") depth++;
            else if (src[i] === "}") depth--;
            value += src[i++];
          }
          continue;
        }
        value += src[i++];
      }
      tokens.push({ type: "template", value: value });
      continue;
    }

    // Regex literal (only if regex is allowed in this position)
    if (ch === "/" && regexAllowed()) {
      let j = i + 1;
      let inClass = false;
      while (j < n) {
        const c = src[j];
        if (c === "\\") { j += 2; continue; }
        if (c === "[") { inClass = true; j++; continue; }
        if (c === "]") { inClass = false; j++; continue; }
        if (c === "/" && !inClass) { j++; break; }
        if (c === "\n") break; // unterminated -- treat as division fallback
        j++;
      }
      // flags
      while (j < n && isIdentPart(src[j])) j++;
      tokens.push({ type: "regex", value: src.slice(i, j) });
      i = j;
      continue;
    }

    // Identifier / keyword
    if (isIdentStart(ch)) {
      let j = i + 1;
      while (j < n && isIdentPart(src[j])) j++;
      tokens.push({ type: "ident", value: src.slice(i, j) });
      i = j;
      continue;
    }

    // Number (including 0x, 0o, 0b, decimals, exponents -- we don't need to be
    // exact, just enough that a numeric literal doesn't bleed into the next
    // token). BigInt suffix `n` handled by the ident-part loop.
    if (isDigit(ch) || (ch === "." && isDigit(src[i + 1]))) {
      let j = i + 1;
      while (j < n && (isIdentPart(src[j]) || src[j] === ".")) j++;
      tokens.push({ type: "number", value: src.slice(i, j) });
      i = j;
      continue;
    }

    // Spread operator
    if (ch === "." && src[i + 1] === "." && src[i + 2] === ".") {
      tokens.push({ type: "punct", value: "..." });
      i += 3;
      continue;
    }

    // Arrow
    if (ch === "=" && src[i + 1] === ">") {
      tokens.push({ type: "punct", value: "=>" });
      i += 2;
      continue;
    }

    // Single-char punct we care about for parsing
    if ("(){}[],;:.=*".indexOf(ch) >= 0) {
      tokens.push({ type: "punct", value: ch });
      i++;
      continue;
    }

    // Everything else: lump as "other" so regex-allowed-after logic stays sane.
    // Multi-char operators don't matter to us; we record one char at a time.
    tokens.push({ type: "other", value: ch });
    i++;
  }

  return tokens;
}

// ---------------------------------------------------------------------------
// Token-stream scanner -- finds imports, dynamic imports, and exports.
// ---------------------------------------------------------------------------

function findMatching(tokens, openIdx, openChar, closeChar) {
  // tokens[openIdx] is the opening punct; return idx of matching close, or -1.
  let depth = 0;
  for (let k = openIdx; k < tokens.length; k++) {
    const t = tokens[k];
    if (t.type !== "punct") continue;
    if (t.value === openChar) depth++;
    else if (t.value === closeChar) {
      depth--;
      if (depth === 0) return k;
    }
  }
  return -1;
}

function isTopLevelAt(tokens, idx) {
  // True iff tokens[idx] is not nested inside any (), {}, or [].
  let p = 0, b = 0, s = 0;
  for (let k = 0; k < idx; k++) {
    const t = tokens[k];
    if (t.type !== "punct") continue;
    if (t.value === "(") p++;
    else if (t.value === ")") p--;
    else if (t.value === "{") b++;
    else if (t.value === "}") b--;
    else if (t.value === "[") s++;
    else if (t.value === "]") s--;
  }
  return p === 0 && b === 0 && s === 0;
}

function parseImportSpecifiers(tokens, openBraceIdx) {
  // `import { a, b as c }` -> source names are ["a", "b"]
  const close = findMatching(tokens, openBraceIdx, "{", "}");
  if (close < 0) return { sourceNames: [], closeIdx: tokens.length };
  const names = [];
  let k = openBraceIdx + 1;
  while (k < close) {
    const t = tokens[k];
    if (t.type === "ident") {
      names.push(t.value);
      // Skip optional `as <local>`
      if (tokens[k + 1] && tokens[k + 1].type === "ident" && tokens[k + 1].value === "as") {
        k += 3; // ident, "as", local
      } else {
        k++;
      }
      // Skip trailing comma if present
      if (tokens[k] && tokens[k].type === "punct" && tokens[k].value === ",") k++;
    } else {
      k++;
    }
  }
  return { sourceNames: names, closeIdx: close };
}

function parseDestructureSourceNames(tokens, openBraceIdx) {
  // `const { a, b: c, d = 1 } = ...` -> source props are ["a", "b", "d"]
  // (Property names on the LEFT of `:`. With no `:`, the ident IS the prop.)
  const close = findMatching(tokens, openBraceIdx, "{", "}");
  if (close < 0) return { sourceNames: [], closeIdx: tokens.length };
  const names = [];
  let k = openBraceIdx + 1;
  while (k < close) {
    const t = tokens[k];
    if (t.type === "punct" && t.value === "...") { k++; continue; }
    if (t.type !== "ident") { k++; continue; }
    const propName = t.value;
    let j = k + 1;
    // Quoted prop names ("foo": bar) -- we'd see a string before the ident; skip.
    // We already pushed the ident as propName; check if next token is `:` (alias).
    if (tokens[j] && tokens[j].type === "punct" && tokens[j].value === ":") {
      // The actual SOURCE property is propName; alias is whatever follows -- skip past it.
      // The alias may itself be a destructuring target {} or [], not just an ident.
      names.push(propName);
      j++;
      // Skip the alias subtree -- could be ident, {...}, or [...].
      if (tokens[j] && tokens[j].type === "punct" && tokens[j].value === "{") {
        j = findMatching(tokens, j, "{", "}") + 1;
      } else if (tokens[j] && tokens[j].type === "punct" && tokens[j].value === "[") {
        j = findMatching(tokens, j, "[", "]") + 1;
      } else {
        j++;
      }
    } else {
      names.push(propName);
    }
    // Skip optional default: `= <expr>`
    if (tokens[j] && tokens[j].type === "punct" && tokens[j].value === "=") {
      // Skip until the next top-level comma or the closing `}`.
      let depth = 0;
      j++;
      while (j < close) {
        const tt = tokens[j];
        if (tt.type === "punct") {
          if (tt.value === "(" || tt.value === "{" || tt.value === "[") depth++;
          else if (tt.value === ")" || tt.value === "}" || tt.value === "]") depth--;
          else if (tt.value === "," && depth === 0) break;
        }
        j++;
      }
    }
    if (tokens[j] && tokens[j].type === "punct" && tokens[j].value === ",") j++;
    k = j;
  }
  return { sourceNames: names, closeIdx: close };
}

function scan(tokens, file) {
  const refs = [];   // { url, names: string[]|null, kind: "static"|"dynamic" }
  const exports = new Set();

  for (let i = 0; i < tokens.length; i++) {
    const t = tokens[i];
    if (t.type !== "ident") continue;

    // ------------------------------------------------------------------
    // Static `import ...` (only valid at top level in ESM, and we only
    // care about top-level here -- nested `import(...)` is dynamic ESM,
    // not used in this codebase, but we handle it below as well).
    // ------------------------------------------------------------------
    if (t.value === "import" && isTopLevelAt(tokens, i)) {
      // Look ahead for the `from "..."` source, or a bare `import "..."`.
      // Patterns:
      //   import "src";
      //   import X from "src";
      //   import * as X from "src";
      //   import { a, b as c } from "src";
      //   import X, { a } from "src";
      //   import X, * as Y from "src";
      let j = i + 1;
      const named = [];
      let sawDefault = false;
      let sawNamespace = false;

      // Bare `import "src"`
      if (tokens[j] && tokens[j].type === "string") {
        refs.push({ url: tokens[j].value, names: null, kind: "static" });
        i = j;
        continue;
      }

      while (j < tokens.length) {
        const tt = tokens[j];
        if (tt.type === "ident" && tt.value === "from") { j++; break; }
        if (tt.type === "string") break; // shouldn't hit before `from`, but bail
        if (tt.type === "punct" && tt.value === ";") { j = -1; break; }
        if (tt.type === "ident" && tt.value !== "as") {
          sawDefault = true;
          j++;
          continue;
        }
        if (tt.type === "punct" && tt.value === "*") {
          sawNamespace = true;
          j++;
          continue;
        }
        if (tt.type === "punct" && tt.value === "{") {
          const parsed = parseImportSpecifiers(tokens, j);
          for (const n of parsed.sourceNames) named.push(n);
          j = parsed.closeIdx + 1;
          continue;
        }
        if (tt.type === "punct" && tt.value === ",") { j++; continue; }
        if (tt.type === "ident" && tt.value === "as") { j += 2; continue; } // skip alias
        j++;
      }
      if (j > 0 && tokens[j] && tokens[j].type === "string") {
        // For static imports we record the named bindings we extracted. If the
        // import was default-only or namespace-only, we record `null` (target
        // exists check still runs, but no per-name check).
        const url = tokens[j].value;
        const names = named.length > 0 ? named : null;
        refs.push({ url, names, kind: "static" });
        i = j;
      }
      continue;
    }

    // ------------------------------------------------------------------
    // `export ...`
    // ------------------------------------------------------------------
    if (t.value === "export" && isTopLevelAt(tokens, i)) {
      const next = tokens[i + 1];
      if (!next) continue;

      // `export default ...`
      if (next.type === "ident" && next.value === "default") {
        exports.add("default");
        continue;
      }

      // `export * from "src"` or `export * as X from "src"`
      if (next.type === "punct" && next.value === "*") {
        // Re-export -- doesn't add named exports we can know about statically.
        // (Could be expanded by following the `from`-target's exports, but
        // not needed for our codebase.)
        continue;
      }

      // `export { a, b as c }` or `export { a } from "src"`
      if (next.type === "punct" && next.value === "{") {
        const close = findMatching(tokens, i + 1, "{", "}");
        if (close > 0) {
          let k = i + 2;
          while (k < close) {
            const tk = tokens[k];
            if (tk.type === "ident") {
              // local name; check for `as <exported>`
              let exported = tk.value;
              if (tokens[k + 1] && tokens[k + 1].type === "ident" && tokens[k + 1].value === "as") {
                if (tokens[k + 2] && tokens[k + 2].type === "ident") {
                  exported = tokens[k + 2].value;
                }
                k += 3;
              } else {
                k++;
              }
              exports.add(exported);
              if (tokens[k] && tokens[k].type === "punct" && tokens[k].value === ",") k++;
            } else {
              k++;
            }
          }
        }
        continue;
      }

      // `export const|let|var <name>` (also handles destructuring patterns,
      // though we only need to surface what's exported as a named binding;
      // for our codebase plain ident declarations are universal).
      // `export function <name>` / `export async function <name>`
      // `export class <name>`
      if (next.type === "ident") {
        let k = i + 1;
        if (next.value === "async") k++;
        const decl = tokens[k];
        if (!decl) continue;
        if (decl.type === "ident" &&
            (decl.value === "const" || decl.value === "let" || decl.value === "var" ||
             decl.value === "function" || decl.value === "class")) {
          // Walk declarators: `const a = ..., b = ...`
          let m = k + 1;
          if (decl.value === "function" && tokens[m] && tokens[m].type === "punct" && tokens[m].value === "*") {
            m++;
          }
          while (m < tokens.length) {
            const tm = tokens[m];
            if (!tm) break;
            if (tm.type === "ident") {
              exports.add(tm.value);
              // For function/class, only the first ident is the name.
              if (decl.value === "function" || decl.value === "class") break;
              // For const/let/var, skip past initializer to next comma at this depth.
              let depth = 0;
              m++;
              while (m < tokens.length) {
                const tt = tokens[m];
                if (tt.type === "punct") {
                  if (tt.value === "(" || tt.value === "{" || tt.value === "[") depth++;
                  else if (tt.value === ")" || tt.value === "}" || tt.value === "]") depth--;
                  else if (tt.value === "," && depth === 0) break;
                  else if (tt.value === ";" && depth === 0) { m = tokens.length; break; }
                }
                m++;
              }
              if (m < tokens.length && tokens[m].type === "punct" && tokens[m].value === ",") {
                m++;
                continue;
              }
              break;
            }
            // Bail on anything we don't recognize as a declarator name.
            break;
          }
        }
      }
      continue;
    }

    // ------------------------------------------------------------------
    // ChromeUtils.importESModule(...)
    //   - bare:        ChromeUtils.importESModule("src")
    //   - destructured (any of):
    //         const { a, b } = ChromeUtils.importESModule("src");
    //         let   { a, b } = ChromeUtils.importESModule(\n "src" \n);
    //         var   { a, b: c } = ChromeUtils.importESModule("src");
    //         ({ a } = ChromeUtils.importESModule("src"));
    // ------------------------------------------------------------------
    if (t.value === "ChromeUtils" &&
        tokens[i + 1] && tokens[i + 1].type === "punct" && tokens[i + 1].value === "." &&
        tokens[i + 2] && tokens[i + 2].type === "ident" && tokens[i + 2].value === "importESModule" &&
        tokens[i + 3] && tokens[i + 3].type === "punct" && tokens[i + 3].value === "(") {
      // The first argument should be a string literal. Walk forward, skipping
      // any whitespace already eaten by the tokenizer.
      const close = findMatching(tokens, i + 3, "(", ")");
      let urlTok = null;
      for (let k = i + 4; k < close; k++) {
        if (tokens[k].type === "string") { urlTok = tokens[k]; break; }
        if (tokens[k].type === "template") { urlTok = null; break; } // dynamic url, can't validate
        // Anything else (numbers, idents) means the call is built dynamically.
        if (tokens[k].type === "ident" || tokens[k].type === "number") { urlTok = null; break; }
      }
      if (!urlTok) {
        // Skip past the call.
        i = close;
        continue;
      }
      const url = urlTok.value;

      // Now look BACKWARDS to detect destructuring assignment that wraps this call.
      // Pattern A: `<decl-kw> { ... } = ChromeUtils.importESModule(...)`
      // Pattern B: `({ ... } = ChromeUtils.importESModule(...))` (no decl kw)
      // We scan back over: `=`, then expect `}` (close of destructure pattern).
      let names = null;
      let k = i - 1;
      // Allow `(` directly before for the parenthesised assignment form.
      while (k >= 0 && tokens[k].type === "punct" && tokens[k].value === "(") k--;
      if (k >= 0 && tokens[k].type === "punct" && tokens[k].value === "=") {
        // Walk back to the matching `{`.
        let m = k - 1;
        let depth = 0;
        let braceOpen = -1;
        while (m >= 0) {
          const tm = tokens[m];
          if (tm.type === "punct") {
            if (tm.value === "}") depth++;
            else if (tm.value === "{") {
              depth--;
              if (depth === 0) { braceOpen = m; break; }
            }
          }
          m--;
        }
        if (braceOpen >= 0) {
          const parsed = parseDestructureSourceNames(tokens, braceOpen);
          names = parsed.sourceNames;
        }
      }

      refs.push({ url, names, kind: "dynamic" });
      i = close;
      continue;
    }
  }

  return { refs, exports };
}

// ---------------------------------------------------------------------------
// Driver
// ---------------------------------------------------------------------------

const files = walk(EXT_DIR);
const moduleExports = new Map();    // absPath => Set<name>
const fileImports = new Map();      // absPath => refs[]

for (const f of files) {
  const src = fs.readFileSync(f, "utf8");
  let tokens;
  try {
    tokens = tokenize(src, f);
  } catch (e) {
    console.error(`tokenize failed for ${relExt(f)}: ${e.message}`);
    process.exit(1);
  }
  const { refs, exports } = scan(tokens, f);
  moduleExports.set(f, exports);
  // Only thunderbird-mcp imports participate in graph + per-name validation;
  // Mozilla SDK URLs (resource://gre/, resource:///, chrome://) are skipped.
  fileImports.set(f, refs.filter(r => r.url.startsWith(RES_PREFIX)));
}

const errors = [];

// Existence + named-symbol checks
for (const [importer, refs] of fileImports) {
  for (const ref of refs) {
    const target = resolveResourceUrl(ref.url);
    if (!target) continue;
    if (!fs.existsSync(target)) {
      errors.push(`${relExt(importer)}: import "${ref.url}" -> file not found at ${relExt(target)}`);
      continue;
    }
    if (ref.names && ref.names.length > 0) {
      const exp = moduleExports.get(target);
      if (!exp) continue; // target wasn't walked (shouldn't happen for in-tree files)
      for (const name of ref.names) {
        if (!exp.has(name)) {
          errors.push(`${relExt(importer)}: import { ${name} } from "${ref.url}" -- target does not export "${name}"`);
        }
      }
    }
  }
}

// Cycle detection over thunderbird-mcp imports only
const adj = new Map();
for (const [f, refs] of fileImports) {
  const targets = refs
    .map(r => resolveResourceUrl(r.url))
    .filter(t => t && fs.existsSync(t));
  adj.set(f, targets);
}

const WHITE = 0, GRAY = 1, BLACK = 2;
const color = new Map();
for (const f of files) color.set(f, WHITE);
let cycleCount = 0;

function dfs(node, stack) {
  color.set(node, GRAY);
  for (const next of adj.get(node) || []) {
    if (color.get(next) === GRAY) {
      const cycleStart = stack.indexOf(next);
      const cycle = stack.slice(cycleStart).concat(next);
      errors.push(`Import cycle: ${cycle.map(relExt).join(" -> ")}`);
      cycleCount++;
    } else if (color.get(next) === WHITE) {
      dfs(next, stack.concat(next));
    }
  }
  color.set(node, BLACK);
}
for (const f of files) {
  if (color.get(f) === WHITE) dfs(f, [f]);
}

const totalImports = [...fileImports.values()].reduce((n, r) => n + r.length, 0);

if (errors.length > 0) {
  console.error("Module-graph validation failed:");
  for (const e of errors) console.error("  " + e);
  process.exit(1);
}
console.log(`Module graph OK: ${files.length} files, ${totalImports} thunderbird-mcp imports, cycles=${cycleCount}`);
