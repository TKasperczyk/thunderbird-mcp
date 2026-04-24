#!/usr/bin/env node
/**
 * Static module-graph validator for the Thunderbird MCP extension.
 *
 * Walks every .js / .mjs / .sys.mjs / .cjs file under extension/, parses
 * each with acorn, and extracts:
 *
 *   - static `import { foo } from "..."` declarations (incl. default,
 *     namespace, side-effect-only, named, mixed)
 *   - dynamic `ChromeUtils.importESModule("...")` calls, including the
 *     destructuring pattern `const { foo } = ChromeUtils.importESModule(...)`
 *   - `export { ... }` and `export const|let|var|function|class` declarations
 *
 * For each import targeting `resource://thunderbird-mcp/...` it then:
 *   - resolves to extension/<rest> (the substitution api.js installs)
 *   - asserts the file exists
 *   - asserts each named symbol the importer pulls is actually exported
 *   - participates in cycle detection over the thunderbird-mcp subgraph
 *
 * Mozilla SDK URLs (resource://gre/, resource:///, chrome://) are intentionally
 * NOT validated -- those exist in Thunderbird at runtime, not in this repo.
 *
 * Exits 0 on clean, 1 on any finding. Wired into the build via
 * `node scripts/check-imports.cjs` -- safe to add as a pre-commit hook.
 */

const fs = require('fs');
const path = require('path');
const acorn = require('acorn');

const ROOT = path.resolve(__dirname, '..');
const EXT_DIR = path.join(ROOT, 'extension');
const RES_PREFIX = 'resource://thunderbird-mcp/';

function walk(dir, out = []) {
  for (const entry of fs.readdirSync(dir, { withFileTypes: true })) {
    const full = path.join(dir, entry.name);
    if (entry.isDirectory()) walk(full, out);
    else if (/\.(?:js|mjs|cjs)$/.test(entry.name)) out.push(full);
  }
  return out;
}

function rel(p) { return path.relative(ROOT, p).replace(/\\/g, '/'); }

function resolveResourceUrl(url) {
  if (!url.startsWith(RES_PREFIX)) return null;
  const tail = url.slice(RES_PREFIX.length).split('?')[0];
  return path.join(EXT_DIR, tail);
}

function parse(file, src) {
  // Try module first (ESM imports/exports), fall back to script for legacy CJS.
  for (const sourceType of ['module', 'script']) {
    try {
      return acorn.parse(src, {
        ecmaVersion: 'latest',
        sourceType,
        allowReturnOutsideFunction: sourceType === 'script',
        allowAwaitOutsideFunction: true,
        allowImportExportEverywhere: true,
        locations: true,
      });
    } catch (e) {
      if (sourceType === 'script') {
        throw new Error(`${rel(file)}: parse failed (${e.message})`);
      }
    }
  }
  return null;
}

function isChromeUtilsImportESModule(node) {
  return node.type === 'CallExpression'
    && node.callee.type === 'MemberExpression'
    && node.callee.object.type === 'Identifier'
    && node.callee.object.name === 'ChromeUtils'
    && node.callee.property.type === 'Identifier'
    && node.callee.property.name === 'importESModule';
}

function literalString(node) {
  if (node && node.type === 'Literal' && typeof node.value === 'string') return node.value;
  if (node && node.type === 'TemplateLiteral' && node.quasis.length === 1 && node.expressions.length === 0) {
    return node.quasis[0].value.cooked;
  }
  return null;
}

function namesFromObjectPattern(pat) {
  // const { foo, bar: baz, ...rest } = ...   -- collect imported symbol names
  const names = new Set();
  if (!pat || pat.type !== 'ObjectPattern') return names;
  for (const prop of pat.properties) {
    if (prop.type === 'Property' && prop.key.type === 'Identifier' && !prop.computed) {
      names.add(prop.key.name);
    }
  }
  return names;
}

// Recursive walker -- acorn's AST is small and stable.
function walkAst(node, visit) {
  if (!node || typeof node !== 'object') return;
  if (Array.isArray(node)) {
    for (const child of node) walkAst(child, visit);
    return;
  }
  if (typeof node.type !== 'string') return;
  visit(node);
  for (const key of Object.keys(node)) {
    if (key === 'loc' || key === 'start' || key === 'end' || key === 'range') continue;
    walkAst(node[key], visit);
  }
}

function extract(file, ast) {
  const exports = new Set();
  const imports = []; // { url, names: Set|null, line }
  // CallExpression nodes that are the RHS of a `const {...} = ChromeUtils.importESModule(...)`
  // pattern. We collect names from the destructure and skip the bare CallExpression
  // visit so each call yields exactly one import entry.
  const destructuredCalls = new WeakSet();

  // First pass: find destructured ChromeUtils.importESModule calls and record
  // them with their imported names. Marks the call node so the second pass
  // doesn't double-count it.
  walkAst(ast, (node) => {
    if (node.type === 'VariableDeclaration') {
      for (const decl of node.declarations) {
        if (decl.init && isChromeUtilsImportESModule(decl.init)) {
          const url = literalString(decl.init.arguments && decl.init.arguments[0]);
          if (url) {
            const names = namesFromObjectPattern(decl.id);
            imports.push({
              url,
              names: names.size > 0 ? names : null,
              line: decl.loc.start.line,
            });
            destructuredCalls.add(decl.init);
          }
        }
      }
    }
    // Same idea for `({ foo } = ChromeUtils.importESModule(...))` AssignmentExpressions
    if (node.type === 'AssignmentExpression'
        && node.left.type === 'ObjectPattern'
        && isChromeUtilsImportESModule(node.right)) {
      const url = literalString(node.right.arguments && node.right.arguments[0]);
      if (url) {
        const names = namesFromObjectPattern(node.left);
        imports.push({
          url,
          names: names.size > 0 ? names : null,
          line: node.loc.start.line,
        });
        destructuredCalls.add(node.right);
      }
    }
  });

  // Second pass: everything else.
  walkAst(ast, (node) => {
    // ---------- static imports ----------
    if (node.type === 'ImportDeclaration') {
      const url = node.source.value;
      const names = new Set();
      let namespaceOnly = false;
      for (const spec of node.specifiers) {
        if (spec.type === 'ImportSpecifier') {
          names.add(spec.imported.name);
        } else if (spec.type === 'ImportDefaultSpecifier') {
          names.add('default');
        } else if (spec.type === 'ImportNamespaceSpecifier') {
          // import * as X -- can't validate individual names
          namespaceOnly = true;
        }
      }
      imports.push({
        url,
        names: namespaceOnly ? null : (names.size > 0 ? names : null),
        line: node.loc.start.line,
      });
    }

    // ---------- bare ChromeUtils.importESModule(...) (not in a destructure) ----------
    if (isChromeUtilsImportESModule(node) && !destructuredCalls.has(node)) {
      const url = literalString(node.arguments && node.arguments[0]);
      if (url) imports.push({ url, names: null, line: node.loc.start.line });
    }

    // ---------- exports ----------
    if (node.type === 'ExportNamedDeclaration') {
      if (node.declaration) {
        const d = node.declaration;
        if (d.type === 'VariableDeclaration') {
          for (const decl of d.declarations) {
            if (decl.id.type === 'Identifier') exports.add(decl.id.name);
          }
        } else if ((d.type === 'FunctionDeclaration' || d.type === 'ClassDeclaration') && d.id) {
          exports.add(d.id.name);
        }
      }
      for (const spec of node.specifiers || []) {
        exports.add(spec.exported.name);
      }
    }
    if (node.type === 'ExportDefaultDeclaration') {
      exports.add('default');
    }
  });

  return { exports, imports };
}

// ---------- run ----------

const files = walk(EXT_DIR);
const moduleData = new Map(); // file -> { exports, imports }
const errors = [];

for (const file of files) {
  try {
    const src = fs.readFileSync(file, 'utf8');
    const ast = parse(file, src);
    moduleData.set(file, extract(file, ast));
  } catch (e) {
    errors.push(`${rel(file)}: ${e.message}`);
  }
}

// Existence + named-symbol checks
for (const [importer, { imports }] of moduleData) {
  for (const ref of imports) {
    if (!ref.url.startsWith(RES_PREFIX)) continue;
    const target = resolveResourceUrl(ref.url);
    if (!target) continue;
    if (!fs.existsSync(target)) {
      errors.push(`${rel(importer)}:${ref.line}: import "${ref.url}" -> file not found at ${rel(target)}`);
      continue;
    }
    const targetData = moduleData.get(target);
    if (!targetData || !ref.names) continue;
    for (const name of ref.names) {
      if (!targetData.exports.has(name)) {
        errors.push(`${rel(importer)}:${ref.line}: import { ${name} } from "${ref.url}" -- target does not export "${name}" (target exports: ${[...targetData.exports].sort().join(', ') || '<nothing>'})`);
      }
    }
  }
}

// Cycle detection over thunderbird-mcp imports only
const adj = new Map();
for (const [file, { imports }] of moduleData) {
  const targets = imports
    .map(r => resolveResourceUrl(r.url))
    .filter(t => t && fs.existsSync(t));
  adj.set(file, [...new Set(targets)]);
}

const WHITE = 0, GRAY = 1, BLACK = 2;
const color = new Map();
for (const f of files) color.set(f, WHITE);
let cycles = 0;

function dfs(node, stack) {
  color.set(node, GRAY);
  for (const next of adj.get(node) || []) {
    if (!color.has(next)) continue;
    if (color.get(next) === GRAY) {
      const cycle = stack.slice(stack.indexOf(next)).concat(next);
      errors.push(`Import cycle: ${cycle.map(rel).join(' -> ')}`);
      cycles++;
    } else if (color.get(next) === WHITE) {
      dfs(next, stack.concat(next));
    }
  }
  color.set(node, BLACK);
}
for (const f of files) if (color.get(f) === WHITE) dfs(f, [f]);

if (errors.length > 0) {
  console.error('Module-graph validation failed:');
  for (const e of errors) console.error('  ' + e);
  process.exit(1);
}

const totalImports = [...moduleData.values()].reduce(
  (n, { imports }) => n + imports.filter(r => r.url.startsWith(RES_PREFIX)).length,
  0
);
console.log(`Module graph OK: ${files.length} files, ${totalImports} thunderbird-mcp imports, cycles=${cycles}.`);
