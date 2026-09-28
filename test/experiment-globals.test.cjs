"use strict";

const { describe, it } = require("node:test");
const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const vm = require("node:vm");

const extensionDir = path.resolve(__dirname, "../extension");
const manifest = JSON.parse(fs.readFileSync(path.join(extensionDir, "manifest.json"), "utf8"));
const source = fs.readFileSync(path.join(extensionDir, "mcp_server/api.js"), "utf8");
const START = "// BEGIN EXPERIMENT GLOBAL IMPORTS";
const END = "// END EXPERIMENT GLOBAL IMPORTS";

// Web/Node globals do not come with the Experiment sandbox. JS builtins and
// Thunderbird-provided globals (Cc, Ci, Cu, ChromeUtils, Services, console, etc.)
// are deliberately absent. Keep the Node names here even though Cu cannot import them.
const HOST_GLOBALS = new Set([
  "DOMParser", "atob", "btoa", "TextDecoder", "TextEncoder", "URL", "URLSearchParams",
  "crypto", "fetch", "Blob", "File", "FileReader", "FormData", "Headers", "Request", "Response",
  "AbortController", "AbortSignal", "XMLHttpRequest", "WebSocket", "DOMException",
  "setTimeout", "setInterval", "clearTimeout", "clearInterval", "setImmediate", "clearImmediate",
  "queueMicrotask", "structuredClone", "Buffer", "process", "require", "module", "exports",
  "__dirname", "__filename", "window", "document", "self", "navigator", "performance",
  "Event", "CustomEvent", "EventTarget", "Node", "NodeFilter", "XMLSerializer",
  "ReadableStream", "WritableStream", "TransformStream", "CompressionStream", "DecompressionStream",
  "MessageChannel", "MessagePort", "BroadcastChannel", "Image", "ImageData", "MutationObserver",
  "indexedDB", "caches", "localStorage", "sessionStorage", "requestAnimationFrame", "cancelAnimationFrame",
]);

// Reviewed against xpc::GlobalProperties::Parse (not the much larger browser
// global list): https://searchfox.org/mozilla-central/source/js/xpconnect/src/Sandbox.cpp
// Only allow reviewed names; expanding this list requires checking Gecko support.
const IMPORTABLE_GLOBALS = new Set([
  "AbortController", "Blob", "DOMException", "DOMParser", "Event", "File", "FileReader",
  "FormData", "Headers", "MessageChannel", "Node", "NodeFilter", "ReadableStream",
  "TextDecoder", "TextEncoder", "URL", "URLSearchParams", "WebSocket", "XMLHttpRequest",
  "XMLSerializer", "atob", "btoa", "caches", "crypto", "fetch", "indexedDB", "structuredClone",
]);

// A small lexer, not a regexp over raw source: quoted strings, comments, regexp
// bodies and template text are inert, while ${...} expressions remain code.
function tokensFor(code) {
  const tokens = [];
  let offset = 0;
  function read(untilBrace = false) {
    let depth = 0;
    let expressionStart = true;
    const controlParens = [];
    while (offset < code.length) {
      const start = offset;
      const char = code[offset];
      if (/\s/.test(char)) { offset++; continue; }
      if (code.startsWith("//", offset)) {
        while (offset < code.length && code[offset] !== "\n") offset++;
        continue;
      }
      if (code.startsWith("/*", offset)) {
        const end = code.indexOf("*/", offset + 2);
        offset = end < 0 ? code.length : end + 2;
        continue;
      }
      if (char === "\"" || char === "'") {
        offset++;
        while (offset < code.length) {
          if (code[offset++] === "\\") offset++;
          else if (code[offset - 1] === char) break;
        }
        tokens.push({ type: "string", value: code.slice(start, offset), start });
        expressionStart = false;
        continue;
      }
      if (char === "`") {
        offset++;
        while (offset < code.length) {
          if (code[offset] === "\\") { offset += 2; continue; }
          if (code[offset++] === "`") break;
          if (code[offset - 1] === "$" && code[offset] === "{") {
            offset++;
            read(true);
          }
        }
        expressionStart = false;
        continue;
      }
      if (char === "/" && expressionStart) {
        offset++;
        let inClass = false;
        while (offset < code.length) {
          const next = code[offset++];
          if (next === "\\") { offset++; continue; }
          if (next === "[") inClass = true;
          if (next === "]") inClass = false;
          if (next === "/" && !inClass) break;
        }
        while (/[a-z]/i.test(code[offset] || "")) offset++;
        expressionStart = false;
        continue;
      }
      if (char === "}" && untilBrace && depth === 0) { offset++; return; }
      const identifier = /^[A-Za-z_$][\w$]*/.exec(code.slice(offset));
      if (identifier) {
        offset += identifier[0].length;
        tokens.push({ type: "identifier", value: identifier[0], start });
        expressionStart = /^(return|throw|case|delete|void|typeof|yield|await|new|in|of)$/.test(identifier[0]);
        continue;
      }
      const value = code.startsWith("=>", offset) || code.startsWith("?.", offset)
        ? code.slice(offset, offset + 2) : char;
      offset += value.length;
      if (char === "{") depth++;
      if (char === "}") depth--;
      if (char === "(") controlParens.push(/^(if|while|for|with|switch|catch)$/.test(tokens[tokens.length - 1]?.value || ""));
      tokens.push({ type: "punctuation", value, start });
      expressionStart = char === ")" ? controlParens.pop() === true : !/[\w\])}]/.test(value);
    }
  }
  read();
  return tokens;
}

function referencedGlobals(code) {
  const tokens = tokensFor(code);
  const pairs = new Map();
  const containers = [];
  const stack = [];
  for (let i = 0; i < tokens.length; i++) {
    containers[i] = stack[stack.length - 1];
    if (["(", "[", "{"].includes(tokens[i].value)) stack.push(i);
    if ([")", "]", "}"].includes(tokens[i].value)) {
      const start = stack.pop();
      pairs.set(start, i);
      pairs.set(i, start);
    }
  }
  const scopes = [{ start: -1, end: tokens.length, bindings: new Set() }];
  for (const [start, end] of pairs) {
    if (tokens[start]?.value === "{" && end > start) scopes.push({ start, end, bindings: new Set() });
  }
  function scopeAt(index) {
    return scopes.filter(scope => scope.start < index && scope.end > index)
      .sort((a, b) => b.start - a.start)[0];
  }
  const bindingTokens = new Set();
  function bind(index, scope) {
    if (tokens[index]?.type !== "identifier") return;
    bindingTokens.add(index);
    scope.bindings.add(tokens[index].value);
  }
  function bindPattern(start, end, scope) {
    for (let i = start; i < end; i++) {
      if (tokens[i].value === "=") {
        while (i + 1 < end && tokens[i + 1].value !== ",") {
          i++;
          if (pairs.get(i) > i) i = pairs.get(i);
        }
      } else if (tokens[i + 1]?.value !== ":") {
        bind(i, scope);
      }
    }
  }
  function isDeclaration(index) {
    if (tokens[index - 1]?.value === "async") index--;
    return index === 0 || ["{", "}", ";", "export", "default"].includes(tokens[index - 1]?.value);
  }
  for (let i = 0; i < tokens.length; i++) {
    const value = tokens[i].value;
    if (["const", "let", "var"].includes(value)) {
      const end = pairs.get(i + 1);
      if (end > i + 1) bindPattern(i + 2, end, scopeAt(i));
      else bind(i + 1, scopeAt(i));
    }
    if (value === "function" || value === "catch") {
      let params = i + 1;
      if (tokens[params]?.value === "*") params++;
      const name = tokens[params]?.type === "identifier" ? params++ : null;
      const close = pairs.get(params);
      const body = scopes.find(scope => scope.start === close + 1);
      if (body) {
        if (name !== null) bind(name, isDeclaration(i) ? scopeAt(i) : body);
        bindPattern(params + 1, close, body);
      }
    }
    if (value === "class" && tokens[i + 1]?.type === "identifier" && tokens[i + 1].value !== "extends") {
      let bodyStart = i + 2;
      while (bodyStart < tokens.length && tokens[bodyStart].value !== "{") {
        bodyStart = (pairs.get(bodyStart) > bodyStart ? pairs.get(bodyStart) : bodyStart) + 1;
      }
      const body = scopes.find(scope => scope.start === bodyStart);
      if (body) bind(i + 1, isDeclaration(i) ? scopeAt(i) : body);
    }
    if (value === "=>") {
      const start = tokens[i - 1]?.value === ")" ? pairs.get(i - 1) : i - 1;
      const body = scopes.find(scope => scope.start === i + 1);
      if (body) bindPattern(start + (tokens[start]?.value === "(" ? 1 : 0), i - (tokens[i - 1]?.value === ")" ? 1 : 0), body);
    }
  }
  function shadowed(name, index) {
    return scopes.some(scope => scope.start < index && scope.end > index && scope.bindings.has(name));
  }
  function isPropertyKey(index) {
    if (tokens[index + 1]?.value !== ":" || !["{", ","].includes(tokens[index - 1]?.value)) return false;
    const container = containers[index];
    if (tokens[container]?.value !== "{") return false;
    // A comma in a ternary's middle expression does not begin an object key.
    let conditionals = 0;
    for (let i = container + 1; i < index; i++) {
      if (containers[i] !== container) continue;
      if (tokens[i].value === "?" && tokens[i - 1]?.value !== "?" && tokens[i + 1]?.value !== "?") conditionals++;
      if (tokens[i].value === ":" && conditionals > 0) conditionals--;
    }
    return conditionals === 0;
  }
  const used = new Set();
  for (let i = 0; i < tokens.length; i++) {
    const token = tokens[i];
    if (token.value === "globalThis" && !shadowed("globalThis", i)) {
      let member;
      if ([".", "?."].includes(tokens[i + 1]?.value)) member = tokens[i + 2]?.value;
      if (tokens[i + 1]?.value === "[" && tokens[i + 2]?.type === "string") {
        member = vm.runInNewContext(tokens[i + 2].value, Object.create(null), { timeout: 100 });
      }
      if (HOST_GLOBALS.has(member)) used.add(member);
    }
    if (token.type !== "identifier" || !HOST_GLOBALS.has(token.value) || bindingTokens.has(i)) continue;
    if ([".", "?."].includes(tokens[i - 1]?.value) || isPropertyKey(i)) continue;
    if (!shadowed(token.value, i)) used.add(token.value);
  }
  return used;
}

function importedGlobals(code) {
  const start = code.indexOf(START);
  const end = code.indexOf(END, start);
  assert.ok(start >= 0 && end > start, "Missing production Experiment import markers");
  assert.equal(code.indexOf(START, start + START.length), -1, "Duplicate import marker");
  const imports = [];
  let calls = 0;
  vm.runInNewContext(code.slice(start, end), {
    Cu: { importGlobalProperties(names) { calls++; imports.push(...names); } },
  }, { timeout: 1000 });
  assert.equal(calls, 1, "Import web globals exactly once");
  assert.equal(new Set(imports).size, imports.length, "Duplicate imported global");
  const tokens = tokensFor(code);
  const importCalls = tokens.filter((token, index) => token.value === "Cu" &&
    tokens[index + 1]?.value === "." && tokens[index + 2]?.value === "importGlobalProperties");
  assert.equal(importCalls.length, 1, "Keep the sole import inside its production marker");
  // No executable declaration or statement may precede the import besides strict mode.
  assert.deepEqual(tokensFor(code.slice(0, start)).map(token => token.value), ['"use strict"', ";"]);
  return new Set(imports);
}

function auditGlobals(code) {
  const imports = importedGlobals(code);
  const used = referencedGlobals(code);
  return {
    missing: [...used].filter(name => !imports.has(name)).sort(),
    unused: [...imports].filter(name => !used.has(name)).sort(),
    invalid: [...imports].filter(name => !IMPORTABLE_GLOBALS.has(name)).sort(),
  };
}

function changeImports(code, names) {
  return code.slice(0, code.indexOf(START)) + START + "\nCu.importGlobalProperties(" +
    JSON.stringify(names) + ");\n" + code.slice(code.indexOf(END));
}

describe("Experiment parent globals", () => {
  it("imports every used web global, only used globals, from Gecko's supported list", () => {
    for (const experiment of Object.values(manifest.experiment_apis)) {
      if (!experiment.parent) continue;
      const code = fs.readFileSync(path.join(extensionDir, experiment.parent.script), "utf8");
      assert.deepEqual(auditGlobals(code), { missing: [], unused: [], invalid: [] }, experiment.parent.script);
    }
  });

  it("rejects missing imports for every audited web/Node global, including typeof probes", () => {
    const code = changeImports(source, []);
    const bare = auditGlobals(code + [...HOST_GLOBALS].map(name => `\nvoid typeof ${name};`).join(""));
    const members = auditGlobals(code + [...HOST_GLOBALS].map(name => `\nvoid globalThis.${name};`).join(""));
    assert.deepEqual(bare.missing, [...HOST_GLOBALS].sort());
    assert.deepEqual(members.missing, [...HOST_GLOBALS].sort());
  });

  it("rejects unused and unsupported imports", () => {
    const imports = [...importedGlobals(source)];
    assert.ok(auditGlobals(changeImports(source, [...imports, "URL"])).unused.includes("URL"));
    assert.ok(auditGlobals(changeImports(source, [...imports, "setTimeout"])).invalid.includes("setTimeout"));
    assert.ok(auditGlobals(changeImports(source, [...imports, "Buffer"])).invalid.includes("Buffer"));
  });

  it("ignores inert text and object members, but reads template expressions and computed global access", () => {
    assert.deepEqual([...referencedGlobals([
      "// DOMParser fetch\n/* btoa TextEncoder */",
      "const literal = 'URL TextDecoder'; const re = /fetch[\\/]DOMParser/gi;",
      "if (ready) /Blob|FileReader/.test(literal);",
      "const template = `URL ${'fetch'} ${`Blob ${crypto}`}`;",
      "object.fetch(); object?.DOMParser; const value = { URL: 'value' };",
      "globalThis['TextEncoder'];",
    ].join("\n"))].sort(), ["TextEncoder", "crypto"]);
  });

  it("distinguishes local declarations and parameters from free global references", () => {
    assert.deepEqual([...referencedGlobals([
      "function sample(URL, { fetch: localFetch }) { return new URL(localFetch()); }",
      "const local = (TextEncoder) => { return new TextEncoder(); };",
      "{ const Blob = class {}; new Blob(); }",
      "const { setTimeout } = ChromeUtils.importESModule('resource://gre/modules/Timer.sys.mjs');",
      "setTimeout(() => {}, 1);",
      "new DOMParser();",
    ].join("\n"))], ["DOMParser"]);
  });

  it("keeps ternary values and switch cases distinct from object property keys", () => {
    for (const statement of [
      "enabled ? TextEncoder : null;",
      "switch (value) { case TextEncoder: break; }",
      "const options = { value: enabled ? first, TextEncoder : null };",
    ]) {
      assert.ok(auditGlobals(`${source}\n${statement}`).missing.includes("TextEncoder"), statement);
    }
    assert.deepEqual([...referencedGlobals("const options = { TextEncoder: null, URL: enabled ?? null, Blob: null };" )], []);
  });

  it("limits named function and class expression bindings to the expression itself", () => {
    for (const expression of [
      "function TextEncoder() { return new TextEncoder(); }",
      "class TextEncoder { static create() { return new TextEncoder(); } }",
    ]) {
      const declaration = `const local = ${expression};`;
      assert.deepEqual([...referencedGlobals(declaration)], [], declaration);
      assert.ok(auditGlobals(`${source}\n${declaration}\nnew TextEncoder();`).missing.includes("TextEncoder"), declaration);
    }
    assert.deepEqual([...referencedGlobals([
      "function TextEncoder() {} new TextEncoder();",
      "class Blob {} new Blob();",
    ].join("\n"))], []);
    assert.deepEqual([...referencedGlobals("function local() { function URL() {} new URL(); } new URL();")], ["URL"]);
  });
});
