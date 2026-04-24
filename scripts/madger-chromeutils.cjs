#!/usr/bin/env node
/**
 * Madger-alike for the Thunderbird MCP codebase.
 *
 * madge / dependency-cruiser / rollup / esbuild all assume either ES
 * static `import` declarations or CJS `require()` calls. Our modules
 * are .sys.mjs loaded via Mozilla's `ChromeUtils.importESModule(...)`
 * with `resource://thunderbird-mcp/...` URLs, and those are invisible
 * to every standard JS graph tool -- running `madge --orphans` on
 * extension/ reports ALL 22 files as orphans (confirmed locally).
 *
 * This script is the parser + analyzer those tools can't be. It
 * reuses the acorn-based ChromeUtils detective that already lives in
 * scripts/check-imports.cjs and emits a proper dependency graph plus
 * a "tool audit" cross-check over the four places a tool has to be
 * wired: schema (lib/tools-schema.sys.mjs), registry (lib/tools/
 * *.sys.mjs), handler bag (api.js toolHandlers), and dispatch switch
 * (lib/dispatch.sys.mjs).
 *
 * Commands:
 *   node scripts/madger-chromeutils.cjs              # summary (default)
 *   node scripts/madger-chromeutils.cjs graph        # full "file -> deps" map
 *   node scripts/madger-chromeutils.cjs orphans      # modules no one imports
 *   node scripts/madger-chromeutils.cjs leaves       # modules that import nothing in our graph
 *   node scripts/madger-chromeutils.cjs circular     # cycles
 *   node scripts/madger-chromeutils.cjs tool-audit   # cross-check tool wiring
 *   node scripts/madger-chromeutils.cjs json         # full graph as JSON
 *
 * Exit code mirrors madge:
 *   0 = no findings (or summary mode always)
 *   1 = findings (orphans / cycles / audit mismatches)
 */

const fs = require('fs');
const path = require('path');
const acorn = require('acorn');

const ROOT = path.resolve(__dirname, '..');
const EXT_DIR = path.join(ROOT, 'extension');
const RES_PREFIX = 'resource://thunderbird-mcp/';

// ── Shared plumbing (mirrors check-imports.cjs deliberately so that
//    file stays the single source of truth for module validation; this
//    file adds the graph + audit views on top) ──────────────────────────

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

function walkAst(node, visit) {
  if (!node || typeof node !== 'object') return;
  if (Array.isArray(node)) { for (const c of node) walkAst(c, visit); return; }
  if (typeof node.type !== 'string') return;
  visit(node);
  for (const key of Object.keys(node)) {
    if (key === 'loc' || key === 'start' || key === 'end' || key === 'range') continue;
    walkAst(node[key], visit);
  }
}

function literalString(n) {
  if (n && n.type === 'Literal' && typeof n.value === 'string') return n.value;
  if (n && n.type === 'TemplateLiteral' && n.quasis.length === 1 && n.expressions.length === 0) {
    return n.quasis[0].value.cooked;
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

/**
 * Collect every import edge a file emits. Returns { imports: [{url, line}],
 * staticImports: [url], dynImports: [url] }. Cross-file kind split is only
 * used for the summary view; the graph itself is kind-agnostic.
 */
function extractImports(ast) {
  const out = [];
  walkAst(ast, (node) => {
    if (node.type === 'ImportDeclaration') {
      out.push({ url: node.source.value, line: node.loc.start.line, kind: 'static' });
    }
    if (isChromeUtilsImportESModule(node)) {
      const url = literalString(node.arguments && node.arguments[0]);
      if (url) out.push({ url, line: node.loc.start.line, kind: 'chromeutils' });
    }
  });
  return out;
}

// ── Build the graph ──────────────────────────────────────────────────

const files = walk(EXT_DIR);
/** file -> { imports: [{target, kind, line}], importers: Set<file> } */
const nodes = new Map();
for (const f of files) nodes.set(f, { imports: [], importers: new Set() });

for (const file of files) {
  const src = fs.readFileSync(file, 'utf8');
  let ast;
  try { ast = parse(file, src); } catch (e) { console.error(e.message); continue; }
  const imports = extractImports(ast);
  for (const imp of imports) {
    if (!imp.url.startsWith(RES_PREFIX)) continue;
    const target = resolveResourceUrl(imp.url);
    if (!target || !nodes.has(target)) continue;
    nodes.get(file).imports.push({ target, kind: imp.kind, line: imp.line });
    nodes.get(target).importers.add(file);
  }
}

// ── Virtual edges for the auto-registry pattern ──────────────────────
//
// tool-registry.sys.mjs calls ChromeUtils.importESModule(manifestPath)
// and ChromeUtils.importESModule(moduleBasePath + relPath) -- neither
// has a literal URL, so the parser can't see the edges. But the runtime
// behavior is deterministic: the manifest lists every module in
// TOOL_MODULES, tool-registry imports each one, and api.js imports
// tool-registry. So we synthesize:
//
//   api.js                  -> lib/tools/_manifest.sys.mjs
//   lib/tool-registry.sys.mjs -> lib/tools/_manifest.sys.mjs
//   lib/tool-registry.sys.mjs -> lib/tools/<each manifest entry>
//
// That turns the three legitimate orphans into reachable nodes and
// mirrors the actual runtime dependency graph.

function addVirtualEdge(from, to, kind) {
  if (!nodes.has(from) || !nodes.has(to)) return;
  nodes.get(from).imports.push({ target: to, kind, line: 0 });
  nodes.get(to).importers.add(from);
}

const manifestFile = path.join(EXT_DIR, 'mcp_server', 'lib', 'tools', '_manifest.sys.mjs');
const toolRegFile = path.join(EXT_DIR, 'mcp_server', 'lib', 'tool-registry.sys.mjs');
const apiFile = path.join(EXT_DIR, 'mcp_server', 'api.js');
if (fs.existsSync(manifestFile)) {
  addVirtualEdge(apiFile, manifestFile, 'registry');
  addVirtualEdge(toolRegFile, manifestFile, 'registry');
  // Parse the manifest to discover which tool files are "imported" at runtime.
  try {
    const ast = parse(manifestFile, fs.readFileSync(manifestFile, 'utf8'));
    walkAst(ast, (node) => {
      if (node.type === 'ExportNamedDeclaration'
          && node.declaration
          && node.declaration.type === 'VariableDeclaration') {
        for (const decl of node.declaration.declarations) {
          if (decl.id.type === 'Identifier' && decl.id.name === 'TOOL_MODULES' && decl.init
              && decl.init.type === 'ArrayExpression') {
            for (const el of decl.init.elements) {
              const fname = literalString(el);
              if (!fname) continue;
              const target = path.join(EXT_DIR, 'mcp_server', 'lib', 'tools', fname);
              addVirtualEdge(toolRegFile, target, 'registry');
            }
          }
        }
      }
    });
  } catch (e) {
    console.error(`Virtual edges: failed to parse manifest (${e.message})`);
  }
}

// ── View helpers ─────────────────────────────────────────────────────

/**
 * Orphan = nothing in our tracked graph imports it AND it's not an
 * entry point. Entry points are known by filename convention:
 *   extension/background.js        -- declared in manifest.json
 *   extension/options.js           -- declared in manifest.json
 *   extension/mcp_server/api.js    -- declared in manifest.json (experiment API)
 * Plus any test harness / CLI helpers which live outside extension/
 * and aren't scanned here.
 */
const ENTRY_POINTS = new Set([
  path.join(EXT_DIR, 'background.js'),
  path.join(EXT_DIR, 'options.js'),
  path.join(EXT_DIR, 'mcp_server', 'api.js'),
  // httpd.sys.mjs is the Mozilla-provided HTTP server shipped alongside
  // our code under extension/. Mozilla's chrome loader pulls it at
  // runtime from a resource:// URL our code constructs dynamically; no
  // static edge exists in our source. Treat as an implicit entry.
  path.join(EXT_DIR, 'httpd.sys.mjs'),
]);

function listOrphans() {
  const orphans = [];
  for (const [file, data] of nodes) {
    if (ENTRY_POINTS.has(file)) continue;
    if (data.importers.size === 0) orphans.push(file);
  }
  return orphans.sort();
}

function listLeaves() {
  const leaves = [];
  for (const [file, data] of nodes) {
    if (data.imports.length === 0) leaves.push(file);
  }
  return leaves.sort();
}

function findCycles() {
  const WHITE = 0, GRAY = 1, BLACK = 2;
  const color = new Map();
  for (const f of nodes.keys()) color.set(f, WHITE);
  const cycles = [];
  function dfs(node, stack) {
    color.set(node, GRAY);
    for (const { target } of nodes.get(node).imports) {
      if (!color.has(target)) continue;
      if (color.get(target) === GRAY) {
        const cycle = stack.slice(stack.indexOf(target)).concat(target);
        cycles.push(cycle);
      } else if (color.get(target) === WHITE) {
        dfs(target, stack.concat(target));
      }
    }
    color.set(node, BLACK);
  }
  for (const f of nodes.keys()) if (color.get(f) === WHITE) dfs(f, [f]);
  return cycles;
}

function graphAsObject() {
  const g = {};
  for (const [file, data] of nodes) {
    g[rel(file)] = data.imports.map(i => rel(i.target));
  }
  return g;
}

// ── Tool audit: the cross-check madger was summoned for ──────────────
//
// Four places a tool has to be wired:
//   1. lib/tools-schema.sys.mjs exports `tools` array (legacy path)
//   2. lib/tools/*.sys.mjs + _manifest.sys.mjs (registry path)
//   3. api.js toolHandlers bag (the function the dispatcher calls)
//   4. lib/dispatch.sys.mjs switch (binds tool name -> toolHandlers.fn(...))
//   5. (bonus) constants.sys.mjs DEFAULT_DISABLED_TOOLS
//
// For each tool we check every place and flag mismatches. These are the
// silent-failure cases: a handler without a schema is uncallable; a
// schema without a handler shows up in tools/list but throws "Unknown
// tool" on tools/call; a switch case that points at a missing
// handlerKey would throw TypeError at runtime.

function extractStringsFromArrayLiteral(node) {
  if (!node || node.type !== 'ArrayExpression') return [];
  return node.elements
    .map(el => literalString(el))
    .filter(v => typeof v === 'string');
}

function toolAudit() {
  const report = {
    schemaTools: new Map(),    // name -> { file, line }
    registryTools: new Map(),  // name -> { file, line }
    handlerBagKeys: new Set(), // strings used as keys in api.js toolHandlers = { ... }
    switchCases: new Map(),    // toolName -> { handlerKey, file, line }
    disabledDefault: new Set(),
    registryManifest: new Set(), // filenames listed in _manifest.sys.mjs
    findings: [],
  };

  // ---- (1) schema tools from tools-schema.sys.mjs
  const schemaFile = path.join(EXT_DIR, 'mcp_server', 'lib', 'tools-schema.sys.mjs');
  if (fs.existsSync(schemaFile)) {
    const ast = parse(schemaFile, fs.readFileSync(schemaFile, 'utf8'));
    walkAst(ast, (node) => {
      if (node.type === 'ObjectExpression') {
        for (const prop of node.properties) {
          if (prop.type === 'Property'
              && !prop.computed
              && prop.key.type === 'Identifier'
              && prop.key.name === 'name'
              && prop.value.type === 'Literal'
              && typeof prop.value.value === 'string') {
            // Heuristic: only count object literals inside the exported `tools`
            // array -- those objects have 'name' + 'inputSchema' together.
            const hasInputSchema = node.properties.some(p =>
              p.type === 'Property' && !p.computed
              && p.key.type === 'Identifier' && p.key.name === 'inputSchema'
            );
            if (hasInputSchema) {
              report.schemaTools.set(prop.value.value, {
                file: rel(schemaFile),
                line: prop.loc.start.line,
              });
            }
          }
        }
      }
    });
  }

  // ---- (2) registry tools from lib/tools/*.sys.mjs (export const tool = { name: ... })
  const toolsDir = path.join(EXT_DIR, 'mcp_server', 'lib', 'tools');
  if (fs.existsSync(toolsDir)) {
    for (const entry of fs.readdirSync(toolsDir)) {
      if (!entry.endsWith('.sys.mjs') || entry === '_manifest.sys.mjs') continue;
      const tf = path.join(toolsDir, entry);
      const ast = parse(tf, fs.readFileSync(tf, 'utf8'));
      walkAst(ast, (node) => {
        if (node.type === 'ExportNamedDeclaration'
            && node.declaration
            && node.declaration.type === 'VariableDeclaration') {
          for (const decl of node.declaration.declarations) {
            if (decl.id.type === 'Identifier' && decl.id.name === 'tool'
                && decl.init && decl.init.type === 'ObjectExpression') {
              const nameProp = decl.init.properties.find(p =>
                p.type === 'Property' && !p.computed
                && p.key.type === 'Identifier' && p.key.name === 'name');
              if (nameProp && literalString(nameProp.value)) {
                report.registryTools.set(nameProp.value.value, {
                  file: rel(tf),
                  line: decl.loc.start.line,
                });
              }
            }
          }
        }
      });
    }
  }

  // ---- registry manifest
  const manifestFile = path.join(toolsDir, '_manifest.sys.mjs');
  if (fs.existsSync(manifestFile)) {
    const ast = parse(manifestFile, fs.readFileSync(manifestFile, 'utf8'));
    walkAst(ast, (node) => {
      if (node.type === 'ExportNamedDeclaration'
          && node.declaration
          && node.declaration.type === 'VariableDeclaration') {
        for (const decl of node.declaration.declarations) {
          if (decl.id.type === 'Identifier' && decl.id.name === 'TOOL_MODULES' && decl.init) {
            for (const f of extractStringsFromArrayLiteral(decl.init)) {
              report.registryManifest.add(f);
            }
          }
        }
      }
    });
  }

  // ---- (3) api.js toolHandlers bag keys
  const apiFile = path.join(EXT_DIR, 'mcp_server', 'api.js');
  if (fs.existsSync(apiFile)) {
    const ast = parse(apiFile, fs.readFileSync(apiFile, 'utf8'));
    walkAst(ast, (node) => {
      if (node.type === 'VariableDeclaration') {
        for (const decl of node.declarations) {
          if (decl.id.type === 'Identifier'
              && decl.id.name === 'toolHandlers'
              && decl.init
              && decl.init.type === 'ObjectExpression') {
            for (const prop of decl.init.properties) {
              if (prop.type === 'Property' && !prop.computed) {
                if (prop.key.type === 'Identifier') {
                  report.handlerBagKeys.add(prop.key.name);
                } else if (literalString(prop.key)) {
                  report.handlerBagKeys.add(prop.key.value);
                }
              }
              // Shorthand property `{ foo }` is Property with shorthand:true
              // -- still an Identifier key so covered above.
            }
          }
        }
      }
    });
  }

  // ---- (4) dispatch.sys.mjs switch cases
  const dispatchFile = path.join(EXT_DIR, 'mcp_server', 'lib', 'dispatch.sys.mjs');
  if (fs.existsSync(dispatchFile)) {
    const ast = parse(dispatchFile, fs.readFileSync(dispatchFile, 'utf8'));
    walkAst(ast, (node) => {
      if (node.type === 'SwitchStatement') {
        for (const kase of node.cases) {
          if (!kase.test) continue; // default
          const toolName = literalString(kase.test);
          if (!toolName) continue;
          // Find the first toolHandlers.X(...) call inside this case
          let handlerKey = null;
          walkAst(kase.consequent, (n) => {
            if (handlerKey) return;
            if (n.type === 'CallExpression'
                && n.callee.type === 'MemberExpression'
                && n.callee.object.type === 'Identifier'
                && n.callee.object.name === 'toolHandlers'
                && n.callee.property.type === 'Identifier') {
              handlerKey = n.callee.property.name;
            }
          });
          report.switchCases.set(toolName, {
            handlerKey,
            file: rel(dispatchFile),
            line: kase.loc.start.line,
          });
        }
      }
    });
  }

  // ---- (5) DEFAULT_DISABLED_TOOLS
  const constantsFile = path.join(EXT_DIR, 'mcp_server', 'lib', 'constants.sys.mjs');
  if (fs.existsSync(constantsFile)) {
    const ast = parse(constantsFile, fs.readFileSync(constantsFile, 'utf8'));
    walkAst(ast, (node) => {
      if (node.type === 'ExportNamedDeclaration'
          && node.declaration
          && node.declaration.type === 'VariableDeclaration') {
        for (const decl of node.declaration.declarations) {
          if (decl.id.type === 'Identifier'
              && decl.id.name === 'DEFAULT_DISABLED_TOOLS'
              && decl.init
              && decl.init.type === 'NewExpression'
              && decl.init.callee.type === 'Identifier'
              && decl.init.callee.name === 'Set') {
            const arg = decl.init.arguments && decl.init.arguments[0];
            for (const s of extractStringsFromArrayLiteral(arg)) {
              report.disabledDefault.add(s);
            }
          }
        }
      }
    });
  }

  // ── Cross-check ──
  const allToolNames = new Set([
    ...report.schemaTools.keys(),
    ...report.registryTools.keys(),
  ]);

  for (const name of allToolNames) {
    const inSchema = report.schemaTools.has(name);
    const inRegistry = report.registryTools.has(name);
    const inSwitch = report.switchCases.has(name);
    // A tool is "callable via the switch" if it has a switch case AND the
    // case resolves to a handlerKey that exists in the toolHandlers bag.
    // Registry tools don't need either -- they go through registryHandlers.
    if (inRegistry) {
      if (!report.registryManifest.has(path.basename(report.registryTools.get(name).file))) {
        report.findings.push({
          severity: 'error',
          tool: name,
          msg: `registry tool ${name} declared in ${report.registryTools.get(name).file} is NOT in _manifest.sys.mjs -- build-xpi will not pick it up; tools/list will not advertise it.`,
        });
      }
      continue;
    }
    if (inSchema && !inSwitch) {
      report.findings.push({
        severity: 'error',
        tool: name,
        msg: `schema tool ${name} (${report.schemaTools.get(name).file}:${report.schemaTools.get(name).line}) has NO switch case in dispatch.sys.mjs -- tools/list would advertise it but tools/call would throw "Unknown tool".`,
      });
      continue;
    }
    if (inSwitch) {
      const { handlerKey, line } = report.switchCases.get(name);
      if (!handlerKey) {
        report.findings.push({
          severity: 'error',
          tool: name,
          msg: `dispatch case for ${name} (dispatch.sys.mjs:${line}) has no toolHandlers.* call -- dead case.`,
        });
      } else if (!report.handlerBagKeys.has(handlerKey)) {
        report.findings.push({
          severity: 'error',
          tool: name,
          msg: `dispatch case for ${name} (dispatch.sys.mjs:${line}) calls toolHandlers.${handlerKey}() but that key is NOT assembled in api.js toolHandlers bag -- runtime TypeError on call.`,
        });
      }
    }
  }

  // Orphan handlerKeys (defined in toolHandlers but no switch case calls them).
  const calledHandlerKeys = new Set();
  for (const sc of report.switchCases.values()) if (sc.handlerKey) calledHandlerKeys.add(sc.handlerKey);
  for (const k of report.handlerBagKeys) {
    if (!calledHandlerKeys.has(k)) {
      // Some keys are called indirectly from other modules, not the switch --
      // collect but mark as warning rather than error.
      report.findings.push({
        severity: 'warning',
        tool: null,
        msg: `toolHandlers.${k} is assembled in api.js but no dispatch.sys.mjs switch case calls it (may be a helper or called indirectly; if it backs a tool the tool is uncallable).`,
      });
    }
  }

  // Orphan switch cases (switch mentions a name not in schema AND not in registry).
  for (const [name, sc] of report.switchCases) {
    if (!allToolNames.has(name)) {
      report.findings.push({
        severity: 'error',
        tool: name,
        msg: `dispatch case for "${name}" (dispatch.sys.mjs:${sc.line}) has no matching schema OR registry entry -- tools/list won't advertise it, so this case is unreachable.`,
      });
    }
  }

  // Disabled-default tools that don't exist at all.
  for (const n of report.disabledDefault) {
    if (!allToolNames.has(n)) {
      report.findings.push({
        severity: 'warning',
        tool: n,
        msg: `DEFAULT_DISABLED_TOOLS mentions "${n}" which is neither in schema nor registry -- stale entry.`,
      });
    }
  }

  return report;
}

// ── Output ───────────────────────────────────────────────────────────

const cmd = process.argv[2] || 'summary';

function fmtSummary() {
  const orphans = listOrphans();
  const leaves = listLeaves();
  const cycles = findCycles();
  console.log(`Graph: ${nodes.size} modules`);
  console.log(`  entry points:     ${[...ENTRY_POINTS].map(rel).join(', ')}`);
  console.log(`  orphans:          ${orphans.length}`);
  console.log(`  leaves:           ${leaves.length}`);
  console.log(`  cycles:           ${cycles.length}`);
  return { orphans, cycles };
}

function runCmd(c) {
  switch (c) {
    case 'summary': {
      const { orphans, cycles } = fmtSummary();
      return orphans.length + cycles.length > 0 ? 0 : 0; // summary never exits non-zero
    }
    case 'orphans': {
      const orphans = listOrphans();
      for (const o of orphans) console.log(rel(o));
      return orphans.length > 0 ? 1 : 0;
    }
    case 'leaves': {
      for (const l of listLeaves()) console.log(rel(l));
      return 0;
    }
    case 'circular': {
      const cycles = findCycles();
      for (const cy of cycles) console.log(cy.map(rel).join(' -> '));
      return cycles.length > 0 ? 1 : 0;
    }
    case 'graph': {
      const g = graphAsObject();
      for (const [k, v] of Object.entries(g).sort()) {
        console.log(k);
        for (const d of v) console.log('  -> ' + d);
      }
      return 0;
    }
    case 'json': {
      console.log(JSON.stringify(graphAsObject(), null, 2));
      return 0;
    }
    case 'tool-audit': {
      const r = toolAudit();
      console.log(`Tool wiring audit:`);
      console.log(`  schema tools:     ${r.schemaTools.size}`);
      console.log(`  registry tools:   ${r.registryTools.size}`);
      console.log(`  handler bag keys: ${r.handlerBagKeys.size}`);
      console.log(`  switch cases:     ${r.switchCases.size}`);
      console.log(`  disabled default: ${r.disabledDefault.size}`);
      console.log(`  registry manifest:${r.registryManifest.size}`);
      console.log('');
      if (r.findings.length === 0) {
        console.log('  no findings -- every tool is wired correctly.');
        return 0;
      }
      const errors = r.findings.filter(f => f.severity === 'error');
      const warnings = r.findings.filter(f => f.severity === 'warning');
      if (errors.length > 0) {
        console.log(`  ERRORS (${errors.length}):`);
        for (const f of errors) console.log(`    - ${f.msg}`);
      }
      if (warnings.length > 0) {
        console.log(`  WARNINGS (${warnings.length}):`);
        for (const f of warnings) console.log(`    - ${f.msg}`);
      }
      return errors.length > 0 ? 1 : 0;
    }
    default:
      console.error(`Unknown command: ${c}`);
      console.error(`Usage: ${path.basename(process.argv[1])} [summary|graph|orphans|leaves|circular|tool-audit|json]`);
      return 2;
  }
}

process.exit(runCmd(cmd));
