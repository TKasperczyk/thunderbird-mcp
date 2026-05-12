// dependency-cruiser config.
//
// Why this exists:
//   This repo has three loaders that share no module graph at runtime:
//     1. Node CJS  — mcp-bridge.cjs, scripts/*.cjs, test/*.cjs
//     2. WebExtension page context — extension/{background,options}.js
//     3. Mozilla chrome module loader — extension/mcp_server/api.js +
//        (post-refactor) extension/mcp_server/lib/*.sys.mjs
//   ESLint catches per-file issues. dependency-cruiser catches the
//   cross-file ones: circular imports, orphans, and accidental
//   coupling between zones that aren't supposed to import each other.
//
// What is NOT here (deliberately):
//   - No rule pinning extension/mcp_server/api.js to "must not import lib/".
//     The api.js→lib/ split lives in PR #2 (refactor/api-modular-split),
//     not this PR. Adding lib/ rules here would create a dead clause until
//     that PR lands. Once it does, a follow-up tightens this config.
//   - No "no-deprecated-core" / "no-unresolved" yet. Those need a tsconfig
//     or jsconfig (we have neither). Holding for after refactor.

module.exports = {
  forbidden: [
    {
      name: "no-circular",
      severity: "error",
      comment:
        "Circular deps survive in CJS because require() returns a partial module on cycle; they break instantly under ESM. This repo runs both — fix at the boundary.",
      from: {},
      to: { circular: true },
    },
    {
      name: "no-orphans",
      severity: "warn",
      comment:
        "Orphans are files no other file imports. Either entry points (listed below in 'allowed') or dead code. knip enforces dead-code rigor; this rule surfaces the dead-code candidates from the graph side.",
      from: {
        orphan: true,
        pathNot: [
          // Entry points — these are LOADED, not imported.
          "^mcp-bridge\\.cjs$",
          "^extension/background\\.js$",
          "^extension/options\\.js$",
          "^extension/mcp_server/api\\.js$",
          // Build / dev scripts run via npm scripts, not import graph.
          "^scripts/",
          // Tests run via node:test runner, not import graph.
          "^test/",
          // Config files.
          "^(eslint\\.config\\.mjs|\\.dependency-cruiser\\.cjs|commitlint\\.config\\.cjs|knip\\.json)$",
          // Vendored XPCOM web server used in tests.
          "^extension/httpd\\.sys\\.mjs$",
        ],
      },
      to: {},
    },
    {
      name: "no-extension-from-bridge",
      severity: "error",
      comment:
        "mcp-bridge.cjs is the Node-side stdio<->WebSocket relay. It must never reach into extension/ — those files run in a completely different sandbox (chrome / page context) and importing them via Node would either fail or produce subtly broken stubs.",
      from: { path: "^mcp-bridge\\.cjs$" },
      to: { path: "^extension/" },
    },
  ],

  options: {
    doNotFollow: { path: "node_modules" },
    exclude: { path: "(^|/)node_modules/" },
    includeOnly: "^(mcp-bridge\\.cjs|extension/|scripts/|test/)",
    moduleSystems: ["cjs", "es6"],
    tsPreCompilationDeps: false,
    progress: { type: "none" },
    reporterOptions: {
      text: { highlightFocused: true },
    },
  },
};
