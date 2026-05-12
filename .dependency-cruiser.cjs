// dependency-cruiser config.
//
// Why this exists:
//   This repo has three loaders that share no module graph at runtime:
//     1. Node CJS  — mcp-bridge.cjs, scripts/*.cjs, test/*.cjs
//     2. WebExtension page context — extension/{background,options}.js
//     3. Mozilla chrome module loader — extension/mcp_server/api.js +
//        extension/mcp_server/lib/*.sys.mjs
//   The lib/ modules are loaded via ChromeUtils.importESModule with
//   resource:// URLs which dep-cruise can't statically resolve. They
//   show up as orphans from the cruise's perspective; we whitelist
//   them in the no-orphans rule below.
//
//   ESLint catches per-file issues. dependency-cruiser catches the
//   cross-file ones: circular imports, orphans, and accidental
//   coupling between zones that aren't supposed to import each other.
//
// What is NOT here (deliberately):
//   - No "no-deprecated-core" / "no-unresolved". Both need a tsconfig
//     or jsconfig (we have neither).

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
          // lib/*.sys.mjs are loaded via ChromeUtils.importESModule with
          // resource:// URLs that dep-cruise can't follow statically.
          // madger-chromeutils.cjs validates the resource:// edges.
          "^extension/mcp_server/lib/",
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
