/**
 * ESLint flat config (eslint v9+).
 *
 * Three file groups with different runtime expectations:
 *
 *   1. extension/  -- runs inside Thunderbird's chrome / WebExtension /
 *      experiment-API sandboxes. Has XPCOM globals (Cc, Ci, Services,
 *      ChromeUtils, MailServices, dump). WebExtension page scripts
 *      (background.js, options.js) get the `browser` and `messenger`
 *      globals from the WebExtensions framework.
 *
 *      Imports go via ChromeUtils.importESModule("resource://...") --
 *      not standard ES `import`. ESLint can see static `import` but
 *      not the resource:// edges. That's fine for the rules we
 *      enable: no-undef catches typos / orphan refs, no-unused-vars
 *      flags dead destructures, the promise/* rules catch missing
 *      .catch() on Promise chains. Boundary enforcement happens via
 *      dependency-cruiser (see .dependency-cruiser.cjs).
 *
 *   2. scripts/, test/, mcp-bridge.cjs -- standard Node code. eslint-
 *      plugin-n covers Node-specific patterns.
 *
 *   3. .mjs files under scripts/ are ESM, sourceType:module override.
 *
 * Rules deliberately conservative -- enabling this mid-project means we
 * want signal over noise. Stylistic formatting is intentionally NOT
 * enforced (prettier-style); the goal is bug detection, not whitespace.
 *
 * Notably useful in this codebase:
 *   - no-undef:    catches free references in api.js's nested closures
 *                  (anything from the old monolithic structure where a
 *                  helper was defined in one IIFE and called from a
 *                  sibling IIFE that doesn't close over it).
 *   - promise/catch-or-return: catches .then() chains without .catch().
 *                  In the experiment-API surface this is the difference
 *                  between a clear { error } response and a 60s
 *                  tools/call timeout that hangs the client.
 */

import js from "@eslint/js";
import nodePlugin from "eslint-plugin-n";
import promisePlugin from "eslint-plugin-promise";
import globals from "globals";

export default [
  // ── Global ignores ────────────────────────────────────────────────
  {
    ignores: [
      "node_modules/",
      "dist/",
      "extension/buildinfo.json",
      "extension/httpd.sys.mjs",  // vendored Mozilla httpd, not ours
      ".claude/",
    ],
  },

  // ── Base JS recommendations apply to everything ──────────────────
  js.configs.recommended,

  // ── Extension chrome-context code ────────────────────────────────
  {
    files: ["extension/**/*.{js,mjs}"],
    languageOptions: {
      ecmaVersion: "latest",
      sourceType: "module",
      globals: {
        // ── XPCOM / chrome ──
        Components: "readonly",
        ChromeUtils: "readonly",
        Services: "readonly",
        Cc: "readonly", Ci: "readonly", Cr: "readonly", Cu: "readonly",
        MailServices: "readonly",
        dump: "readonly",
        // ── WebExtension globals ──
        browser: "readonly",
        messenger: "readonly",
        ExtensionAPI: "readonly",
        ExtensionCommon: "readonly",
        // ── Standard browser-ish (background.js + options.js) ──
        ...globals.browser,
      },
    },
    plugins: {
      promise: promisePlugin,
    },
    rules: {
      "no-undef": "error",
      "no-unused-vars": ["warn", {
        argsIgnorePattern: "^_",
        varsIgnorePattern: "^_",
        caughtErrorsIgnorePattern: "^(_|e$|err$|ex$)",
      }],
      "no-redeclare": "error",
      "no-implicit-globals": "off",  // chrome module loader has its own scope
      "no-prototype-builtins": "warn",

      // Promise hygiene -- the bug class that hides 30s tools/call hangs.
      "promise/catch-or-return": ["warn", { allowFinally: true }],
      "promise/no-nesting": "warn",
      "promise/no-return-wrap": "warn",
      "promise/param-names": "warn",
      // always-return is too strict for chrome XPCOM fire-and-forget patterns.
      "promise/always-return": "off",

      "no-empty": ["warn", { allowEmptyCatch: true }],
      "no-control-regex": "off",  // we filter JSON-RPC control chars intentionally
    },
  },

  // ── Node-side scripts and tests ──────────────────────────────────
  {
    files: ["scripts/**/*.{cjs,mjs,js}", "test/**/*.{cjs,mjs,js}", "mcp-bridge.cjs"],
    languageOptions: {
      ecmaVersion: "latest",
      sourceType: "commonjs",
      globals: {
        ...globals.node,
        ...globals.es2024,
      },
    },
    plugins: {
      n: nodePlugin,
      promise: promisePlugin,
    },
    rules: {
      "no-unused-vars": ["warn", {
        argsIgnorePattern: "^_",
        varsIgnorePattern: "^_",
        caughtErrorsIgnorePattern: "^(_|e$|err$|ex$)",
      }],
      "n/no-unsupported-features/node-builtins": "off",
      "n/no-process-exit": "off",          // mcp-bridge.cjs uses it intentionally
      "n/no-missing-require": "off",        // native deps confuse the resolver
      "n/no-extraneous-require": "off",
      "promise/catch-or-return": ["warn", { allowFinally: true }],
      "promise/no-nesting": "off",
      "no-empty": ["warn", { allowEmptyCatch: true }],
    },
  },

  // ── Scripts that are real ESM (rare, but supported) ──────────────
  {
    files: ["scripts/**/*.mjs", "eslint.config.mjs"],
    languageOptions: { sourceType: "module" },
  },

  // ── Transitional: api.js is being split into lib/*.sys.mjs ───────
  //
  // The monolith has at least one ghost-reference bug surfaced by
  // no-undef (`removeConnectionInfo` at line ~5982, defined inside a
  // sibling closure at ~999). It's a real bug, but the fix lives in
  // PR #2 (refactor/api-modular-split) where the code moves into
  // lib/connection.sys.mjs and the scope problem disappears.
  //
  // Until that PR lands, downgrade no-undef from error to warn ONLY
  // for api.js so it doesn't block CI on every subsequent PR. After
  // PR #2 merges, delete this whole block (api.js itself is gone).
  //
  // Promise rules are already warn-level globally, so they need no
  // override.
  {
    files: ["extension/mcp_server/api.js"],
    rules: {
      "no-undef": "warn",
    },
  },
];
