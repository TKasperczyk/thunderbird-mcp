/**
 * ESLint flat config (industry-standard JavaScript linter, eslint v9+).
 *
 * Two file groups need different treatment:
 *
 *   1. extension/  -- chrome-context code that runs inside Thunderbird.
 *      Has access to XPCOM globals (Cc, Ci, Services, ChromeUtils,
 *      MailServices, dump, ...) and WebExtension globals (browser,
 *      messenger). Imports are mostly via ChromeUtils.importESModule
 *      with resource:// URLs which static linters can't resolve --
 *      that's what scripts/madger-chromeutils.cjs is for. ESLint here
 *      catches the kinds of bugs that DO show up syntactically:
 *      typos, unused vars, unhandled promises, free references.
 *
 *      no-undef is the rule that pays for itself: it would have
 *      caught `isAccountAllowed is not defined` in compose.sys.mjs
 *      at lint time instead of runtime, saving a CI iteration.
 *
 *   2. scripts/, test/, mcp-bridge.cjs -- Node-side code. Standard
 *      Node.js + better-sqlite3 + acorn deps. eslint-plugin-n covers
 *      Node-specific rules (callback shapes, deprecated APIs).
 *
 * Rules deliberately conservative -- this is enabled mid-project on
 * a working codebase, so we want to surface real bugs (no-undef, no
 * floating promises) without drowning in stylistic noise. Prettier-
 * style formatting is intentionally NOT enforced here.
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
      "extension/mcp_server/lib/tools/_manifest.sys.mjs",  // generated
      "extension/httpd.sys.mjs",                            // vendored Mozilla
      ".claude/",
    ],
  },

  // ── Base JS recommendations apply to everything ──────────────────
  js.configs.recommended,

  // ── Extension chrome-context code ────────────────────────────────
  // Includes api.js (experiment API sandbox), background.js (WebExt),
  // options.js (WebExt page), all .sys.mjs modules under
  // extension/mcp_server/lib/, and the embedded httpd.sys.mjs.
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
        // dump() prints to stderr from chrome JS; gated by
        // browser.dom.window.dump.enabled.
        dump: "readonly",
        // ── WebExtension page globals ──
        browser: "readonly",
        messenger: "readonly",
        ExtensionAPI: "readonly",
        ExtensionCommon: "readonly",
        // ── Standard browser-ish globals (background.js + options.js) ──
        ...globals.browser,
      },
    },
    plugins: {
      promise: promisePlugin,
    },
    rules: {
      // ── Real bugs ──
      "no-undef": "error",
      "no-unused-vars": ["warn", {
        argsIgnorePattern: "^_",
        varsIgnorePattern: "^_",
        caughtErrorsIgnorePattern: "^(_|e$|err$|ex$)",
      }],
      "no-redeclare": "error",
      "no-implicit-globals": "off",  // TB module loader has its own scope
      "no-prototype-builtins": "warn",

      // ── Promise hygiene -- the bug class that ate createDrafts ──
      "promise/catch-or-return": ["warn", { allowFinally: true }],
      "promise/no-nesting": "warn",
      "promise/no-return-wrap": "warn",
      "promise/param-names": "warn",
      // always-return is too strict for chrome XPCOM patterns where
      // a Promise is intentionally fire-and-forget into XPCOM glue.
      "promise/always-return": "off",

      // ── Off ──
      // The codebase predates this lint setup; turn off purely
      // stylistic warnings so we surface real issues only.
      "no-empty": ["warn", { allowEmptyCatch: true }],
      "no-control-regex": "off",
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
      "n/no-unsupported-features/node-builtins": "off",  // we run latest Node
      "n/no-process-exit": "off",  // mcp-bridge.cjs uses it intentionally
      "n/no-missing-require": "off",  // better-sqlite3 native dep, plays funny
      "n/no-extraneous-require": "off",
      "promise/catch-or-return": ["warn", { allowFinally: true }],
      "promise/no-nesting": "off",  // common in async iteration helpers
      "no-empty": ["warn", { allowEmptyCatch: true }],
    },
  },

  // ── Special-case: ESM scripts ────────────────────────────────────
  // .mjs files under scripts/ are real ESM, override sourceType
  {
    files: ["scripts/**/*.mjs", "eslint.config.mjs"],
    languageOptions: { sourceType: "module" },
  },
];
