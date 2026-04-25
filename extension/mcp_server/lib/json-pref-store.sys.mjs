/**
 * Reusable JSON-document-in-a-pref storage pattern for WebExtension
 * experiment APIs.
 *
 * There are two Mozilla-specific gotchas that this module codifies so we
 * don't reinvent the hack every time we want to expose a structured
 * policy / config document through the experiment API:
 *
 *  1. `additionalProperties: true` is NOT honored in schema.json parameter
 *     declarations. The Mozilla Schemas validator pre-rejects any nested
 *     object that doesn't exactly match the declared property shape, and
 *     that rejection surfaces to the options-page caller as the unhelpful
 *     generic "An unexpected error occurred" with no diagnostic.
 *
 *     WORKAROUND: declare the parameter as { "type": "any" } in the
 *     experiment-API schema.json. Client-side validation is this module's
 *     job, NOT the schema validator's.
 *
 *  2. Errors thrown from an experiment-API handler body also surface as
 *     "An unexpected error occurred" with no diagnostic. Every handler
 *     must therefore try/catch and return { error: <message> } instead
 *     of throwing.
 *
 *     WORKAROUND: use wrapApi() from this module to wrap the handler
 *     body. Returns { error } on any throw, with a prefixed message that
 *     names the failing method so option-page errors are traceable.
 *
 * The factory also owns the three-path load / save / clear lifecycle that
 * permissions.sys.mjs had inline, so new stores (future tool-access,
 * account-scope, whatever) don't have to reimplement prefHasUserValue
 * fallback, corrupt-pref handling, or user-choice-vs-default semantics.
 *
 * Typical usage in a lib/*.sys.mjs factory:
 *
 *   const store = makeJsonPrefStore({
 *     Services,
 *     prefKey: PREF_FOO,
 *     validate: normalizeFooPolicy,
 *     defaultValue: { version: 1, rules: {} },
 *   });
 *   // store.load() / store.save(v) / store.clear()
 *
 * And in api.js (experiment-API surface):
 *
 *   getFooConfig: async function() {
 *     return wrapApi("getFooConfig", () => ({ policy: fooStore.load() }));
 *   },
 *   setFooConfig: async function(policy) {
 *     return wrapApi("setFooConfig", () => ({ success: true, policy: fooStore.save(policy) }));
 *   },
 *   clearFooConfig: async function() {
 *     return wrapApi("clearFooConfig", () => { fooStore.clear(); return { success: true }; });
 *   },
 *
 * with the matching schema.json parameter declared as { "type": "any" }.
 */

/**
 * Build a load/save/clear trio over a single string-pref holding a JSON
 * document.
 *
 * - load():        returns the parsed document, or a fresh deep-cloned
 *                  defaultValue when the pref has no user value OR is
 *                  corrupt. Never throws.
 * - save(value):   runs the optional validate() first, then persists the
 *                  result. Throws on validation failure (caller wraps in
 *                  wrapApi or its own try/catch).
 * - clear():       clears the pref (returns to defaultValue on next load).
 *                  Never throws.
 *
 * `defaultValue` is deep-cloned on every load so callers can't mutate the
 * shared instance. If you need a *frozen* default for reference equality,
 * freeze it yourself before passing it in.
 */
export function makeJsonPrefStore({
  Services,
  prefKey,
  validate,       // optional: (raw) => normalized  (throws on invalid)
  defaultValue,   // required: the shape returned when pref is unset / corrupt
  logPrefix = "thunderbird-mcp",
  // Optional logger sink. Defaults to console; unit tests pass
  // { warn() {}, error() {} } to silence the intentional fallback
  // logs from corrupt-pref / validator-throws negative-path tests.
  logger = console,
}) {
  if (!Services || !prefKey) {
    throw new Error("makeJsonPrefStore requires { Services, prefKey }");
  }

  function cloneDefault() {
    return structuredClone(defaultValue);
  }

  function load() {
    try {
      if (!Services.prefs.prefHasUserValue(prefKey)) return cloneDefault();
      const raw = Services.prefs.getStringPref(prefKey, "");
      if (!raw) return cloneDefault();
      const parsed = JSON.parse(raw);
      return validate ? validate(parsed) : parsed;
    } catch (e) {
      logger.warn(`${logPrefix}: ${prefKey} unparseable, using default:`, e);
      return cloneDefault();
    }
  }

  function save(value) {
    const normalized = validate ? validate(value) : value;
    Services.prefs.setStringPref(prefKey, JSON.stringify(normalized));
    return normalized;
  }

  function clear() {
    try { Services.prefs.clearUserPref(prefKey); } catch { /* not user-set */ }
  }

  return { load, save, clear, prefKey };
}

/**
 * Wrap an experiment-API handler body so any thrown error becomes a
 * structured { error } response instead of the Mozilla-generic
 * "An unexpected error occurred" that hides diagnostics. Synchronous and
 * async operations both work.
 *
 * @param {string} name   Handler name for the error-message prefix + log.
 * @param {Function} op   Handler body. May return a value or a Promise.
 * @returns {Promise<any>} Resolves to op()'s result, or { error: string }
 *                         if op threw / rejected.
 */
export async function wrapApi(name, op, { logger = console } = {}) {
  try {
    return await op();
  } catch (e) {
    const msg = e && e.message ? e.message : String(e);
    logger.error(`thunderbird-mcp: ${name} failed:`, e);
    return { error: `${name} failed: ${msg}` };
  }
}
