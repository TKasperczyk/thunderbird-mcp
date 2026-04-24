/**
 * Fine-grained permission engine for Thunderbird MCP tools.
 *
 * Sits in front of the dispatch.callTool switch. For every tool call,
 * `evaluate(toolName, args, context)` resolves a hierarchical policy
 * (per-arg whitelist > per-folder rule > per-account rule > global rule
 * > sandbox default) and returns either { allow: true, args: <maybe-mutated> }
 * or { allow: false, reason: <string> }.
 *
 * Why a single module instead of inline checks: every level of the
 * resolution can override the level above (an account can override a
 * global allow with a deny, a folder rule can override an account rule,
 * etc.). Keeping the resolution in one place makes the precedence
 * provable rather than emergent from scattered if-statements.
 *
 * Storage: a single JSON pref `extensions.thunderbird-mcp.permissions`
 * holding the full policy document. The schema:
 *
 *   {
 *     "version": 1,
 *     "tools": {
 *       "<toolName>": <toolRule>          // global per-tool rule
 *     },
 *     "accounts": {
 *       "<accountKey>": {
 *         "tools": { "<toolName>": <toolRule>, "*": <toolRule> },
 *         "default": "deny" | "allow"     // when no per-tool rule matches
 *       }
 *     },
 *     "folders": {
 *       "<folderUri>": {
 *         "tools": { "<toolName>": <toolRule>, "*": <toolRule> }
 *       }
 *     }
 *   }
 *
 * `<toolRule>`:
 *   {
 *     "enabled": bool,                                    // Level 1
 *     "alwaysReview": bool,                                // Level 2 (compose only -- forces skipReview off)
 *     "maxBatchSize": number,                              // Level 2 (deleteMessages, updateMessage messageIds)
 *     "requireFolderArg": bool,                            // Level 2 (block tools that take optional folderPath when omitted)
 *     "argRules": {                                        // Level 4
 *       "<argName>": {
 *         "match"?: <regex string>,                        // value (or each entry of array value) must match
 *         "anyOf"?: [<glob string>, ...],                  // value must equal one of (glob: * = any sequence, ? = any char)
 *         "maxLength"?: number                             // string or array length cap
 *       }
 *     }
 *   }
 *
 * Resolution order for one tool call:
 *   1. find effective rule by merging:
 *        global.tools[name]                  (lowest precedence)
 *        accounts[ctx.accountKey].tools[name] (or "*")
 *        folders[ctx.folderUri].tools[name]   (or "*")
 *      (fields set in a higher-precedence layer override the same fields below)
 *   2. account.default if nothing matched
 *   3. fall back to undisableable / default-disabled sandbox
 *
 * Factory pattern matches the rest of lib/*.sys.mjs.
 */

export function makePermissions({
  Services,
  MailServices,
  PREF_PERMISSIONS,
  UNDISABLEABLE_TOOLS,
  DEFAULT_DISABLED_TOOLS,
}) {
  // ---- pref I/O ----------------------------------------------------------

  const EMPTY_POLICY = Object.freeze({ version: 1, tools: {}, accounts: {}, folders: {} });

  function loadPolicy() {
    try {
      if (!Services.prefs.prefHasUserValue(PREF_PERMISSIONS)) {
        return EMPTY_POLICY;
      }
      const raw = Services.prefs.getStringPref(PREF_PERMISSIONS, "");
      if (!raw) return EMPTY_POLICY;
      const parsed = JSON.parse(raw);
      if (!parsed || typeof parsed !== "object") return EMPTY_POLICY;
      return {
        version: parsed.version || 1,
        tools: parsed.tools && typeof parsed.tools === "object" ? parsed.tools : {},
        accounts: parsed.accounts && typeof parsed.accounts === "object" ? parsed.accounts : {},
        folders: parsed.folders && typeof parsed.folders === "object" ? parsed.folders : {},
      };
    } catch (e) {
      console.warn("thunderbird-mcp: permissions pref unparseable, using empty policy:", e);
      return EMPTY_POLICY;
    }
  }

  function savePolicy(policy) {
    if (!policy || typeof policy !== "object") {
      throw new Error("policy must be an object");
    }
    const normalized = {
      version: 1,
      tools: policy.tools && typeof policy.tools === "object" ? policy.tools : {},
      accounts: policy.accounts && typeof policy.accounts === "object" ? policy.accounts : {},
      folders: policy.folders && typeof policy.folders === "object" ? policy.folders : {},
    };
    Services.prefs.setStringPref(PREF_PERMISSIONS, JSON.stringify(normalized));
    return normalized;
  }

  function clearPolicy() {
    try { Services.prefs.clearUserPref(PREF_PERMISSIONS); } catch { /* not set, fine */ }
  }

  // ---- arg-rule helpers --------------------------------------------------

  // Glob -> regex, anchored. * = any chars, ? = single char. Other regex
  // metachars are escaped so users can't accidentally write half-broken
  // regex into the policy.
  function globToRegex(glob) {
    const esc = String(glob).replace(/[.+^${}()|[\]\\]/g, "\\$&");
    const pat = "^" + esc.replace(/\*/g, ".*").replace(/\?/g, ".") + "$";
    return new RegExp(pat, "i");
  }

  function valueMatchesArgRule(value, argRule) {
    if (!argRule) return { ok: true };

    // Apply rule to each element if value is an array.
    const values = Array.isArray(value) ? value : [value];
    if (argRule.maxLength != null) {
      if (Array.isArray(value) && value.length > argRule.maxLength) {
        return { ok: false, reason: `array exceeds maxLength=${argRule.maxLength}` };
      }
      if (typeof value === "string" && value.length > argRule.maxLength) {
        return { ok: false, reason: `string exceeds maxLength=${argRule.maxLength}` };
      }
    }
    let matchRe = null;
    if (argRule.match) {
      try { matchRe = new RegExp(argRule.match, "i"); }
      catch (e) { return { ok: false, reason: `invalid match regex: ${e.message}` }; }
    }
    let anyOfRes = null;
    if (Array.isArray(argRule.anyOf) && argRule.anyOf.length > 0) {
      anyOfRes = argRule.anyOf.map(globToRegex);
    }
    if (!matchRe && !anyOfRes) return { ok: true };

    for (const v of values) {
      if (v == null || v === "") continue; // empty values not whitelisted unless explicitly tested
      const s = String(v);
      if (matchRe && !matchRe.test(s)) {
        return { ok: false, reason: `value "${s}" does not match required pattern ${argRule.match}` };
      }
      if (anyOfRes && !anyOfRes.some(re => re.test(s))) {
        return { ok: false, reason: `value "${s}" not in allowed list (${argRule.anyOf.join(", ")})` };
      }
    }
    return { ok: true };
  }

  // ---- rule merging ------------------------------------------------------

  function mergeRule(lower, upper) {
    if (!upper) return lower || null;
    if (!lower) return { ...upper };
    return {
      ...lower,
      ...upper,
      argRules: { ...(lower.argRules || {}), ...(upper.argRules || {}) },
    };
  }

  // Resolve folderPath -> account key via MailServices.folderLookup.
  function accountKeyForFolderUri(folderUri) {
    if (!folderUri || !MailServices) return null;
    try {
      const folder = MailServices.folderLookup.getFolderForURL(folderUri);
      if (!folder || !folder.server) return null;
      const account = MailServices.accounts.findAccountForServer(folder.server);
      return account ? account.key : null;
    } catch {
      return null;
    }
  }

  function resolveContext(toolName, args) {
    // Best-effort: tools that pass an accountId arg use it directly; tools
    // that pass folderPath have their account derived. Some tools take
    // both, in which case folderPath wins (it's more specific).
    let accountKey = null;
    let folderUri = null;
    if (args && typeof args === "object") {
      if (typeof args.accountId === "string") accountKey = args.accountId;
      if (typeof args.folderPath === "string") folderUri = args.folderPath;
      if (typeof args.parentFolderPath === "string" && !folderUri) folderUri = args.parentFolderPath;
      if (typeof args.newParentPath === "string" && !folderUri) folderUri = args.newParentPath;
      if (folderUri && !accountKey) accountKey = accountKeyForFolderUri(folderUri);
    }
    return { accountKey, folderUri };
  }

  function effectiveRule(policy, toolName, ctx) {
    let rule = mergeRule(null, policy.tools && policy.tools[toolName]);
    if (ctx.accountKey && policy.accounts && policy.accounts[ctx.accountKey]) {
      const acct = policy.accounts[ctx.accountKey];
      const acctRule = (acct.tools && (acct.tools[toolName] || acct.tools["*"])) || null;
      rule = mergeRule(rule, acctRule);
      // account-level "default: deny" applies when no per-tool rule matched
      if (!acctRule && acct.default === "deny" && (!rule || rule.enabled !== false)) {
        rule = mergeRule(rule, { enabled: false });
      }
    }
    if (ctx.folderUri && policy.folders && policy.folders[ctx.folderUri]) {
      const fld = policy.folders[ctx.folderUri];
      const fldRule = (fld.tools && (fld.tools[toolName] || fld.tools["*"])) || null;
      rule = mergeRule(rule, fldRule);
    }
    return rule;
  }

  // ---- the main entry point ---------------------------------------------

  /**
   * Evaluate whether a tool call is allowed under the current policy.
   * Returns { allow: true, args: <possibly-mutated args> } or
   *         { allow: false, reason: <human-readable string> }.
   *
   * The bulk-size and alwaysReview rules can MUTATE args (e.g. force
   * skipReview=false). Caller must use the returned args.
   */
  function evaluate(toolName, rawArgs, baseAllowed = true) {
    const args = (rawArgs && typeof rawArgs === "object") ? { ...rawArgs } : {};

    // Layer 0: undisableable tools always go through, without further checks.
    if (UNDISABLEABLE_TOOLS && UNDISABLEABLE_TOOLS.has(toolName)) {
      return { allow: true, args };
    }

    const policy = loadPolicy();
    const ctx = resolveContext(toolName, args);
    const rule = effectiveRule(policy, toolName, ctx);

    // Layer 1: explicit enabled toggle (rule wins over baseAllowed)
    if (rule && rule.enabled === false) {
      return { allow: false, reason: `Tool '${toolName}' is disabled by policy` + (ctx.accountKey ? ` for account ${ctx.accountKey}` : "") + (ctx.folderUri ? ` in folder ${ctx.folderUri}` : "") };
    }
    if (rule && rule.enabled === true) {
      // explicit allow -- continue to constraint checks
    } else if (!baseAllowed) {
      // No explicit allow and the caller told us this tool is in the
      // sandbox-default-disabled set -- deny.
      return { allow: false, reason: `Tool '${toolName}' is off by default. Enable it in the extension settings page.` };
    }

    if (!rule) return { allow: true, args };

    // Layer 2: per-tool constraints
    if (rule.requireFolderArg && !ctx.folderUri) {
      return { allow: false, reason: `Tool '${toolName}' requires a folderPath argument under current policy` };
    }
    if (rule.maxBatchSize != null) {
      const bulk = Array.isArray(args.messageIds) ? args.messageIds
                 : Array.isArray(args.contactIds) ? args.contactIds
                 : Array.isArray(args.eventIds) ? args.eventIds
                 : null;
      if (bulk && bulk.length > rule.maxBatchSize) {
        return { allow: false, reason: `Tool '${toolName}' batch size ${bulk.length} exceeds policy maxBatchSize=${rule.maxBatchSize}` };
      }
    }
    if (rule.alwaysReview === true && args.skipReview === true) {
      // Force compose tools to open the review window regardless of what
      // the caller passed. Mutate args, log nothing here -- compose code
      // already logs a clear message when blockSkipReview kicks in.
      args.skipReview = false;
    }

    // Layer 4: argument whitelists
    if (rule.argRules && typeof rule.argRules === "object") {
      for (const [argName, argRule] of Object.entries(rule.argRules)) {
        if (!Object.prototype.hasOwnProperty.call(args, argName)) continue;
        const r = valueMatchesArgRule(args[argName], argRule);
        if (!r.ok) {
          return { allow: false, reason: `Tool '${toolName}' arg '${argName}' rejected by policy: ${r.reason}` };
        }
      }
    }

    return { allow: true, args };
  }

  return { evaluate, loadPolicy, savePolicy, clearPolicy };
}
