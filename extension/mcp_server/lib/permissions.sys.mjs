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
  makeJsonPrefStore,   // injected factory from lib/json-pref-store.sys.mjs
}) {
  // ---- pref I/O via the shared JSON-pref-store pattern -------------------

  // Shape normalizer: this runs on both load (after JSON.parse) and save.
  // Anything that isn't the expected object-of-objects shape is coerced to
  // an empty sub-section rather than throwing, so a partial document from
  // the options-page editor doesn't lose the user's other sections.
  function normalizePolicy(raw) {
    if (!raw || typeof raw !== "object") {
      throw new Error("policy must be an object");
    }
    return {
      version: raw.version || 1,
      tools:    (raw.tools    && typeof raw.tools    === "object") ? raw.tools    : {},
      accounts: (raw.accounts && typeof raw.accounts === "object") ? raw.accounts : {},
      folders:  (raw.folders  && typeof raw.folders  === "object") ? raw.folders  : {},
    };
  }

  const store = makeJsonPrefStore({
    Services,
    prefKey: PREF_PERMISSIONS,
    validate: normalizePolicy,
    defaultValue: { version: 1, tools: {}, accounts: {}, folders: {} },
  });

  // ---- regex compilation + catastrophic-backtracking guard ---------------
  //
  // CodeRabbit flagged ReDoS risk: argRule.match was compiled with `new RegExp`
  // on every tool call, and the input is operator-controlled. We pre-compile
  // and cache per arg-rule object via a WeakMap keyed by the rule, plus reject
  // obviously-dangerous shapes (nested quantifiers, overlong patterns) at load
  // time. Invalid rules are demoted to "<<invalid-skip>>" which never matches.

  const _compiledMatch = new WeakMap();
  const SUSPICIOUS_RE = /\([^)]*[+*][^)]*\)[+*]|\(\.\*\)\*|\(\.\+\)\+/;

  function isLikelyCatastrophic(pattern) {
    if (typeof pattern !== "string") return true;
    if (pattern.length > 200) return true;
    if (SUSPICIOUS_RE.test(pattern)) return true;
    return false;
  }

  function precompileArgRuleMatch(argRule) {
    if (!argRule || typeof argRule !== "object") return null;
    if (_compiledMatch.has(argRule)) return _compiledMatch.get(argRule);
    if (!argRule.match) {
      _compiledMatch.set(argRule, null);
      return null;
    }
    let compiled = null;
    if (isLikelyCatastrophic(argRule.match)) {
      try {
        console.warn(
          `[permissions] arg-rule match pattern rejected as potentially catastrophic; ` +
          `treating as <<invalid-skip>>: ${String(argRule.match).slice(0, 80)}`
        );
      } catch { /* console missing -- ignore */ }
      compiled = { invalid: true };
    } else {
      try {
        compiled = { re: new RegExp(argRule.match, "i") };
      } catch (e) {
        try {
          console.warn(`[permissions] invalid regex in arg-rule match: ${e.message}`);
        } catch { /* ignore */ }
        compiled = { invalid: true, reason: e.message };
      }
    }
    _compiledMatch.set(argRule, compiled);
    return compiled;
  }

  // Walk the whole policy once after load to pre-compile every match-regex.
  // Cheap because each rule object is only walked the first time it appears;
  // subsequent loadPolicy() calls return the same cached objects.
  function precompileAllRuleRegexes(policy) {
    if (!policy || typeof policy !== "object") return policy;
    const visit = (toolMap) => {
      if (!toolMap || typeof toolMap !== "object") return;
      for (const ruleKey of Object.keys(toolMap)) {
        const rule = toolMap[ruleKey];
        if (!rule || typeof rule !== "object" || !rule.argRules) continue;
        for (const argName of Object.keys(rule.argRules)) {
          precompileArgRuleMatch(rule.argRules[argName]);
        }
      }
    };
    visit(policy.tools);
    if (policy.accounts) {
      for (const a of Object.values(policy.accounts)) visit(a && a.tools);
    }
    if (policy.folders) {
      for (const f of Object.values(policy.folders)) visit(f && f.tools);
    }
    return policy;
  }

  function loadPolicy() {
    const p = store.load();
    precompileAllRuleRegexes(p);
    return p;
  }
  const savePolicy = store.save;
  const clearPolicy = store.clear;

  // ---- arg-rule helpers --------------------------------------------------

  // Some MCP clients (and our own deleteMessages) accept arrays serialized as
  // JSON strings. The permission engine must apply array-length caps to the
  // *parsed* array, otherwise a caller can ship `messageIds: "[...10000 ids...]"`
  // and slip past Array.isArray-gated checks entirely. (Claude review §2a.)
  function asArrayOrNull(v) {
    if (Array.isArray(v)) return v;
    if (typeof v === "string") {
      try {
        const parsed = JSON.parse(v);
        return Array.isArray(parsed) ? parsed : null;
      } catch { return null; }
    }
    return null;
  }

  // True iff any per-folder policy entry has a rule for this tool. Used to
  // reject includeSubfolders=true when per-folder rules exist, because a
  // recursive walk would otherwise process subfolders that may be denied
  // by their own policy entries. Pragmatic mitigation -- a full per-subfolder
  // expansion would require reaching into the folder tree from here. (Claude
  // review §2d.)
  function policyHasPerFolderRulesForTool(policy, toolName) {
    if (!policy || !policy.folders) return false;
    for (const folderUri of Object.keys(policy.folders)) {
      const fld = policy.folders[folderUri];
      const rule = fld && fld.tools && (fld.tools[toolName] || fld.tools["*"]);
      if (rule) return true;
    }
    return false;
  }

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
      const compiled = precompileArgRuleMatch(argRule);
      if (compiled && compiled.invalid) {
        // Rule was demoted at load-time (catastrophic-backtracking shape or
        // unparseable). Treat as never-matching so we always deny rather than
        // accidentally letting unfiltered input through.
        return { ok: false, reason: `arg-rule match regex was rejected at policy load; value not whitelisted` };
      }
      matchRe = compiled ? compiled.re : null;
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

  // Tools that operate on TWO folders (source + destination). The per-folder
  // policy must be evaluated against BOTH or the destination rule is silently
  // ignored. (Claude review §2b.) Today only bulk_move_by_query has this
  // shape; the table is here so future tools (move_message, copyMessage, …)
  // are an additive change.
  const DEST_FOLDER_ARG_BY_TOOL = Object.freeze({
    bulk_move_by_query: "to",
  });

  function isKnownAccountKey(accountKey) {
    if (!accountKey || !MailServices || !MailServices.accounts) return false;
    try {
      const accounts = MailServices.accounts.accounts;
      if (!accounts) return false;
      // .accounts can be an nsIArray on real Thunderbird or a plain JS array
      // in the test harness; support both shapes.
      if (typeof accounts.some === "function") {
        return accounts.some(a => a && a.key === accountKey);
      }
      // nsIArray-style: iterate via .length / queryElementAt is overkill --
      // most callers use `for...of`. Fall through to that.
      for (const a of accounts) {
        if (a && a.key === accountKey) return true;
      }
      return false;
    } catch { return false; }
  }

  function resolveContext(toolName, args) {
    // Best-effort: tools that pass an accountId arg use it directly; tools
    // that pass folderPath have their account derived. Some tools take
    // both, in which case folderPath wins (it's more specific).
    let accountKey = null;
    let folderUri = null;
    let destFolderUri = null;
    if (args && typeof args === "object") {
      if (typeof args.accountId === "string") accountKey = args.accountId;
      if (typeof args.folderPath === "string") folderUri = args.folderPath;
      if (typeof args.parentFolderPath === "string" && !folderUri) folderUri = args.parentFolderPath;
      if (typeof args.newParentPath === "string" && !folderUri) folderUri = args.newParentPath;
      if (folderUri && !accountKey) accountKey = accountKeyForFolderUri(folderUri);

      // Destination folder for tools that have one (e.g. bulk_move_by_query.to).
      const destArg = DEST_FOLDER_ARG_BY_TOOL[toolName];
      if (destArg && typeof args[destArg] === "string" && args[destArg]) {
        destFolderUri = args[destArg];
      }
    }
    // folderUris = [src, dst] filtered to truthy + deduped. Evaluate denies if
    // any one of them is rule-denied.
    const folderUris = [];
    if (folderUri) folderUris.push(folderUri);
    if (destFolderUri && destFolderUri !== folderUri) folderUris.push(destFolderUri);
    return { accountKey, folderUri, destFolderUri, folderUris };
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

  // Returns just the per-folder rule for a single folderUri (no merging with
  // globals/account). Used to independently check destination-folder denials
  // for tools like bulk_move_by_query, where the source rule was already
  // merged into the main `rule` via effectiveRule. (Claude review §2b.)
  function perFolderRule(policy, toolName, folderUri) {
    if (!folderUri || !policy.folders) return null;
    const fld = policy.folders[folderUri];
    if (!fld || !fld.tools) return null;
    return fld.tools[toolName] || fld.tools["*"] || null;
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

    // Reject unknown accountId early: otherwise effectiveRule would silently
    // skip the account.default=deny branch because the bogus key isn't in
    // policy.accounts. (Claude review §2c.) We only validate if the caller
    // explicitly passed an accountId arg; tools that derive accountKey from
    // folderUri can't trigger this -- a valid folder always maps to a real
    // account or returns null.
    if (typeof args.accountId === "string" && args.accountId && !isKnownAccountKey(args.accountId)) {
      return { allow: false, reason: `Unknown accountId '${args.accountId}'` };
    }

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

    // Destination folder check: if the tool has a destination folder arg and
    // that destination has its own per-folder rule that disables this tool,
    // deny regardless of the source-folder decision. (Claude review §2b.)
    if (ctx.destFolderUri && ctx.destFolderUri !== ctx.folderUri) {
      const dstRule = perFolderRule(policy, toolName, ctx.destFolderUri);
      if (dstRule && dstRule.enabled === false) {
        return { allow: false, reason: `Tool '${toolName}' is disabled by policy in destination folder ${ctx.destFolderUri}` };
      }
    }

    // Recursive-walk safeguard: if the policy defines per-folder rules for
    // this tool, do not let includeSubfolders=true bypass them. (Claude
    // review §2d.) A complete fix would expand the folder tree here, but
    // that requires nsIMsgFolder access this module deliberately doesn't
    // have. The pragmatic policy: opt out of recursion as soon as ANY
    // per-folder rule exists for the tool.
    if (args.includeSubfolders === true && policyHasPerFolderRulesForTool(policy, toolName)) {
      return {
        allow: false,
        reason: `Tool '${toolName}': includeSubfolders disabled because per-folder policy rules exist for this tool`
      };
    }

    if (!rule) return { allow: true, args };

    // Layer 2: per-tool constraints
    if (rule.requireFolderArg && !ctx.folderUri) {
      return { allow: false, reason: `Tool '${toolName}' requires a folderPath argument under current policy` };
    }
    if (rule.maxBatchSize != null) {
      // CodeRabbit + Claude §2a: array-args may arrive as JSON strings (the
      // tools themselves JSON.parse them). Use asArrayOrNull so a stringified
      // payload of 10 000 ids can't slip past Array.isArray.
      const bulk = asArrayOrNull(args.messageIds) ||
                   asArrayOrNull(args.contactIds) ||
                   asArrayOrNull(args.eventIds) ||
                   asArrayOrNull(args.drafts);
      if (bulk && bulk.length > rule.maxBatchSize) {
        return { allow: false, reason: `Tool '${toolName}' batch size ${bulk.length} exceeds policy maxBatchSize=${rule.maxBatchSize}` };
      }
      // Special case: bulk_move_by_query doesn't carry an array arg -- its
      // batch size is the `limit` numeric arg. CodeRabbit caught that
      // omitting `limit` entirely fell through to the tool's default of
      // 1000 instead of being capped, so we clamp here (mirrors the
      // alwaysReview->args.skipReview mutation pattern above).
      if (toolName === "bulk_move_by_query") {
        if (typeof args.limit !== "number" || args.limit > rule.maxBatchSize) {
          args.limit = rule.maxBatchSize;
        }
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
