/**
 * Account-, folder-, and tool-access policy for the Thunderbird MCP server.
 *
 * Centralizes the "is this caller allowed to touch X" checks that every
 * read/write tool runs before doing real work. Backed by Thunderbird
 * preferences (PREF_ALLOWED_ACCOUNTS, PREF_DISABLED_TOOLS, PREF_BLOCK_SKIPREVIEW).
 *
 * All checks fail CLOSED on corrupt prefs -- a malformed pref disables
 * everything rather than silently allowing it. The settings UI reads the
 * same values via getAccountAccess() / getDisabledTools() so it
 * accurately reflects the server's actual behavior even when corrupt.
 *
 * Factory: api.js calls makeAccessControl({ Services, MailServices,
 * PREF_ALLOWED_ACCOUNTS, PREF_DISABLED_TOOLS, PREF_BLOCK_SKIPREVIEW,
 * UNDISABLEABLE_TOOLS }).
 */

export function makeAccessControl({
  Services,
  MailServices,
  PREF_ALLOWED_ACCOUNTS,
  PREF_DISABLED_TOOLS,
  PREF_BLOCK_SKIPREVIEW,
  UNDISABLEABLE_TOOLS,
}) {
  /**
   * Return the list of allowed account IDs from the pref. Empty array
   * means "no restriction set". Returns the sentinel ["__invalid__"] on
   * any parse failure so isAccountAllowed denies everything.
   */
  function getAllowedAccountIds() {
    try {
      const pref = Services.prefs.getStringPref(PREF_ALLOWED_ACCOUNTS, "");
      if (!pref) return [];
      const parsed = JSON.parse(pref);
      if (!Array.isArray(parsed)) {
        console.error("thunderbird-mcp: allowed accounts pref is not an array, blocking all accounts");
        return ["__invalid__"];
      }
      return parsed;
    } catch (e) {
      console.error("thunderbird-mcp: failed to parse allowed accounts pref, blocking all accounts:", e);
      return ["__invalid__"];
    }
  }

  function isAccountAllowed(accountKey) {
    const allowed = getAllowedAccountIds();
    if (allowed.length === 0) return true;
    return allowed.includes(accountKey);
  }

  /**
   * Whether the user has disabled the skipReview shortcut. Defaults to
   * blocked on read failure -- the user retains review ability.
   */
  function isSkipReviewBlocked() {
    try {
      return Services.prefs.getBoolPref(PREF_BLOCK_SKIPREVIEW, false);
    } catch {
      return true;
    }
  }

  /**
   * Disabled tool names from the pref. Returns ["__all__"] sentinel on
   * any failure so isToolEnabled denies everything except undisableables.
   */
  function getDisabledTools() {
    try {
      const pref = Services.prefs.getStringPref(PREF_DISABLED_TOOLS, "");
      if (!pref) return [];
      const parsed = JSON.parse(pref);
      if (!Array.isArray(parsed) || !parsed.every(v => typeof v === "string")) {
        console.error("thunderbird-mcp: disabled tools pref is invalid, disabling all tools");
        return ["__all__"];
      }
      return parsed;
    } catch (e) {
      console.error("thunderbird-mcp: failed to parse disabled tools pref, disabling all tools:", e);
      return ["__all__"];
    }
  }

  function isToolEnabled(toolName) {
    if (UNDISABLEABLE_TOOLS.has(toolName)) return true;
    const disabled = getDisabledTools();
    if (disabled.includes("__all__")) return false;
    return !disabled.includes(toolName);
  }

  /**
   * True if the resolved folder belongs to an allowed account.
   */
  function isFolderAccessible(folder) {
    if (!folder || !folder.server) return false;
    const account = MailServices.accounts.findAccountForServer(folder.server);
    return account ? isAccountAllowed(account.key) : false;
  }

  /**
   * Resolve a folder URI and verify accessibility. Returns { folder } on
   * success, or { error } if missing or restricted.
   */
  function getAccessibleFolder(folderPath) {
    const folder = MailServices.folderLookup.getFolderForURL(folderPath);
    if (!folder) return { error: `Folder not found: ${folderPath}` };
    if (!isFolderAccessible(folder)) return { error: `Account not accessible for folder: ${folderPath}` };
    return { folder };
  }

  function getAccessibleAccounts() {
    const result = [];
    for (const account of MailServices.accounts.accounts) {
      if (isAccountAllowed(account.key)) {
        result.push(account);
      }
    }
    return result;
  }

  /**
   * Public tool handler -- exposes the current account-access policy.
   * Restricted accounts are hidden, not just flagged.
   */
  function getAccountAccess() {
    const allowed = getAllowedAccountIds();
    const accessibleAccounts = [];
    for (const account of MailServices.accounts.accounts) {
      if (!isAccountAllowed(account.key)) continue;
      const server = account.incomingServer;
      accessibleAccounts.push({
        id: account.key,
        name: server.prettyName,
        type: server.type,
      });
    }
    return {
      mode: allowed.length === 0 ? "all" : "restricted",
      accounts: accessibleAccounts,
    };
  }

  return {
    getAllowedAccountIds,
    isAccountAllowed,
    isSkipReviewBlocked,
    getDisabledTools,
    isToolEnabled,
    isFolderAccessible,
    getAccessibleFolder,
    getAccessibleAccounts,
    getAccountAccess,
  };
}
