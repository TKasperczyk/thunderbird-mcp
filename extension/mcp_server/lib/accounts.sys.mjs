/**
 * Account + identity lookups for the Thunderbird MCP server.
 *
 * Surfaces accessible accounts (filtered through the access-control
 * policy) and resolves identities by key or email. All identity lookups
 * route through getAccessibleAccounts so restricted accounts stay
 * invisible to MCP callers; the only escape hatch is findIdentityIn,
 * which lets the caller pass MailServices.accounts.accounts directly
 * for diagnostic "exists but blocked" checks.
 *
 * Factory: api.js calls makeAccounts({ MailServices, getAccessibleAccounts }).
 * Cc / Ci / Services aren't needed -- these helpers only walk
 * MailServices objects and the injected accessor.
 */

export function makeAccounts({ MailServices, getAccessibleAccounts }) {
  function listAccounts() {
    const accounts = [];
    for (const account of getAccessibleAccounts()) {
      const server = account.incomingServer;
      const identities = [];
      for (const identity of account.identities) {
        identities.push({
          id: identity.key,
          email: identity.email,
          name: identity.fullName,
          isDefault: identity === account.defaultIdentity
        });
      }
      accounts.push({
        id: account.key,
        name: server.prettyName,
        type: server.type,
        identities
      });
    }
    return accounts;
  }

  /**
   * Searches the given accounts for an identity matching emailOrId
   * (by key or case-insensitive email).
   * Returns the identity object, or null if not found.
   */
  function findIdentityIn(accounts, emailOrId) {
    if (!emailOrId) return null;
    const lowerInput = emailOrId.toLowerCase();
    for (const account of accounts) {
      for (const identity of account.identities) {
        if (identity.key === emailOrId || (identity.email || "").toLowerCase() === lowerInput) {
          return identity;
        }
      }
    }
    return null;
  }

  /**
   * Finds an identity by email address or identity ID
   * among accessible accounts only.  Returns null if not found.
   */
  function findIdentity(emailOrId) {
    return findIdentityIn(getAccessibleAccounts(), emailOrId);
  }

  function getIdentityAutoRecipientHeader(identity, kind) {
    if (!identity) return "";
    try {
      if (kind === "cc") {
        return identity.doCc ? (identity.doCcList || "") : "";
      }
      if (kind === "bcc") {
        return identity.doBcc ? (identity.doBccList || "") : "";
      }
    } catch {}
    return "";
  }

  return { listAccounts, findIdentityIn, findIdentity, getIdentityAutoRecipientHeader };
}
