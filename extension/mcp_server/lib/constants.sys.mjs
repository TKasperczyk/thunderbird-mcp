/**
 * Pure-data constants for the Thunderbird MCP server.
 * No XPCOM, no Services, no Cc/Ci -- safe to import at any point.
 *
 * Loaded by api.js inside getAPI() via:
 *   ChromeUtils.importESModule("resource://thunderbird-mcp/lib/constants.sys.mjs")
 *
 * The MCP_* protocol-version constants must stay in sync with mcp-bridge.cjs.
 */

export const MCP_DEFAULT_PORT = 8765;
export const MCP_MAX_PORT_ATTEMPTS = 10;

// Versions of the MCP protocol this server understands. Behavior never depends
// on the negotiated version inside Thunderbird (the bridge intercepts initialize
// for clients), but for the rare case a client talks directly to the HTTP server
// we still need a spec-compliant negotiated value.
export const MCP_SUPPORTED_PROTOCOL_VERSIONS = new Set([
  "2024-11-05",
  "2025-03-26",
  "2025-06-18",
  "2025-11-25",
]);
export const MCP_LATEST_PROTOCOL_VERSION = "2025-11-25";

// 25 MB limit for inline base64 data (encoded).
export const MAX_BASE64_SIZE = 25 * 1024 * 1024;
// Must be large enough to carry MAX_BASE64_SIZE plus JSON-RPC framing overhead.
// The httpd.sys.mjs pre-buffer cap uses the same value.
export const MAX_REQUEST_BODY = 32 * 1024 * 1024;

// Delay before injecting attachments into a newly opened compose window.
export const COMPOSE_WINDOW_LOAD_DELAY_MS = 1500;

export const DEFAULT_MAX_RESULTS = 50;
export const MAX_SEARCH_RESULTS_CAP = 200;
export const SEARCH_COLLECTION_CAP = 10000;

export const PREF_ALLOWED_ACCOUNTS = "extensions.thunderbird-mcp.allowedAccounts";
export const PREF_DISABLED_TOOLS = "extensions.thunderbird-mcp.disabledTools";
export const PREF_BLOCK_SKIPREVIEW = "extensions.thunderbird-mcp.blockSkipReview";

// Valid group and CRUD values for tool metadata validation.
export const VALID_GROUPS = ["messages", "folders", "contacts", "calendar", "filters", "system"];
export const VALID_CRUD = ["create", "read", "update", "delete"];
// CRUD sort order: read first, then create, update, delete (safe -> destructive).
export const CRUD_ORDER = { read: 0, create: 1, update: 2, delete: 3 };

// Tools that cannot be disabled via the settings page (infrastructure tools).
export const UNDISABLEABLE_TOOLS = new Set(["listAccounts", "listFolders", "getAccountAccess"]);

// Internal IMAP/Thunderbird keywords that should not appear as user-visible tags.
export const INTERNAL_KEYWORDS = new Set([
  "junk", "notjunk", "$forwarded", "$replied",
  "\\seen", "\\answered", "\\flagged", "\\deleted", "\\draft", "\\recent",
  // Some IMAP servers store flags without the backslash prefix
  "seen", "answered", "flagged", "deleted", "draft", "recent",
]);
