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

// MCP protocol versions this server negotiates. This list tracks the bundled
// @modelcontextprotocol/sdk inside Claude Desktop, NOT the latest published
// spec. See mcp-bridge.cjs for the full reasoning -- echoing "2025-11-25"
// back to a client SDK that was pinned at LATEST_PROTOCOL_VERSION='2025-06-18'
// before the November spec landed makes the client silently reject the
// negotiated version and time out at 60 seconds. Filesystem MCP returns
// "2025-06-18" and attaches fine -- we mirror that.
export const MCP_SUPPORTED_PROTOCOL_VERSIONS = new Set([
  "2024-10-07",
  "2024-11-05",
  "2025-03-26",
  "2025-06-18",
]);
export const MCP_LATEST_PROTOCOL_VERSION = "2025-06-18";

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
// Fine-grained permission policy (JSON document, see lib/permissions.sys.mjs).
export const PREF_PERMISSIONS = "extensions.thunderbird-mcp.permissions";

// Valid group and CRUD values for tool metadata validation.
export const VALID_GROUPS = ["messages", "folders", "contacts", "calendar", "filters", "system"];
export const VALID_CRUD = ["create", "read", "update", "delete"];
// CRUD sort order: read first, then create, update, delete (safe -> destructive).
export const CRUD_ORDER = { read: 0, create: 1, update: 2, delete: 3 };

// Tools that cannot be disabled via the settings page (infrastructure tools).
export const UNDISABLEABLE_TOOLS = new Set(["listAccounts", "listFolders", "getAccountAccess"]);

// Tools that ship DISABLED by default. The user has to explicitly opt-in via
// the settings page before AI can call them. Only applies when the user has
// NEVER touched the pref -- once they save any tool-access choice (even an
// empty disabled list = "I want everything enabled"), this default no longer
// overrides their explicit choice.
//
// Policy: any tool that sends mail off the machine OR irreversibly destroys
// user data is opt-in. read/list/search/update tools stay enabled by default
// because they're either safe or recoverable. Create tools for single items
// (createContact, createEvent, ...) stay enabled because creating clutter is
// annoying but not destructive.
//
// Big-hammer batch tools are also opt-in even when their individual actions
// would be default-enabled -- a mistake scales from "one bad move" to "all
// messages in a folder moved to the wrong place" or "40 duplicate drafts".
export const DEFAULT_DISABLED_TOOLS = new Set([
  // send: leaves the machine, can't be undone
  "sendMail", "replyToMessage", "forwardMessage",
  // delete: irreversible (or near-irreversible: trash gets emptied eventually)
  "deleteContact", "deleteEvent", "deleteFilter", "deleteMessages",
  "deleteFolder", "emptyTrash", "emptyJunk",
  // batch / bulk: dry_run=true default caps the first call, but once the caller
  // flips it to false it affects an unbounded number of user objects. Opt-in.
  "bulk_move_by_query", "createDrafts",
]);

// Internal IMAP/Thunderbird keywords that should not appear as user-visible tags.
export const INTERNAL_KEYWORDS = new Set([
  "junk", "notjunk", "$forwarded", "$replied",
  "\\seen", "\\answered", "\\flagged", "\\deleted", "\\draft", "\\recent",
  // Some IMAP servers store flags without the backslash prefix
  "seen", "answered", "flagged", "deleted", "draft", "recent",
]);
