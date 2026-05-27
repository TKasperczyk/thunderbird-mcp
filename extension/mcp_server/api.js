/* global ExtensionCommon, ChromeUtils, Services, Cc, Ci */
"use strict";

/**
 * Thunderbird MCP Server Extension
 * Exposes email, calendar, and contacts via MCP protocol over HTTP.
 *
 * Architecture: MCP Client <-> mcp-bridge.cjs (stdio<->HTTP) <-> This extension (port 8765)
 *
 * Key quirks documented inline:
 * - MIME header decoding (mime2Decoded* properties)
 * - HTML body charset handling (emojis require HTML entity encoding)
 * - Compose window body preservation (must use New type, not Reply)
 * - IMAP folder sync (msgDatabase may be stale)
 */

const resProto = Cc[
  "@mozilla.org/network/protocol;1?name=resource"
].getService(Ci.nsISubstitutingProtocolHandler);

const MCP_DEFAULT_PORT = 8765;
const MCP_MAX_PORT_ATTEMPTS = 10;

// Versions of the MCP protocol this server understands. Behavior never depends
// on the negotiated version inside Thunderbird (the bridge intercepts initialize
// for clients), but for the rare case a client talks directly to the HTTP server
// we still need a spec-compliant negotiated value. Keep in sync with mcp-bridge.cjs.
const MCP_SUPPORTED_PROTOCOL_VERSIONS = new Set([
  "2024-10-07",
  "2024-11-05",
  "2025-03-26",
  "2025-06-18",
  "2025-11-25",
]);
const MCP_LATEST_PROTOCOL_VERSION = "2025-11-25";

// Bridged into serverInfo.version on initialize. Resolved lazily from the
// extension manifest so a single bump in extension/manifest.json propagates here.
let _cachedExtVersion = null;
function getExtVersion() {
  if (_cachedExtVersion) return _cachedExtVersion;
  try {
    const uri = Services.io.newURI("resource://thunderbird-mcp/manifest.json");
    const channel = Services.io.newChannelFromURI(uri, null,
      Services.scriptSecurityManager.getSystemPrincipal(), null,
      Ci.nsILoadInfo.SEC_ALLOW_CROSS_ORIGIN_SEC_CONTEXT_IS_NULL,
      Ci.nsIContentPolicy.TYPE_OTHER);
    const sis = Cc["@mozilla.org/scriptableinputstream;1"]
      .createInstance(Ci.nsIScriptableInputStream);
    sis.init(channel.open());
    const text = sis.read(sis.available());
    sis.close();
    _cachedExtVersion = JSON.parse(text).version || "0.0.0";
  } catch (e) {
    console.warn("thunderbird-mcp: could not read extension manifest version:", e);
    _cachedExtVersion = "0.0.0";
  }
  return _cachedExtVersion;
}
// Keep references to active attach timers to prevent GC before they fire.
const _attachTimers = new Set();
// Track temp files created for inline base64 attachments (cleaned up on shutdown).
const _tempAttachFiles = new Set();
// Track compose windows already claimed by an in-flight replyToMessage or
// forwardMessage call, so concurrent reply/forward operations on the same
// original message never bind two observers to the same compose window
// (which would double-inject the body/attachments).
// WeakSet so entries are collected automatically when the window is destroyed.
const _claimedComposeWindows = new WeakSet();
// Must be large enough to carry the inline-attachment base64 cap plus JSON-RPC
// framing overhead. The httpd.sys.mjs pre-buffer cap uses the same value.
const MAX_REQUEST_BODY = 32 * 1024 * 1024; // 32 MB limit for incoming HTTP request bodies

// Pure helpers (deny-list, header sanitizer, URL allow-lists, schema walker,
// audit helpers) live in security_helpers.js so the Node test suite can
// exercise the same code that ships in production. The module is loaded
// inside getAPI() below because resource:// URIs aren't usable at file scope.
const DEFAULT_MAX_RESULTS = 50;
const PREF_ALLOWED_ACCOUNTS = "extensions.thunderbird-mcp.allowedAccounts";
const PREF_DISABLED_TOOLS = "extensions.thunderbird-mcp.disabledTools";
const PREF_BLOCK_SKIPREVIEW = "extensions.thunderbird-mcp.blockSkipReview";
const PREF_STABLE_AUTH_TOKEN = "extensions.thunderbird-mcp.stableAuthToken";
// Gate `forward` and `reply` actions on filter creation/update. Filters run on
// every incoming message and never open a review UI, so a single createFilter
// call can install a permanent silent-exfiltration rule. Default is to refuse
// these action types via the MCP API; users who legitimately need them can
// flip this pref or create the filter manually in Thunderbird.
const PREF_BLOCK_FILTER_FORWARD_REPLY = "extensions.thunderbird-mcp.blockFilterForwardReply";
// Gate write operations on address books. createContact / updateContact /
// deleteContact have no UI confirmation, span every address book the user has
// configured, and could be used to spoof an existing contact (change Boss's
// email to attacker@evil.com) so the user's future replies are silently
// misrouted. Default is to refuse contact writes; users who want LLM-driven
// contact management opt in via the options page.
const PREF_BLOCK_CONTACT_WRITES = "extensions.thunderbird-mcp.blockContactWrites";
// Gate exportMailbox. Bulk-exports walk thousands of messages, may run for
// many seconds, and write the user's mail content to a JSON file on disk.
// Default off-by-pref so an LLM cannot mass-extract mailbox content silently.
const PREF_BLOCK_MAILBOX_EXPORT = "extensions.thunderbird-mcp.blockMailboxExport";
const AUTH_TOKEN_PATTERN = /^[0-9a-f]{64}$/;
// Valid group and CRUD values for tool metadata validation
const VALID_GROUPS = ["messages", "folders", "contacts", "calendar", "filters", "system"];
const VALID_CRUD = ["create", "read", "update", "delete"];
// CRUD sort order: read first, then create, update, delete (safe → destructive)
const CRUD_ORDER = { read: 0, create: 1, update: 2, delete: 3 };
// Tools that cannot be disabled via the settings page (infrastructure tools)
const UNDISABLEABLE_TOOLS = new Set(["listAccounts", "listFolders", "getAccountAccess"]);
const MAX_SEARCH_RESULTS_CAP = 200;
// Internal IMAP/Thunderbird keywords that should not appear as user-visible tags
const INTERNAL_KEYWORDS = new Set([
  "junk", "notjunk", "$forwarded", "$replied",
  "\\seen", "\\answered", "\\flagged", "\\deleted", "\\draft", "\\recent",
  // Some IMAP servers store flags without the backslash prefix
  "seen", "answered", "flagged", "deleted", "draft", "recent",
]);

var mcpServer = class extends ExtensionCommon.ExtensionAPI {
  getAPI(context) {
    const extensionRoot = context.extension.rootURI;
    const resourceName = "thunderbird-mcp";

    resProto.setSubstitutionWithFlags(
      resourceName,
      extensionRoot,
      resProto.ALLOW_CONTENT_ACCESS
    );

    // Load pure helpers from security_helpers.js. We pass a scope object with
    // a `module.exports` shim so the file can use standard CommonJS exports
    // syntax; that lets the Node test suite require() it directly without a
    // dual-format wrapper.
    const __helperScope = { module: { exports: {} } };
    Services.scriptloader.loadSubScript(
      "resource://thunderbird-mcp/mcp_server/security_helpers.js?" + Date.now(),
      __helperScope
    );
    const {
      isSensitiveFilePath,
      sanitizeHeaderLine,
      UNTRUSTED_CLOSE,
      wrapUntrustedBody,
      wrapUntrustedPreview,
      countRecipients,
      summarizeAttachmentsForAudit,
      isSafeMarkdownHref,
      isSafeImageSrc,
      escapeMarkdownLinkText,
      renderMarkdownLink,
      isSystemPrincipalFetchAllowed,
      validateAgainstSchema,
      createRateLimiterState,
      consumeRateLimit,
      inspectRateLimits,
    } = __helperScope.module.exports;

    // One rate-limiter state per server instance. Resets on extension reload
    // -- intentional: that's typically a developer action and we don't want a
    // restart to feel "stuck" for the windowMs duration.
    const __rateLimiterState = createRateLimiterState();

    const tools = [
      {
        name: "listAccounts",
        group: "system", crud: "read",
        title: "List Accounts",
        description: "List all email accounts and their identities",
        inputSchema: { type: "object", properties: {}, required: [] },
      },
      {
        name: "listFolders",
        group: "system", crud: "read",
        title: "List Folders",
        description: "List all mail folders with URIs and message counts",
        inputSchema: {
          type: "object",
          properties: {
            accountId: { type: "string", description: "Optional account ID (from listAccounts) to limit results to a single account" },
            folderPath: { type: "string", description: "Optional folder URI (from listFolders) to list only that folder and its subfolders" },
          },
          required: [],
        },
      },
      {
        name: "searchMessages",
        group: "messages", crud: "read",
        title: "Search Mail",
        description: "Search message headers and return IDs/folder paths you can use with getMessage to read full email content",
        inputSchema: {
          type: "object",
          properties: {
            query: { type: "string", description: "Text to search. Multi-word queries are AND-of-tokens: every word must appear somewhere across subject/author/recipients/ccList/preview (or inside the selected field when an operator is used). Prefix with 'from:', 'subject:', 'to:', or 'cc:' to restrict matching to one field (e.g. 'from:Alice Smith' requires both tokens in the author field). Use empty string to match all." },
            folderPath: { type: "string", description: "Optional folder URI (from listFolders) to limit search to that folder and its subfolders" },
            startDate: { type: "string", description: "Filter messages on or after this ISO 8601 date" },
            endDate: { type: "string", description: "Filter messages on or before this ISO 8601 date. Date-only strings (e.g. '2024-01-15') include the full day." },
            maxResults: { type: "number", description: "Maximum number of results to return (default 50, max 200)" },
            offset: { type: "number", description: "Number of results to skip for pagination (default 0). When provided, returns {messages, totalMatches, offset, limit, hasMore} instead of a plain array. Note: totalMatches is capped at 10000." },
            sortOrder: { type: "string", description: "Date sort order: asc (oldest first) or desc (newest first, default)" },
            unreadOnly: { type: "boolean", description: "Only return unread messages (default: false)" },
            flaggedOnly: { type: "boolean", description: "Only return flagged/starred messages (default: false)" },
            tag: { type: "string", description: "Filter by tag keyword (e.g. '$label1' for Important, or a custom tag). Only messages with this tag are returned." },
            includeSubfolders: { type: "boolean", description: "If false, only search the specified folder — not its subfolders. Default: true." },
            countOnly: { type: "boolean", description: "If true, return only the match count instead of full results. Much faster for 'how many unread?' queries." },
            searchBody: { type: "boolean", description: "If true, search full message bodies using Thunderbird's Gloda index (slower but finds text beyond the ~200 char preview). Requires query. IMAP accounts need offline sync enabled for body indexing." },
          },
          required: ["query"],
        },
      },
      {
        name: "getMessage",
        group: "messages", crud: "read",
        title: "Get Message",
        description: "Read the full content of an email message by its ID",
        inputSchema: {
          type: "object",
          properties: {
            messageId: { type: "string", description: "The message ID (from searchMessages results)" },
            folderPath: { type: "string", description: "The folder URI path (from searchMessages results)" },
            saveAttachments: { type: "boolean", description: "If true, save attachments to <OS temp dir>/thunderbird-mcp/<messageId>/ and include filePath in response (default: false)" },
            bodyFormat: { type: "string", enum: ["markdown", "text", "html"], description: "Body output format: 'markdown' (default, preserves structure), 'text' (plain text), 'html' (raw HTML)" },
            rawSource: { type: "boolean", description: "If true, return the full raw RFC 2822 message source (all headers + MIME parts). Useful for extracting calendar invites, S/MIME data, or debugging. Other fields (body, attachments) are omitted when this is set. Note: requires local/offline message copy; IMAP messages not cached offline may fail." },
          },
          required: ["messageId", "folderPath"],
        },
      },
      {
        name: "getMessageHeaders",
        group: "messages", crud: "read",
        title: "Get Message Headers",
        description: "Read message headers only (subject, author, recipients, date, tags, threading headers) without decoding the body or attachments. Much cheaper than getMessage when you only need to triage or quote a small set of messages.",
        inputSchema: {
          type: "object",
          properties: {
            messageId: { type: "string", description: "The message ID (from searchMessages results)" },
            folderPath: { type: "string", description: "The folder URI path (from searchMessages results)" },
          },
          required: ["messageId", "folderPath"],
        },
      },
      {
        name: "exportMailbox",
        group: "messages", crud: "read",
        title: "Export Mailbox",
        description: "Stream a folder's messages to a JSON-lines file under <ProfD>/thunderbird-mcp/exports/. Default mode is headers-only; pass includeBody:true to also embed each message's plain-text body and attachment metadata (slower). The destination file is named with a UTC timestamp and a sanitized folder name; the path is returned so the caller knows where to find it. Disabled by default via extensions.thunderbird-mcp.blockMailboxExport -- enable it in the options page when you actually need a bulk export.",
        inputSchema: {
          type: "object",
          properties: {
            folderPath: { type: "string", description: "Folder URI to export (required). Subfolders are NOT recursed -- call once per folder." },
            maxMessages: { type: "integer", description: "Cap on exported messages (default 1000, hard cap 50000). Use sortOrder:'desc' (default) to take the newest N." },
            includeBody: { type: "boolean", description: "If true, extract each message's plain-text body and embed it on the JSON line. Default false. Adds a MIME parse per message so a 5000-message export takes substantially longer." },
            includeAttachmentMeta: { type: "boolean", description: "If true and includeBody is true, embed attachment {name, contentType, size} metadata too. Attachment CONTENT is never written -- only metadata. Default false." },
            scanCap: { type: "integer", description: "Hard cap on messages enumerated before stopping (default 50000). Prevents the LLM from walking a 200k-message archive forever." },
            sortOrder: { type: "string", enum: ["asc", "desc"], description: "Date sort order before applying maxMessages. 'desc' (default) = newest first." },
          },
          required: ["folderPath"],
        },
      },
      {
        name: "getSenderHistory",
        group: "messages", crud: "read",
        title: "Get Sender History",
        description: "Return recent message headers from a given email address, scanned across the inbox of every accessible account. Useful for 'have I corresponded with this person before?' or 'have we already approached this outreach target?' decisions before drafting a reply.",
        inputSchema: {
          type: "object",
          properties: {
            email: { type: "string", description: "Sender email address to match (case-insensitive substring match against the author field)" },
            maxResults: { type: "integer", description: "Cap on returned headers (default 50, max 200)" },
            scanCap: { type: "integer", description: "Hard cap on messages enumerated per folder before stopping (default 5000, max 50000)" },
            sinceDays: { type: "integer", description: "Restrict to messages from the last N days (default unlimited)" },
          },
          required: ["email"],
        },
      },
      {
        name: "searchAttachments",
        group: "messages", crud: "read",
        title: "Search Attachments",
        description: "Find messages with attachments matching a filename substring and/or MIME-type pattern. Walks messages in the chosen folder (or the inbox of every accessible account when folderPath is omitted) and inspects allUserAttachments metadata. Returns header objects with the matching attachment names attached. Capped to avoid mailbox scans.",
        inputSchema: {
          type: "object",
          properties: {
            nameContains: { type: "string", description: "Case-insensitive substring matched against the attachment filename (e.g. 'invoice', '.pdf')" },
            contentType: { type: "string", description: "Case-insensitive prefix matched against the attachment Content-Type (e.g. 'application/pdf', 'image/')" },
            folderPath: { type: "string", description: "Optional folder URI to limit the search. Omitted = scan inbox of each accessible account." },
            maxResults: { type: "integer", description: "Cap on matching messages (default 50, max 200)" },
            scanCap: { type: "integer", description: "Hard cap on messages enumerated per folder before stopping (default 5000, max 50000). Prevents the LLM from accidentally walking a 200k-message archive." },
          },
          required: [],
        },
      },
      {
        name: "searchByThread",
        group: "messages", crud: "read",
        title: "Search By Thread",
        description: "Given any messageId + folderPath, return headers for every other message in the same thread. Uses msgHdr.threadId. Results are headers-only -- call getMessage on a specific id for its body.",
        inputSchema: {
          type: "object",
          properties: {
            messageId: { type: "string", description: "Any message in the thread you want to fetch" },
            folderPath: { type: "string", description: "Folder URI containing the message" },
            maxResults: { type: "integer", description: "Cap on returned headers (default 100, max 500)" },
          },
          required: ["messageId", "folderPath"],
        },
      },
      {
        name: "batchGetMessageHeaders",
        group: "messages", crud: "read",
        title: "Batch Get Message Headers",
        description: "Fetch headers for up to 200 messages in a single round-trip. Returns a map of messageId -> { header object | { error: ... } }. Pair with searchMessages to enrich a result set without N round-trips.",
        inputSchema: {
          type: "object",
          properties: {
            messageIds: { type: "array", items: { type: "string" }, description: "Array of message IDs to fetch headers for (hard cap 200)" },
            folderPath: { type: "string", description: "Folder URI shared by all of the IDs" },
          },
          required: ["messageIds", "folderPath"],
        },
      },
      {
        name: "listTemplates",
        group: "system", crud: "read",
        title: "List Compose Templates",
        description: "List user-authored compose templates stored under <ProfD>/thunderbird-mcp/templates/*.md. Each file has YAML-style frontmatter (---\\nname: short-id\\ndescription: ...\\nsubject: ...\\nisHtml: false\\nvars: [target, contact_name]\\n---) followed by the body. The variable names listed in `vars` are the placeholders renderTemplate accepts.",
        inputSchema: { type: "object", properties: {}, required: [] },
      },
      {
        name: "renderTemplate",
        group: "system", crud: "read",
        title: "Render Compose Template",
        description: "Render a template by name with the supplied variable bindings. Returns { subject, body, isHtml } ready to feed into sendMail. Pure read-only -- nothing is sent. Use this to standardize outreach / reply patterns so the LLM doesn't drift across messages.",
        inputSchema: {
          type: "object",
          properties: {
            name: { type: "string", description: "Template `name` value from its frontmatter" },
            vars: {
              type: "object",
              description: "Variable bindings; keys must match the template's declared `vars` list. Unknown keys are silently ignored; missing required keys produce an error.",
            },
          },
          required: ["name"],
        },
      },
      {
        name: "getServerCapabilities",
        group: "system", crud: "read",
        title: "Get Server Capabilities",
        description: "Return a snapshot of the LLM's sandbox: enabled / disabled tool names, accessible account IDs, current safeguard pref states (blockSkipReview, blockFilterForwardReply, blockContactWrites), current rate-limit bucket states with remaining slots per tool, server version, and audit-log location. Call this once at the start of an agent loop so you know what you can do before you try things that will fail.",
        inputSchema: { type: "object", properties: {}, required: [] },
      },
      {
        name: "getAuditLog",
        group: "system", crud: "read",
        title: "Get Audit Log",
        description: "Read parsed entries from the MCP audit log newest-first. Every outbound compose (sendMail / replyToMessage / forwardMessage / saveDraft) and every contact write are recorded as JSON lines with timestamp + tool + metadata (no body, no attachment content, no recipient addresses beyond counts). Useful for the agent to answer 'what did I send today?' or 'have I already contacted this target?'",
        inputSchema: {
          type: "object",
          properties: {
            maxEntries: { type: "integer", description: "Max entries to return (default 200, hard cap 10000)" },
            tool: { type: "string", description: "Filter by exact tool name (e.g. 'sendMail')" },
            since: { type: "string", description: "Filter to entries on or after this ISO 8601 timestamp" },
            until: { type: "string", description: "Filter to entries on or before this ISO 8601 timestamp" },
          },
          required: [],
        },
      },
      {
        name: "dryRunCompose",
        group: "messages", crud: "read",
        title: "Dry-Run Compose",
        description: "Validate compose parameters WITHOUT sending or saving. Resolves the from identity, parses recipient counts, evaluates every attachment through the same path / size / deny-list pipeline that sendMail uses, and reports the rendered subject (after header sanitization). Useful for an agent that wants to self-check a compose call before triggering the review window or skipReview send.",
        inputSchema: {
          type: "object",
          properties: {
            to: { type: "string", description: "Recipient email address" },
            subject: { type: "string", description: "Email subject line" },
            body: { type: "string", description: "Email body text (length is reported back; content is not echoed)" },
            cc: { type: "string", description: "CC recipients (comma-separated)" },
            bcc: { type: "string", description: "BCC recipients (comma-separated)" },
            isHtml: { type: "boolean", description: "Treat body as HTML (default: false)" },
            from: { type: "string", description: "Sender identity (email address or identity ID)" },
            attachments: {
              type: "array",
              description: "Attachments to evaluate (same shape as sendMail). Each entry is checked but neither read nor copied; only path/size/deny-list status is returned.",
              items: {
                oneOf: [
                  { type: "string", description: "Absolute file path to attach" },
                  {
                    type: "object",
                    properties: {
                      name: { type: "string" },
                      contentType: { type: "string" },
                      base64: { type: "string" },
                    },
                    required: ["name", "base64"],
                    additionalProperties: false,
                  },
                ],
              },
            },
          },
          required: ["to", "subject", "body"],
        },
      },
      {
        name: "sendMail",
        group: "messages", crud: "create",
        title: "Compose Mail",
        description: "Compose a new email. By default opens a compose window for review; set skipReview to send directly.",
        inputSchema: {
          type: "object",
          properties: {
            to: { type: "string", description: "Recipient email address" },
            subject: { type: "string", description: "Email subject line" },
            body: { type: "string", description: "Email body text" },
            cc: { type: "string", description: "CC recipients (comma-separated)" },
            bcc: { type: "string", description: "BCC recipients (comma-separated)" },
            isHtml: { type: "boolean", description: "Set to true if body contains HTML markup (default: false)" },
            from: { type: "string", description: "Sender identity (email address or identity ID from listAccounts)" },
            skipReview: { type: "boolean", description: "If true, send the message directly without opening a compose window (default: false)" },
            idempotencyKey: { type: "string", description: "Optional client-supplied key (max 256 chars). If sendMail was previously called with this same key within the last 24 hours AND succeeded, the prior result is returned instead of sending again. Use this to make retries safe across crashes / network errors -- especially for outreach where re-sending to a real target is costly." },
            attachments: {
              type: "array",
              description: "Attachments: file paths (strings) or inline objects ({name, contentType, base64})",
              items: {
                oneOf: [
                  { type: "string", description: "Absolute file path to attach" },
                  {
                    type: "object",
                    properties: {
                      name: { type: "string", description: "Attachment filename" },
                      contentType: { type: "string", description: "MIME type, e.g. application/pdf" },
                      base64: { type: "string", description: "Base64-encoded file content" },
                    },
                    required: ["name", "base64"],
                    additionalProperties: false,
                  },
                ],
              },
            },
          },
          required: ["to", "subject", "body"],
        },
      },
      {
        name: "saveDraft",
        group: "messages", crud: "create",
        title: "Save Draft",
        description: "Save a composed message to the identity's Drafts folder without sending or opening a compose window. Useful when a human will review and send the message later from Thunderbird.",
        inputSchema: {
          type: "object",
          properties: {
            to: { type: "string", description: "Recipient email address(es), comma-separated. Optional -- a draft can have no recipient." },
            subject: { type: "string", description: "Email subject line (optional)" },
            body: { type: "string", description: "Email body (optional)" },
            cc: { type: "string", description: "CC recipients (comma-separated)" },
            bcc: { type: "string", description: "BCC recipients (comma-separated)" },
            isHtml: { type: "boolean", description: "Set to true if body contains HTML markup (default: false)" },
            from: { type: "string", description: "Sender identity (email address or identity ID from listAccounts)" },
            attachments: {
              type: "array",
              description: "Attachments: file paths (strings) or inline objects ({name, contentType, base64})",
              items: {
                oneOf: [
                  { type: "string", description: "Absolute file path to attach" },
                  {
                    type: "object",
                    properties: {
                      name: { type: "string", description: "Attachment filename" },
                      contentType: { type: "string", description: "MIME type, e.g. application/pdf" },
                      base64: { type: "string", description: "Base64-encoded file content" },
                    },
                    required: ["name", "base64"],
                    additionalProperties: false,
                  },
                ],
              },
            },
          },
          required: [],
        },
      },
      {
        name: "listCalendars",
        group: "calendar", crud: "read",
        title: "List Calendars",
        description: "Return the user's calendars",
        inputSchema: { type: "object", properties: {}, required: [] },
      },
      {
        name: "createEvent",
        group: "calendar", crud: "create",
        title: "Create Event",
        description: "Create a calendar event. By default opens a review dialog; set skipReview to add directly.",
        inputSchema: {
          type: "object",
          properties: {
            title: { type: "string", description: "Event title" },
            startDate: { type: "string", description: "Start date/time in ISO 8601 format" },
            endDate: { type: "string", description: "End date/time in ISO 8601 (defaults to startDate + 1h for timed, +1 day for all-day)" },
            location: { type: "string", description: "Event location" },
            description: { type: "string", description: "Event description" },
            calendarId: { type: "string", description: "Target calendar ID (from listCalendars, defaults to first writable calendar)" },
            allDay: { type: "boolean", description: "Create an all-day event (default: false)" },
            status: { type: "string", description: "VEVENT STATUS: 'tentative', 'confirmed', or 'cancelled'. Defaults to confirmed if omitted." },
            skipReview: { type: "boolean", description: "If true, add the event directly without opening a review dialog (default: false)" },
          },
          required: ["title", "startDate"],
        },
      },
      {
        name: "listEvents",
        group: "calendar", crud: "read",
        title: "List Events",
        description: "List calendar events within a date range",
        inputSchema: {
          type: "object",
          properties: {
            calendarId: { type: "string", description: "Calendar ID to query (from listCalendars). If omitted, queries all calendars." },
            startDate: { type: "string", description: "Start of date range in ISO 8601 format (default: now)" },
            endDate: { type: "string", description: "End of date range in ISO 8601 format (default: 30 days from startDate)" },
            maxResults: { type: "number", description: "Maximum number of events to return (default: 100, max: 500)" },
          },
          required: [],
        },
      },
      {
        name: "updateEvent",
        group: "calendar", crud: "update",
        title: "Update Event",
        description: "Update an existing calendar event's title, dates, location, or description. For recurring events, recurringScope must be supplied explicitly: 'series' edits the entire series, no other scopes are currently supported. The call FAILS for recurring events when recurringScope is omitted, so the agent cannot accidentally rewrite years of past occurrences.",
        inputSchema: {
          type: "object",
          properties: {
            eventId: { type: "string", description: "The event ID (from listEvents results)" },
            calendarId: { type: "string", description: "The calendar ID containing the event (from listEvents results)" },
            title: { type: "string", description: "New event title (optional)" },
            startDate: { type: "string", description: "New start date/time in ISO 8601 format (optional)" },
            endDate: { type: "string", description: "New end date/time in ISO 8601 format (optional)" },
            location: { type: "string", description: "New event location (optional)" },
            description: { type: "string", description: "New event description (optional)" },
            status: { type: "string", description: "New VEVENT STATUS: 'tentative', 'confirmed', or 'cancelled' (optional)" },
            recurringScope: { type: "string", enum: ["series"], description: "Required for recurring events. 'series' rewrites every occurrence past and future. Per-occurrence editing is not supported through this API; use Thunderbird's UI." },
          },
          required: ["eventId", "calendarId"],
        },
      },
      {
        name: "deleteEvent",
        group: "calendar", crud: "delete",
        title: "Delete Event",
        description: "Delete a calendar event. For recurring events, recurringScope must be supplied explicitly: 'series' deletes the entire series. The call FAILS for recurring events when recurringScope is omitted, so the agent cannot accidentally nuke a long-running series.",
        inputSchema: {
          type: "object",
          properties: {
            eventId: { type: "string", description: "The event ID (from listEvents results)" },
            calendarId: { type: "string", description: "The calendar ID containing the event (from listEvents results)" },
            recurringScope: { type: "string", enum: ["series"], description: "Required for recurring events. 'series' deletes every occurrence. Per-occurrence deletion is not supported through this API; use Thunderbird's UI." },
          },
          required: ["eventId", "calendarId"],
        },
      },
      {
        name: "createTask",
        group: "calendar", crud: "create",
        title: "Create Task",
        description: "Open a pre-filled task dialog in Thunderbird for user review before saving, or save directly when skipReview is true.",
        inputSchema: {
          type: "object",
          properties: {
            title: { type: "string", description: "Task title" },
            dueDate: { type: "string", description: "Due date in ISO 8601 format (optional)" },
            calendarId: { type: "string", description: "Target calendar ID (from listCalendars, must have supportsTasks=true)" },
            description: { type: "string", description: "Task description/body (optional)" },
            priority: { type: "integer", description: "Priority: 1=high, 5=normal, 9=low (optional)" },
            categories: { type: "array", items: { type: "string" }, description: "Category labels (optional). Use listCategories to get exact existing names before setting." },
            skipReview: { type: "boolean", description: "If true, save the task directly without opening a review dialog (default: false)" },
          },
          required: ["title"],
        },
      },
      {
        name: "listCategories",
        group: "calendar", crud: "read",
        title: "List Categories",
        description: "Return all calendar category names defined in Thunderbird preferences. Use this before creating tasks or events to get exact category names (case-sensitive).",
        inputSchema: { type: "object", properties: {} },
      },
      {
        name: "listTasks",
        group: "calendar", crud: "read",
        title: "List Tasks",
        description: "List tasks/to-dos from Thunderbird calendars, optionally filtered by completion status or due date",
        inputSchema: {
          type: "object",
          properties: {
            calendarId: { type: "string", description: "Calendar ID to query (from listCalendars). If omitted, queries all task-capable calendars." },
            completed: { type: "boolean", description: "Filter by completion status. true = completed only, false = outstanding only. Omit for all tasks." },
            dueBefore: { type: "string", description: "Return tasks due before this ISO 8601 date" },
            maxResults: { type: "integer", description: "Maximum number of tasks to return (default: 100, max: 500)" },
          },
          required: [],
        },
      },
      {
        name: "updateTask",
        group: "calendar", crud: "update",
        title: "Update Task",
        description: "Update an existing task/to-do: change title, due date, description, priority, completion status, or percent complete",
        inputSchema: {
          type: "object",
          properties: {
            taskId: { type: "string", description: "Task ID (from listTasks results)" },
            calendarId: { type: "string", description: "Calendar ID containing the task (from listTasks results)" },
            title: { type: "string", description: "New task title (optional)" },
            dueDate: { type: "string", description: "New due date in ISO 8601 format (optional)" },
            description: { type: "string", description: "New task description/body (optional)" },
            completed: { type: "boolean", description: "Set to true to mark the task done (sets percentComplete=100 and records completedDate), false to reopen it (optional)" },
            percentComplete: { type: "integer", description: "Completion percentage 0–100 (optional)" },
            priority: { type: "integer", description: "Priority: 1=high, 5=normal, 9=low (optional)" },
          },
          required: ["taskId", "calendarId"],
        },
      },
      {
        name: "searchContacts",
        group: "contacts", crud: "read",
        title: "Search Contacts",
        description: "Search contacts across all address books by email address or name",
        inputSchema: {
          type: "object",
          properties: {
            query: { type: "string", description: "Email address or name to search for" },
            maxResults: { type: "number", description: "Maximum number of results to return (default 50, max 200). If truncated, response includes hasMore: true." },
          },
          required: ["query"],
        },
      },
      {
        name: "createContact",
        group: "contacts", crud: "create",
        title: "Create Contact",
        description: "Create a new contact in an address book",
        inputSchema: {
          type: "object",
          properties: {
            email: { type: "string", description: "Primary email address" },
            displayName: { type: "string", description: "Display name" },
            firstName: { type: "string", description: "First name" },
            lastName: { type: "string", description: "Last name" },
            addressBookId: { type: "string", description: "Address book directory ID (from searchContacts results). Defaults to the first writable address book." },
          },
          required: ["email"],
        },
      },
      {
        name: "updateContact",
        group: "contacts", crud: "update",
        title: "Update Contact",
        description: "Update an existing contact's properties",
        inputSchema: {
          type: "object",
          properties: {
            contactId: { type: "string", description: "Contact UID (from searchContacts results)" },
            email: { type: "string", description: "New primary email address" },
            displayName: { type: "string", description: "New display name" },
            firstName: { type: "string", description: "New first name" },
            lastName: { type: "string", description: "New last name" },
          },
          required: ["contactId"],
        },
      },
      {
        name: "deleteContact",
        group: "contacts", crud: "delete",
        title: "Delete Contact",
        description: "Delete a contact from its address book",
        inputSchema: {
          type: "object",
          properties: {
            contactId: { type: "string", description: "Contact UID (from searchContacts results)" },
          },
          required: ["contactId"],
        },
      },
      {
        name: "replyToMessage",
        group: "messages", crud: "create",
        title: "Reply to Message",
        description: "Reply to a message. By default opens a compose window with quoted original text for review; set skipReview to send directly.",
        inputSchema: {
          type: "object",
          properties: {
            messageId: { type: "string", description: "The message ID to reply to (from searchMessages results)" },
            folderPath: { type: "string", description: "The folder URI path (from searchMessages results)" },
            body: { type: "string", description: "Reply body text" },
            replyAll: { type: "boolean", description: "Reply to all recipients (default: false)" },
            isHtml: { type: "boolean", description: "Set to true if body contains HTML markup (default: false)" },
            to: { type: "string", description: "Override recipient email (default: original sender)" },
            cc: { type: "string", description: "CC recipients (comma-separated)" },
            bcc: { type: "string", description: "BCC recipients (comma-separated)" },
            from: { type: "string", description: "Sender identity (email address or identity ID from listAccounts)" },
            skipReview: { type: "boolean", description: "If true, send the reply directly without opening a compose window (default: false)" },
            idempotencyKey: { type: "string", description: "Optional dedup key (max 256 chars). See sendMail for semantics." },
            attachments: {
              type: "array",
              description: "Attachments: file paths (strings) or inline objects ({name, contentType, base64})",
              items: {
                oneOf: [
                  { type: "string", description: "Absolute file path to attach" },
                  {
                    type: "object",
                    properties: {
                      name: { type: "string", description: "Attachment filename" },
                      contentType: { type: "string", description: "MIME type, e.g. application/pdf" },
                      base64: { type: "string", description: "Base64-encoded file content" },
                    },
                    required: ["name", "base64"],
                    additionalProperties: false,
                  },
                ],
              },
            },
          },
          required: ["messageId", "folderPath", "body"],
        },
      },
      {
        name: "forwardMessage",
        group: "messages", crud: "create",
        title: "Forward Message",
        description: "Forward a message. By default opens a compose window with original content for review; set skipReview to send directly.",
        inputSchema: {
          type: "object",
          properties: {
            messageId: { type: "string", description: "The message ID to forward (from searchMessages results)" },
            folderPath: { type: "string", description: "The folder URI path (from searchMessages results)" },
            to: { type: "string", description: "Recipient email address" },
            body: { type: "string", description: "Additional text to prepend (optional)" },
            isHtml: { type: "boolean", description: "Set to true if body contains HTML markup (default: false)" },
            cc: { type: "string", description: "CC recipients (comma-separated)" },
            bcc: { type: "string", description: "BCC recipients (comma-separated)" },
            from: { type: "string", description: "Sender identity (email address or identity ID from listAccounts)" },
            skipReview: { type: "boolean", description: "If true, send the forward directly without opening a compose window (default: false)" },
            idempotencyKey: { type: "string", description: "Optional dedup key (max 256 chars). See sendMail for semantics." },
            attachments: {
              type: "array",
              description: "Additional attachments: file paths (strings) or inline objects ({name, contentType, base64})",
              items: {
                oneOf: [
                  { type: "string", description: "Absolute file path to attach" },
                  {
                    type: "object",
                    properties: {
                      name: { type: "string", description: "Attachment filename" },
                      contentType: { type: "string", description: "MIME type, e.g. application/pdf" },
                      base64: { type: "string", description: "Base64-encoded file content" },
                    },
                    required: ["name", "base64"],
                    additionalProperties: false,
                  },
                ],
              },
            },
          },
          required: ["messageId", "folderPath", "to"],
        },
      },
      {
        name: "refreshFolder",
        group: "messages", crud: "read",
        title: "Refresh Folder",
        description: "Force an IMAP fetch on the given folder so the server-side state syncs to Thunderbird's local cache. Without this, IMAP folders only refresh when the user clicks them in the UI, which means recently-arrived mail is invisible to searchMessages / getRecentMessages until then. Call refreshFolder before a query when you know new traffic just arrived. Non-IMAP folders return success without doing anything. Note: very large IMAP folders like Gmail [Gmail]/Todos (All Mail) may not finish within the timeout cap; consider refreshing the specific subfolder where the message landed (INBOX, a label) instead.",
        inputSchema: {
          type: "object",
          properties: {
            folderPath: { type: "string", description: "Folder URI (from listFolders) to refresh" },
            timeoutMs: { type: "integer", description: "Max ms to wait for the fetch to complete (default 15000). Hard-capped at 25000 to stay below the stdio bridge's HTTP request timeout (30000ms); larger values are clamped." },
          },
          required: ["folderPath"],
        },
      },
      {
        name: "getRecentMessages",
        group: "messages", crud: "read",
        title: "Get Recent Messages",
        description: "Get recent messages sorted newest-first from a specific folder or all Inboxes, with date and unread filtering",
        inputSchema: {
          type: "object",
          properties: {
            folderPath: { type: "string", description: "Folder URI (from listFolders) to list messages from. If omitted, returns messages from all Inboxes." },
            daysBack: { type: "number", description: "Only return messages from the last N days (default: 7). Use a larger value like 365 for older messages." },
            maxResults: { type: "number", description: "Maximum number of results (default: 50, max: 200)" },
            offset: { type: "number", description: "Number of results to skip for pagination (default 0). When provided, returns {messages, totalMatches, offset, limit, hasMore} instead of a plain array. Note: totalMatches is capped at 10000." },
            unreadOnly: { type: "boolean", description: "Only return unread messages (default: false)" },
            flaggedOnly: { type: "boolean", description: "Only return flagged/starred messages (default: false)" },
            includeSubfolders: { type: "boolean", description: "If false, only return messages from the specified folder — not its subfolders. Default: true." },
          },
          required: [],
        },
      },
      {
        name: "displayMessage",
        group: "messages", crud: "read",
        title: "Display Message",
        description: "Open or navigate to a message in the Thunderbird GUI. Use '3pane' (default) to select the message in the mail view, 'tab' to open in a new tab, or 'window' to open in a standalone window.",
        inputSchema: {
          type: "object",
          properties: {
            messageId: { type: "string", description: "The message ID (from searchMessages results)" },
            folderPath: { type: "string", description: "The folder URI path (from searchMessages results)" },
            displayMode: { type: "string", enum: ["3pane", "tab", "window"], description: "How to display: '3pane' (navigate in mail view, default), 'tab' (new tab), or 'window' (new window)" },
          },
          required: ["messageId", "folderPath"],
        },
      },
      {
        name: "deleteMessages",
        group: "messages", crud: "delete",
        title: "Delete Messages",
        description: "Delete messages from a folder. Drafts are moved to Trash instead of permanently deleted.",
        inputSchema: {
          type: "object",
          properties: {
            messageIds: { type: "array", items: { type: "string" }, description: "Array of message IDs to delete" },
            folderPath: { type: "string", description: "The folder URI containing the messages (from listFolders or searchMessages results)" },
          },
          required: ["messageIds", "folderPath"],
        },
      },
      {
        name: "updateMessage",
        group: "messages", crud: "update",
        title: "Update Message",
        description: "Update one or more messages' read/flagged/tagged state and optionally move them. Supply messageId for a single message or messageIds for bulk operations. Tags are Thunderbird keywords (e.g. '$label1' for Important, '$label2' for Work, or any custom string). Note: combining tags with moveTo/trash on IMAP may not preserve tags on the moved copy — use separate calls if needed.",
        inputSchema: {
          type: "object",
          properties: {
            messageId: { type: "string", description: "A single message ID (from searchMessages results). Required unless messageIds is provided." },
            messageIds: { type: "array", items: { type: "string" }, description: "Array of message IDs for bulk operations. Required unless messageId is provided." },
            folderPath: { type: "string", description: "The folder URI containing the message(s) (from searchMessages results)" },
            read: { type: "boolean", description: "Set to true/false to mark read/unread (optional)" },
            flagged: { type: "boolean", description: "Set to true/false to flag/unflag (optional)" },
            addTags: { type: "array", items: { type: "string" }, description: "Tag keywords to add (e.g. ['$label1', 'project-x']). Thunderbird built-in tags: $label1 (Important), $label2 (Work), $label3 (Personal), $label4 (To Do), $label5 (Later)" },
            removeTags: { type: "array", items: { type: "string" }, description: "Tag keywords to remove from the message(s)" },
            moveTo: { type: "string", description: "Destination folder URI (optional). Cannot be used with trash." },
            trash: { type: "boolean", description: "Set to true to move message to Trash (optional). Cannot be used with moveTo." },
          },
          required: ["folderPath"],
        },
      },
      {
        name: "createFolder",
        group: "folders", crud: "create",
        title: "Create Folder",
        description: "Create a new mail subfolder under an existing folder. Note: on IMAP accounts, server-side completion is asynchronous; verify with listFolders.",
        inputSchema: {
          type: "object",
          properties: {
            parentFolderPath: { type: "string", description: "URI of the parent folder (from listFolders)" },
            name: { type: "string", description: "Name for the new subfolder" },
          },
          required: ["parentFolderPath", "name"],
        },
      },
      {
        name: "renameFolder",
        group: "folders", crud: "update",
        title: "Rename Folder",
        description: "Rename an existing mail folder. Note: on IMAP accounts, server-side completion is asynchronous; verify with listFolders.",
        inputSchema: {
          type: "object",
          properties: {
            folderPath: { type: "string", description: "URI of the folder to rename (from listFolders)" },
            newName: { type: "string", description: "New name for the folder" },
          },
          required: ["folderPath", "newName"],
        },
      },
      {
        name: "deleteFolder",
        group: "folders", crud: "delete",
        title: "Delete Folder",
        description: "Delete a mail folder and all its contents. Moves to Trash, or permanently deletes if already in Trash. Note: permanent deletion may prompt the user for confirmation. On IMAP accounts, server-side completion is asynchronous; verify with listFolders.",
        inputSchema: {
          type: "object",
          properties: {
            folderPath: { type: "string", description: "URI of the folder to delete (from listFolders)" },
          },
          required: ["folderPath"],
        },
      },
      {
        name: "emptyTrash",
        group: "folders", crud: "delete",
        title: "Empty Trash",
        description: "Permanently delete all messages in the Trash folder for an account.",
        inputSchema: {
          type: "object",
          properties: {
            accountId: { type: "string", description: "Account ID (from listAccounts). If omitted, empties Trash for all accessible accounts." },
          },
          required: [],
        },
      },
      {
        name: "emptyJunk",
        group: "folders", crud: "delete",
        title: "Empty Junk",
        description: "Permanently delete all messages in the Junk/Spam folder for an account.",
        inputSchema: {
          type: "object",
          properties: {
            accountId: { type: "string", description: "Account ID (from listAccounts). If omitted, empties Junk for all accessible accounts." },
          },
          required: [],
        },
      },
      {
        name: "moveFolder",
        group: "folders", crud: "update",
        title: "Move Folder",
        description: "Move a mail folder to a new parent folder within the same account. Note: on IMAP accounts, server-side completion is asynchronous; verify with listFolders.",
        inputSchema: {
          type: "object",
          properties: {
            folderPath: { type: "string", description: "URI of the folder to move (from listFolders)" },
            newParentPath: { type: "string", description: "URI of the destination parent folder (from listFolders)" },
          },
          required: ["folderPath", "newParentPath"],
        },
      },
      {
        name: "listFilters",
        group: "filters", crud: "read",
        title: "List Filters",
        description: "List all mail filters/rules for an account with their conditions and actions",
        inputSchema: {
          type: "object",
          properties: {
            accountId: { type: "string", description: "Account ID from listAccounts (omit for all accounts)" },
          },
          required: [],
        },
      },
      {
        name: "createFilter",
        group: "filters", crud: "create",
        title: "Create Filter",
        description: "Create a new mail filter rule on an account",
        inputSchema: {
          type: "object",
          properties: {
            accountId: { type: "string", description: "Account ID" },
            name: { type: "string", description: "Filter name" },
            enabled: { type: "boolean", description: "Whether filter is active (default: true)" },
            type: { type: "number", description: "Filter type bitmask (default: 17 = inbox + manual). 1=inbox, 16=manual, 32=post-plugin, 64=post-outgoing" },
            conditions: {
              type: "array",
              items: {
                type: "object",
                properties: {
                  attrib: { type: "string", description: "Attribute: subject, from, to, cc, toOrCc, body, date, priority, status, size, ageInDays, hasAttachment, junkStatus, tag, otherHeader" },
                  op: { type: "string", description: "Operator: contains, doesntContain, is, isnt, isEmpty, beginsWith, endsWith, isGreaterThan, isLessThan, isBefore, isAfter, matches, doesntMatch" },
                  value: { type: "string", description: "Value to match against" },
                  booleanAnd: { type: "boolean", description: "true=AND with previous, false=OR (default: true)" },
                  header: { type: "string", description: "Custom header name (only when attrib is otherHeader)" },
                },
              },
              description: "Array of filter conditions",
            },
            actions: {
              type: "array",
              items: {
                type: "object",
                properties: {
                  type: { type: "string", description: "Action: moveToFolder, copyToFolder, markRead, markUnread, markFlagged, addTag, changePriority, delete, stopExecution, forward, reply" },
                  value: { type: "string", description: "Action parameter (folder URI for move/copy, tag name for addTag, priority for changePriority, email for forward)" },
                },
              },
              description: "Array of actions to perform",
            },
            insertAtIndex: { type: "number", description: "Position to insert (0 = top priority, default: end of list)" },
          },
          required: ["accountId", "name", "conditions", "actions"],
        },
      },
      {
        name: "updateFilter",
        group: "filters", crud: "update",
        title: "Update Filter",
        description: "Modify an existing filter's properties, conditions, or actions",
        inputSchema: {
          type: "object",
          properties: {
            accountId: { type: "string", description: "Account ID" },
            filterIndex: { type: "number", description: "Filter index (from listFilters)" },
            name: { type: "string", description: "New filter name (optional)" },
            enabled: { type: "boolean", description: "Enable/disable (optional)" },
            type: { type: "number", description: "New filter type bitmask (optional)" },
            conditions: {
              type: "array",
              description: "Replace all conditions (optional, same format as createFilter)",
              items: {
                type: "object",
                properties: {
                  attrib: { type: "string", description: "Attribute: subject, from, to, cc, toOrCc, body, date, priority, status, size, ageInDays, hasAttachment, junkStatus, tag, otherHeader" },
                  op: { type: "string", description: "Operator: contains, doesntContain, is, isnt, isEmpty, beginsWith, endsWith, isGreaterThan, isLessThan, isBefore, isAfter, matches, doesntMatch" },
                  value: { type: "string", description: "Value to match against" },
                  booleanAnd: { type: "boolean", description: "true=AND with previous, false=OR (default: true)" },
                  header: { type: "string", description: "Custom header name (only when attrib is otherHeader)" },
                },
              },
            },
            actions: {
              type: "array",
              description: "Replace all actions (optional, same format as createFilter)",
              items: {
                type: "object",
                properties: {
                  type: { type: "string", description: "Action: moveToFolder, copyToFolder, markRead, markUnread, markFlagged, addTag, changePriority, delete, stopExecution, forward, reply" },
                  value: { type: "string", description: "Action parameter (folder URI for move/copy, tag name for addTag, priority for changePriority, email for forward)" },
                },
              },
            },
          },
          required: ["accountId", "filterIndex"],
        },
      },
      {
        name: "deleteFilter",
        group: "filters", crud: "delete",
        title: "Delete Filter",
        description: "Delete a mail filter by index",
        inputSchema: {
          type: "object",
          properties: {
            accountId: { type: "string", description: "Account ID" },
            filterIndex: { type: "number", description: "Filter index to delete (from listFilters)" },
          },
          required: ["accountId", "filterIndex"],
        },
      },
      {
        name: "reorderFilters",
        group: "filters", crud: "update",
        title: "Reorder Filters",
        description: "Move a filter to a different position in the execution order",
        inputSchema: {
          type: "object",
          properties: {
            accountId: { type: "string", description: "Account ID" },
            fromIndex: { type: "number", description: "Current filter index" },
            toIndex: { type: "number", description: "Target index (0 = highest priority)" },
          },
          required: ["accountId", "fromIndex", "toIndex"],
        },
      },
      {
        name: "applyFilters",
        group: "filters", crud: "update",
        title: "Apply Filters",
        description: "Manually run all enabled filters on a folder to organize existing messages",
        inputSchema: {
          type: "object",
          properties: {
            accountId: { type: "string", description: "Account ID (uses its filters)" },
            folderPath: { type: "string", description: "Folder URI to apply filters to (from listFolders)" },
          },
          required: ["accountId", "folderPath"],
        },
      },
      {
        name: "getAccountAccess",
        group: "system", crud: "read",
        title: "Get Account Access",
        description: "Get the current account access control list. Shows which accounts the MCP server can access. Account access is configured by the user in the extension settings page (Tools > Add-ons > Thunderbird MCP > Options) and cannot be changed via MCP tools.",
        inputSchema: { type: "object", properties: {}, required: [] },
      },
    ];

    // Validate tool metadata: every tool must have valid group and crud fields.
    // This prevents tools from being silently hidden in the settings UI.
    const toolErrors = [];
    for (const tool of tools) {
      if (!tool.group || !VALID_GROUPS.includes(tool.group)) {
        toolErrors.push(`Tool "${tool.name}" has invalid or missing group: "${tool.group}" (valid: ${VALID_GROUPS.join(", ")})`);
      }
      if (!tool.crud || !VALID_CRUD.includes(tool.crud)) {
        toolErrors.push(`Tool "${tool.name}" has invalid or missing crud: "${tool.crud}" (valid: ${VALID_CRUD.join(", ")})`);
      }
    }
    if (toolErrors.length > 0) {
      console.error("thunderbird-mcp: Tool metadata validation failed:\n  " + toolErrors.join("\n  "));
    }

    // Derive ALL_TOOL_NAMES from the tools array (single source of truth)
    const ALL_TOOL_NAMES = tools.map(t => t.name);

    // Group display order for settings UI
    const GROUP_ORDER = { system: 0, messages: 1, folders: 2, contacts: 3, calendar: 4, filters: 5 };
    // Group display labels
    const GROUP_LABELS = { system: "System", messages: "Messages", folders: "Folders", contacts: "Contacts", calendar: "Calendar", filters: "Filters" };

    /**
     * Generate a cryptographically random auth token (hex string).
     * Used to authenticate bridge requests to the HTTP server.
     */
    function generateAuthToken() {
      const bytes = new Uint8Array(32);
      // crypto.getRandomValues is not available in Thunderbird experiment API scope;
      // use the XPCOM random generator instead.
      const rng = Cc["@mozilla.org/security/random-generator;1"]
        .createInstance(Ci.nsIRandomGenerator);
      const randomBytes = rng.generateRandomBytes(32);
      for (let i = 0; i < 32; i++) bytes[i] = randomBytes[i];
      return Array.from(bytes, b => b.toString(16).padStart(2, "0")).join("");
    }

    function getStableAuthTokenPref() {
      try {
        const pref = Services.prefs.getStringPref(PREF_STABLE_AUTH_TOKEN, "");
        const token = pref.trim();
        if (!token) {
          if (pref) {
            console.warn("thunderbird-mcp: stableAuthToken preference is malformed; expected 64 lowercase hex characters, ignoring stored value");
          }
          return "";
        }
        if (!AUTH_TOKEN_PATTERN.test(token)) {
          console.warn("thunderbird-mcp: stableAuthToken preference is malformed; expected 64 lowercase hex characters, ignoring stored value");
          return "";
        }
        return token;
      } catch {
        return "";
      }
    }

    function readConnectionInfo() {
      const tmpDir = Services.dirsvc.get("TmpD", Ci.nsIFile);
      tmpDir.append("thunderbird-mcp");
      const connFile = tmpDir.clone();
      connFile.append("connection.json");
      if (!connFile.exists()) {
        return { path: connFile.path, data: null };
      }
      const fis = Cc["@mozilla.org/network/file-input-stream;1"]
        .createInstance(Ci.nsIFileInputStream);
      fis.init(connFile, 0x01, 0, 0);
      const sis = Cc["@mozilla.org/scriptableinputstream;1"]
        .createInstance(Ci.nsIScriptableInputStream);
      sis.init(fis);
      const text = sis.read(sis.available());
      sis.close();
      return { path: connFile.path, data: JSON.parse(text) };
    }

    return {
      mcpServer: {
        start: async function() {
          // Guard against double-start on extension reload (port conflict)
          if (globalThis.__tbMcpStartPromise) {
            return await globalThis.__tbMcpStartPromise;
          }
          const startPromise = (async () => {
          try {
            // Stop any previously running server (e.g. extension reload)
            if (globalThis.__tbMcpServer) {
              try { globalThis.__tbMcpServer.stop(() => {}); } catch { /* ignore */ }
              globalThis.__tbMcpServer = null;
            }
            const { HttpServer } = ChromeUtils.importESModule(
              "resource://thunderbird-mcp/httpd.sys.mjs?" + Date.now()
            );
            const { NetUtil } = ChromeUtils.importESModule(
              "resource://gre/modules/NetUtil.sys.mjs"
            );
            const { MailServices } = ChromeUtils.importESModule(
              "resource:///modules/MailServices.sys.mjs"
            );

            let cal = null;
            let CalEvent = null;
            let CalTodo = null;
            try {
              const calModule = ChromeUtils.importESModule(
                "resource:///modules/calendar/calUtils.sys.mjs"
              );
              cal = calModule.cal;
              const { CalEvent: CE } = ChromeUtils.importESModule(
                "resource:///modules/CalEvent.sys.mjs"
              );
              CalEvent = CE;
              const { CalTodo: CT } = ChromeUtils.importESModule(
                "resource:///modules/CalTodo.sys.mjs"
              );
              CalTodo = CT;
            } catch {
              // Calendar not available
            }

            let GlodaMsgSearcher = null;
            try {
              const glodaModule = ChromeUtils.importESModule(
                "resource:///modules/gloda/GlodaMsgSearcher.sys.mjs"
              );
              GlodaMsgSearcher = glodaModule.GlodaMsgSearcher;
            } catch {
              // Gloda not available
            }

            /**
             * CRITICAL: Must specify { charset: "UTF-8" } or emojis/special chars
             * will be corrupted. NetUtil defaults to Latin-1.
             */
            function readRequestBody(request) {
              const stream = request.bodyInputStream;
              return NetUtil.readInputStreamToString(stream, stream.available(), { charset: "UTF-8" });
            }

            /**
             * Apply offset-based pagination to a sorted results array.
             * Removes the internal _dateTs property from each result.
             *
             * Backward-compatible: when offset is undefined/null (not provided),
             * returns a plain array. When offset is explicitly provided (even 0),
             * returns structured { messages, totalMatches, offset, limit, hasMore }.
             * Note: totalMatches is capped at SEARCH_COLLECTION_CAP and may underreport.
             */
            function paginate(results, offset, effectiveLimit) {
              const offsetProvided = offset !== undefined && offset !== null;
              const effectiveOffset = (offset > 0) ? Math.floor(offset) : 0;
              const page = results.slice(effectiveOffset, effectiveOffset + effectiveLimit).map(r => {
                delete r._dateTs;
                return r;
              });
              if (!offsetProvided) {
                return page;
              }
              return {
                messages: page,
                totalMatches: results.length,
                offset: effectiveOffset,
                limit: effectiveLimit,
                hasMore: effectiveOffset + effectiveLimit < results.length
              };
            }

            /**
             * Write connection info (port + auth token) to a well-known file
             * so the bridge can discover how to connect.
             * File: <TmpD>/thunderbird-mcp/connection.json
             */
            function writeConnectionInfo(port, token) {
              const tmpDir = Services.dirsvc.get("TmpD", Ci.nsIFile);
              tmpDir.append("thunderbird-mcp");
              if (!tmpDir.exists()) {
                tmpDir.create(Ci.nsIFile.DIRECTORY_TYPE, 0o700);
              } else if (tmpDir.isSymlink()) {
                throw new Error("thunderbird-mcp tmp directory is a symlink — refusing to write connection info");
              } else {
                // POSIX hardening: on a shared /tmp another local user could
                // pre-create the directory with group/world bits set, then race
                // the connection file. The O_EXCL on the file itself blocks a
                // straight overwrite, but a permissive directory still lets the
                // attacker read or rename our file. Force perms back to 0o700.
                //
                // Windows uses ACLs, not POSIX modes. The bits reported by
                // nsIFile.permissions on Windows do not correspond to the
                // group/world semantics this check assumes -- a normal Temp
                // subfolder reads as 0o666 or similar and trips a false
                // positive. Skip the chmod entirely on Windows; the POSIX
                // attack model (shared /tmp other-user race) does not apply
                // there anyway since %LOCALAPPDATA%\Temp is per-user.
                const isWindows = (() => {
                  try { return Services.appinfo.OS === "WINNT"; } catch { return false; }
                })();
                if (!isWindows) {
                  try {
                    const mode = tmpDir.permissions;
                    if (mode && (mode & 0o077) !== 0) {
                      try { tmpDir.permissions = 0o700; } catch { /* best-effort */ }
                      if ((tmpDir.permissions & 0o077) !== 0) {
                        throw new Error("thunderbird-mcp tmp directory has group/world permissions — refusing to write connection info");
                      }
                    }
                  } catch (e) {
                    if (e && e.message && e.message.startsWith("thunderbird-mcp tmp directory")) throw e;
                    // ignore: permissions accessor unsupported on this platform
                  }
                }
              }
              const connFile = tmpDir.clone();
              connFile.append("connection.json");
              // Symlink defense: remove any existing file first, then create
              // with O_CREAT|O_EXCL (0x08|0x80) to fail if a symlink appeared
              // between remove and create.
              if (connFile.exists()) {
                connFile.remove(false);
              }
              const data = JSON.stringify({ port, token, pid: Services.appinfo.processID });
              const ostream = Cc["@mozilla.org/network/file-output-stream;1"]
                .createInstance(Ci.nsIFileOutputStream);
              // 0x02 = O_WRONLY, 0x08 = O_CREAT, 0x80 = O_EXCL
              ostream.init(connFile, 0x02 | 0x08 | 0x80, 0o600, 0);
              const converter = Cc["@mozilla.org/intl/converter-output-stream;1"]
                .createInstance(Ci.nsIConverterOutputStream);
              converter.init(ostream, "UTF-8");
              converter.writeString(data);
              converter.close();
              return connFile.path;
            }

            /**
             * Remove the connection info file on shutdown.
             */
            function removeConnectionInfo() {
              try {
                const tmpDir = Services.dirsvc.get("TmpD", Ci.nsIFile);
                tmpDir.append("thunderbird-mcp");
                const connFile = tmpDir.clone();
                connFile.append("connection.json");
                if (connFile.exists()) {
                  connFile.remove(false);
                }
              } catch {
                // Best-effort cleanup
              }
            }

            const authToken = getStableAuthTokenPref() || generateAuthToken();

            /**
             * Constant-time string comparison to prevent timing side-channel attacks.
             */
            function timingSafeEqual(a, b) {
              const aStr = String(a);
              const bStr = String(b);
              const len = Math.max(aStr.length, bStr.length);
              let result = aStr.length ^ bStr.length;
              for (let i = 0; i < len; i++) {
                result |= (aStr.charCodeAt(i) || 0) ^ (bStr.charCodeAt(i) || 0);
              }
              return result === 0;
            }

            // Audit-log cap before rotation. 5 MB of JSON lines is roughly
            // 10k-30k compose events depending on subject length; rotating to
            // a single `.log.1` keeps disk use bounded without losing history.
            const AUDIT_LOG_ROTATE_BYTES = 5 * 1024 * 1024;
            const AUDIT_LOG_SUBDIR = "thunderbird-mcp";
            const AUDIT_LOG_FILENAME = "audit.log";
            const AUDIT_LOG_ROTATED_FILENAME = "audit.log.1";

            /**
             * Append a single JSON line describing an outbound-compose action
             * (sendMail / replyToMessage / forwardMessage / saveDraft) to
             * <ProfD>/thunderbird-mcp/audit.log. Best-effort: any failure is
             * swallowed so disk errors never block a legitimate send.
             *
             * Logged fields are metadata only -- no body, no attachment
             * content, no recipient lists beyond counts. The whole point is
             * incident response, not message archival.
             */
            function appendComposeAudit(entry) {
              try {
                const profDir = Services.dirsvc.get("ProfD", Ci.nsIFile);
                const auditDir = profDir.clone();
                auditDir.append(AUDIT_LOG_SUBDIR);
                if (!auditDir.exists()) {
                  auditDir.create(Ci.nsIFile.DIRECTORY_TYPE, 0o700);
                }

                const logFile = auditDir.clone();
                logFile.append(AUDIT_LOG_FILENAME);

                // Rotate when the active log exceeds the cap.
                if (logFile.exists()) {
                  let size = 0;
                  try { size = logFile.fileSize; } catch { size = 0; }
                  if (size > AUDIT_LOG_ROTATE_BYTES) {
                    const rotated = auditDir.clone();
                    rotated.append(AUDIT_LOG_ROTATED_FILENAME);
                    if (rotated.exists()) {
                      try { rotated.remove(false); } catch { /* best-effort */ }
                    }
                    try { logFile.moveTo(auditDir, AUDIT_LOG_ROTATED_FILENAME); } catch { /* best-effort */ }
                  }
                }

                const line = JSON.stringify({
                  ts: new Date().toISOString(),
                  ...entry,
                }) + "\n";

                const ostream = Cc["@mozilla.org/network/file-output-stream;1"]
                  .createInstance(Ci.nsIFileOutputStream);
                // 0x02 = O_WRONLY, 0x08 = O_CREAT, 0x10 = O_APPEND
                ostream.init(logFile, 0x02 | 0x08 | 0x10, 0o600, 0);
                const converter = Cc["@mozilla.org/intl/converter-output-stream;1"]
                  .createInstance(Ci.nsIConverterOutputStream);
                converter.init(ostream, "UTF-8");
                converter.writeString(line);
                converter.close();
              } catch (e) {
                // Audit failure must never block a send. Surface it once on the
                // console and move on; do NOT propagate.
                try { console.warn("thunderbird-mcp: audit log write failed:", e); } catch { /* ignore */ }
              }
            }

            // summarizeAttachmentsForAudit + countRecipients live in
            // security_helpers.js; loaded above and destructured into scope.

            /**
             * Read the audit log file(s) and return parsed JSON entries newest-
             * first. Reads both audit.log and audit.log.1 if rotation happened
             * so the caller sees the full bounded history.
             *
             * Returns { entries, totalScanned, truncated, errors }:
             *   - entries: array of parsed log objects (most recent first)
             *   - totalScanned: number of bytes read
             *   - truncated: true if maxEntries cut the list
             *   - errors: per-line parse failures (object with line + reason)
             *
             * filter is optional: { tool, since, until } where since/until are
             * ISO timestamps and tool is an exact name match.
             */
            function readAuditLog(maxEntries, filter) {
              const limit = Number.isFinite(maxEntries) && maxEntries > 0
                ? Math.min(Math.floor(maxEntries), 10000)
                : 500;
              const wantTool = filter && typeof filter.tool === "string" ? filter.tool : null;
              const sinceTs = filter && filter.since ? Date.parse(filter.since) : null;
              const untilTs = filter && filter.until ? Date.parse(filter.until) : null;

              const out = { entries: [], totalScanned: 0, truncated: false, errors: [] };
              const profDir = Services.dirsvc.get("ProfD", Ci.nsIFile);
              const auditDir = profDir.clone();
              auditDir.append(AUDIT_LOG_SUBDIR);
              if (!auditDir.exists()) return out;

              // Read the active log first, then the rotated one. Newest-first
              // means we read each file end-to-end, parse, reverse the file's
              // entries, then append. The rotated file (.log.1) is older.
              const fileNames = [AUDIT_LOG_FILENAME, AUDIT_LOG_ROTATED_FILENAME];
              for (const fname of fileNames) {
                if (out.entries.length >= limit) {
                  out.truncated = true;
                  break;
                }
                const f = auditDir.clone();
                f.append(fname);
                if (!f.exists()) continue;
                let text = "";
                try {
                  const fis = Cc["@mozilla.org/network/file-input-stream;1"].createInstance(Ci.nsIFileInputStream);
                  fis.init(f, 0x01, 0, 0);
                  const cis = Cc["@mozilla.org/intl/converter-input-stream;1"].createInstance(Ci.nsIConverterInputStream);
                  cis.init(fis, "UTF-8", 0, 0);
                  const buf = {};
                  let read;
                  while ((read = cis.readString(65536, buf)) > 0) {
                    text += buf.value;
                  }
                  cis.close();
                  out.totalScanned += text.length;
                } catch (e) {
                  out.errors.push({ file: fname, reason: String(e) });
                  continue;
                }
                const lines = text.split("\n");
                // Reverse so newest-first; skip empty trailing line.
                for (let i = lines.length - 1; i >= 0; i--) {
                  const line = lines[i];
                  if (!line) continue;
                  let parsed;
                  try { parsed = JSON.parse(line); }
                  catch { out.errors.push({ line: line.slice(0, 80), reason: "parse failure" }); continue; }
                  if (wantTool && parsed.tool !== wantTool) continue;
                  if (sinceTs && parsed.ts && Date.parse(parsed.ts) < sinceTs) continue;
                  if (untilTs && parsed.ts && Date.parse(parsed.ts) > untilTs) continue;
                  out.entries.push(parsed);
                  if (out.entries.length >= limit) {
                    out.truncated = true;
                    break;
                  }
                }
              }
              return out;
            }

            // Idempotency window: how far back to scan for a matching key
            // before considering the new call a fresh send.
            const IDEMPOTENCY_WINDOW_HOURS = 24;

            /**
             * Find a recent successful audit entry whose idempotencyKey
             * matches `key`. Returns the entry's stored `result` object, or
             * null if no match. Used by sendMail / replyToMessage /
             * forwardMessage to skip duplicate sends after a crash or retry.
             *
             * Only entries with .success === true are considered hits --
             * a previous error MUST allow the caller to retry.
             */
            function findIdempotentEntry(tool, key) {
              if (typeof key !== "string" || !key) return null;
              if (key.length > 256) return null; // schema caps caller input
              const since = new Date(Date.now() - IDEMPOTENCY_WINDOW_HOURS * 3600 * 1000).toISOString();
              const log = readAuditLog(1000, { tool, since });
              for (const e of log.entries) {
                if (e && e.idempotencyKey === key && e.success === true && e.result) {
                  return e.result;
                }
              }
              return null;
            }

            /**
             * Truncate both audit.log and audit.log.1. Returns the number of
             * bytes deleted. Best-effort; missing files are silent successes.
             */
            function clearAuditLog() {
              let bytesRemoved = 0;
              try {
                const profDir = Services.dirsvc.get("ProfD", Ci.nsIFile);
                const auditDir = profDir.clone();
                auditDir.append(AUDIT_LOG_SUBDIR);
                if (!auditDir.exists()) return { success: true, bytesRemoved: 0 };
                for (const fname of [AUDIT_LOG_FILENAME, AUDIT_LOG_ROTATED_FILENAME]) {
                  const f = auditDir.clone();
                  f.append(fname);
                  if (f.exists()) {
                    try { bytesRemoved += f.fileSize; } catch { /* ignore */ }
                    try { f.remove(false); } catch { /* ignore */ }
                  }
                }
              } catch (e) {
                return { error: String(e), bytesRemoved };
              }
              return { success: true, bytesRemoved };
            }

            // Pref-read cache. Services.prefs is hit on every tool dispatch
            // (rate-limit, access-control, safeguards) and the cost adds up
            // under search bursts. We cache the parsed value behind each
            // pref name and register one observer per pref that flips the
            // cached entry to undefined on change, forcing a re-read next
            // call. Cheap: a few microseconds saved per call, but matters
            // for batch tools like batchGetMessageHeaders.
            const __prefCache = Object.create(null);
            const __prefObservers = Object.create(null);

            function __invalidatePrefCache(prefName) {
              return {
                observe(subject, topic, data) {
                  if (topic === "nsPref:changed" && data === prefName) {
                    delete __prefCache[prefName];
                  }
                },
              };
            }
            function __ensurePrefObserver(prefName) {
              if (__prefObservers[prefName]) return;
              const obs = __invalidatePrefCache(prefName);
              try {
                Services.prefs.addObserver(prefName, obs, false);
                __prefObservers[prefName] = obs;
              } catch (e) {
                console.warn("thunderbird-mcp: pref observer registration failed for", prefName, e);
              }
            }
            function __cachedRead(prefName, reader) {
              __ensurePrefObserver(prefName);
              if (__prefCache[prefName] !== undefined) return __prefCache[prefName];
              const value = reader();
              __prefCache[prefName] = value;
              return value;
            }

            /**
             * Get the list of allowed account IDs from preferences.
             * Returns an empty array if no restriction is set (all accounts allowed).
             */
            function getAllowedAccountIds() {
              return __cachedRead(PREF_ALLOWED_ACCOUNTS, () => {
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
                  // Fail closed: corrupt pref means block all accounts, not allow all
                  console.error("thunderbird-mcp: failed to parse allowed accounts pref, blocking all accounts:", e);
                  return ["__invalid__"];
                }
              });
            }

            /**
             * Check if an account is accessible based on the allowed accounts list.
             * When the list is empty, all accounts are accessible (default).
             */
            function isAccountAllowed(accountKey) {
              const allowed = getAllowedAccountIds();
              if (allowed.length === 0) return true;
              return allowed.includes(accountKey);
            }

            /**
             * Check if the user has disabled the skipReview shortcut.
             * When true, send/reply/forward/createEvent/createTask tools must open
             * the review window/dialog even if the caller passed skipReview: true.
             *
             * Default is true: an LLM that reads attacker-controlled email content
             * can be prompt-injected into invoking sendMail with skipReview, so the
             * safe default is to require human review. Users can explicitly opt
             * into silent sends from the options page.
             */
            function isSkipReviewBlocked() {
              return __cachedRead(PREF_BLOCK_SKIPREVIEW, () => {
                try { return Services.prefs.getBoolPref(PREF_BLOCK_SKIPREVIEW, true); }
                catch { return true; }
              });
            }

            /**
             * Check if the user has blocked `forward` and `reply` filter actions
             * from the MCP API. Filters execute on every incoming message with
             * no UI, so allowing an MCP caller to create one is equivalent to
             * giving the LLM write access to a persistent silent-exfil channel.
             * Default true.
             */
            function isFilterForwardReplyBlocked() {
              return __cachedRead(PREF_BLOCK_FILTER_FORWARD_REPLY, () => {
                try { return Services.prefs.getBoolPref(PREF_BLOCK_FILTER_FORWARD_REPLY, true); }
                catch { return true; }
              });
            }

            /**
             * Check if the user has blocked address-book write operations from
             * the MCP API (createContact / updateContact / deleteContact).
             * Contact writes are persistent, cross every configured address
             * book, and have no UI confirmation -- enabling a "spoof a known
             * sender" attack where the LLM edits Boss's email to attacker's.
             * Default true.
             */
            function isContactWritesBlocked() {
              return __cachedRead(PREF_BLOCK_CONTACT_WRITES, () => {
                try { return Services.prefs.getBoolPref(PREF_BLOCK_CONTACT_WRITES, true); }
                catch { return true; }
              });
            }

            /**
             * Check if bulk mailbox export is blocked. Default true:
             * exportMailbox writes the user's mail content to disk in a
             * machine-readable format, which is an attractive primitive for
             * an LLM that has been prompt-injected into "back up everything
             * I have". Off-by-pref keeps it from being a one-call data
             * exfil channel.
             */
            function isMailboxExportBlocked() {
              return __cachedRead(PREF_BLOCK_MAILBOX_EXPORT, () => {
                try { return Services.prefs.getBoolPref(PREF_BLOCK_MAILBOX_EXPORT, true); }
                catch { return true; }
              });
            }

            /**
             * Get the list of disabled tool names from preferences.
             * Returns an empty array if no tools are disabled (all enabled).
             * Fails closed: corrupt pref disables all tools.
             */
            function getDisabledTools() {
              return __cachedRead(PREF_DISABLED_TOOLS, () => {
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
              });
            }

            /**
             * Check if a tool is enabled.
             * Undisableable tools (listAccounts, listFolders, getAccountAccess) always return true.
             */
            function isToolEnabled(toolName) {
              if (UNDISABLEABLE_TOOLS.has(toolName)) return true;
              const disabled = getDisabledTools();
              if (disabled.includes("__all__")) return false;
              return !disabled.includes(toolName);
            }

            /**
             * Check if a resolved folder belongs to an allowed account.
             * Returns true if the folder's account is accessible, false otherwise.
             */
            function isFolderAccessible(folder) {
              if (!folder || !folder.server) return false;
              const account = MailServices.accounts.findAccountForServer(folder.server);
              return account ? isAccountAllowed(account.key) : false;
            }

            /**
             * Lookup a folder by URI and verify it exists and is accessible.
             * Returns { folder } on success, or { error } if not found or restricted.
             */
            function getAccessibleFolder(folderPath) {
              const folder = MailServices.folderLookup.getFolderForURL(folderPath);
              if (!folder) return { error: `Folder not found: ${folderPath}` };
              if (!isFolderAccessible(folder)) return { error: `Account not accessible for folder: ${folderPath}` };
              return { folder };
            }

            /**
             * Get all accessible Thunderbird accounts, filtered by allowed list.
             */
            function getAccessibleAccounts() {
              const result = [];
              for (const account of MailServices.accounts.accounts) {
                if (isAccountAllowed(account.key)) {
                  result.push(account);
                }
              }
              return result;
            }

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
             * Get the current account access control list.
             */
            function getAccountAccess() {
              const allowed = getAllowedAccountIds();
              // Only return accessible accounts — restricted accounts are hidden
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

            /**
             * Lists all folders (optionally limited to a single account).
             * Depth is 0 for root children, increasing for subfolders.
             */
            function listFolders(accountId, folderPath) {
              const results = [];

              function folderType(flags) {
                if (flags & 0x00001000) return "inbox";
                if (flags & 0x00000200) return "sent";
                if (flags & 0x00000400) return "drafts";
                if (flags & 0x00000100) return "trash";
                if (flags & 0x00400000) return "templates";
                if (flags & 0x00000800) return "queue";
                if (flags & 0x40000000) return "junk";
                if (flags & 0x00004000) return "archive";
                return "folder";
              }

              function walkFolder(folder, accountKey, depth) {
                try {
                  // Skip virtual/search folders to avoid duplicates
                  if (folder.flags & 0x00000020) return;

                  const prettyName = folder.prettyName;
                  results.push({
                    name: prettyName || folder.name || "(unnamed)",
                    path: folder.URI,
                    type: folderType(folder.flags),
                    accountId: accountKey,
                    totalMessages: folder.getTotalMessages(false),
                    unreadMessages: folder.getNumUnread(false),
                    depth
                  });
                } catch {
                  // Skip inaccessible folders
                }

                try {
                  if (folder.hasSubFolders) {
                    for (const subfolder of folder.subFolders) {
                      walkFolder(subfolder, accountKey, depth + 1);
                    }
                  }
                } catch {
                  // Skip subfolder traversal errors
                }
              }

              // folderPath filter: list that folder and its subtree
              if (folderPath) {
                const result = getAccessibleFolder(folderPath);
                if (result.error) return result;
                const folder = result.folder;
                const accountKey = folder.server
                  ? (MailServices.accounts.findAccountForServer(folder.server)?.key || "unknown")
                  : "unknown";
                walkFolder(folder, accountKey, 0);
                return results;
              }

              if (accountId) {
                if (!isAccountAllowed(accountId)) {
                  return { error: `Account not accessible: ${accountId}` };
                }
                let target = null;
                for (const account of MailServices.accounts.accounts) {
                  if (account.key === accountId) {
                    target = account;
                    break;
                  }
                }
                if (!target) {
                  return { error: `Account not found: ${accountId}` };
                }
                try {
                  const root = target.incomingServer.rootFolder;
                  if (root && root.hasSubFolders) {
                    for (const subfolder of root.subFolders) {
                      walkFolder(subfolder, target.key, 0);
                    }
                  }
                } catch {
                  // Skip inaccessible account
                }
                return results;
              }

              for (const account of getAccessibleAccounts()) {
                try {
                  const root = account.incomingServer.rootFolder;
                  if (!root) continue;
                  if (root.hasSubFolders) {
                    for (const subfolder of root.subFolders) {
                      walkFolder(subfolder, account.key, 0);
                    }
                  }
                } catch {
                  // Skip inaccessible accounts/folders
                }
              }

              return results;
            }

            /** Returns user-visible tag keywords from a message header, filtering out internal IMAP flags. */
            function getUserTags(msgHdr) {
              return (msgHdr.getStringProperty("keywords") || "").split(/\s+/).filter(k => k && !INTERNAL_KEYWORDS.has(k.toLowerCase()));
            }

            function stripHtml(html) {
              if (!html) return "";
              let text = String(html);

              // Remove style/script blocks
              text = text.replace(/<script\b[^>]*>[\s\S]*?<\/script>/gi, " ");
              text = text.replace(/<style\b[^>]*>[\s\S]*?<\/style>/gi, " ");

              // Convert block-level tags to newlines before stripping
              text = text.replace(/<br\s*\/?>/gi, "\n");
              text = text.replace(/<\/(p|div|li|tr|h[1-6]|blockquote|pre)>/gi, "\n");
              text = text.replace(/<(p|div|li|tr|h[1-6]|blockquote|pre)\b[^>]*>/gi, "\n");

              // Strip remaining tags
              text = text.replace(/<[^>]+>/g, " ");

              // Decode entities in a single pass
              const NAMED_ENTITIES = {
                nbsp: " ", amp: "&", lt: "<", gt: ">", quot: "\"", apos: "'",
                "#39": "'",
                mdash: "\u2014", ndash: "\u2013", hellip: "\u2026",
                lsquo: "\u2018", rsquo: "\u2019", ldquo: "\u201C", rdquo: "\u201D",
                bull: "\u2022", middot: "\u00B7", ensp: "\u2002", emsp: "\u2003",
                thinsp: "\u2009", zwnj: "\u200C", zwj: "\u200D",
                laquo: "\u00AB", raquo: "\u00BB",
                copy: "\u00A9", reg: "\u00AE", trade: "\u2122", deg: "\u00B0",
                plusmn: "\u00B1", times: "\u00D7", divide: "\u00F7",
                micro: "\u00B5", para: "\u00B6", sect: "\u00A7",
                euro: "\u20AC", pound: "\u00A3", yen: "\u00A5", cent: "\u00A2",
              };
              text = text.replace(/&(#x?[0-9a-fA-F]+|[a-zA-Z]+);/gi, (match, entity) => {
                if (entity.startsWith("#x") || entity.startsWith("#X")) {
                  const cp = parseInt(entity.slice(2), 16);
                  if (!cp || cp > 0x10FFFF) return match;
                  try { return String.fromCodePoint(cp); } catch { return match; }
                }
                if (entity.startsWith("#")) {
                  const cp = parseInt(entity.slice(1), 10);
                  if (!cp || cp > 0x10FFFF) return match;
                  try { return String.fromCodePoint(cp); } catch { return match; }
                }
                return NAMED_ENTITIES[entity.toLowerCase()] || match;
              });

              // Normalize newlines/spaces
              text = text.replace(/\r\n/g, "\n").replace(/\r/g, "\n");
              text = text.replace(/\n{3,}/g, "\n\n");
              text = text.replace(/[ \t\f\v]+/g, " ");
              text = text.replace(/ *\n */g, "\n");
              text = text.trim();
              return text;
            }

            // SAFE_HREF_SCHEMES lives in security_helpers.js.

            // isSafeMarkdownHref / isSafeImageSrc / escapeMarkdownLinkText /
            // renderMarkdownLink live in security_helpers.js.

            /**
             * Walks the MIME tree to find the raw body content.
             * Returns { text, isHtml } without any format conversion.
             * Does NOT use coerceBodyToPlaintext -- callers that want
             * the raw HTML (for markdown/html output) need this.
             */
            function extractBodyContent(aMimeMsg) {
              if (!aMimeMsg) return { text: "", isHtml: false };
              try {
                function findBody(part, isRoot = false) {
                  const ct = ((part.contentType || "").split(";")[0] || "").trim().toLowerCase();
                  if (ct === "message/rfc822" && !isRoot) return null;
                  if (ct !== "message/rfc822") {
                    if (ct === "text/plain" && part.body) return { text: part.body, isHtml: false };
                    if (ct === "text/html" && part.body) return { text: part.body, isHtml: true };
                  }
                  if (part.parts) {
                    let htmlFallback = null;
                    for (const sub of part.parts) {
                      const r = findBody(sub);
                      if (r && !r.isHtml) return r;
                      if (r && r.isHtml && !htmlFallback) htmlFallback = r;
                    }
                    if (htmlFallback) return htmlFallback;
                  }
                  return null;
                }
                const found = findBody(aMimeMsg, true);
                if (found) return found;
              } catch { /* give up */ }
              return { text: "", isHtml: false };
            }

            /**
             * Extracts plain text body from a MIME message.
             * Uses coerceBodyToPlaintext as fast path, then MIME tree fallback.
             * Used by reply/forward quoting where plain text is appropriate.
             */
            function extractPlainTextBody(aMimeMsg) {
              if (!aMimeMsg) return "";
              try {
                const text = aMimeMsg.coerceBodyToPlaintext();
                if (text) return text;
              } catch { /* fall through */ }
              const { text, isHtml } = extractBodyContent(aMimeMsg);
              return isHtml ? stripHtml(text) : text;
            }

            // UNTRUSTED_OPEN/CLOSE + wrapUntrustedBody/Preview live in security_helpers.js.

	            /**
	             * Opens a folder and its message database.
	             * Best-effort refresh for IMAP folders (db may be stale).
	             * Returns { folder, db } or { error }.
	             */
	            function openFolder(folderPath) {
	              try {
	                const result = getAccessibleFolder(folderPath);
	                if (result.error) return result;
	                const folder = result.folder;

	                // Attempt to refresh IMAP folders. This is async and may not
	                // complete before we read, but helps with stale data.
	                if (folder.server && folder.server.type === "imap") {
	                  try {
	                    folder.updateFolder(null);
	                  } catch {
	                    // updateFolder may fail, continue anyway
	                  }
	                }

	                const db = folder.msgDatabase;
	                if (!db) {
	                  return { error: "Could not access folder database" };
	                }

	                return { folder, db };
	              } catch (e) {
	                return { error: e.toString() };
	              }
	            }

	            // Hard cap on the linear-enumeration fallback when
	            // getMsgHdrForMessageID misses. A 200k-message archive walked
	            // header-by-header takes tens of seconds; we'd rather fail
	            // fast and tell the caller to pass a smaller folderPath than
	            // burn an agent's clock on a single lookup.
	            const FIND_MESSAGE_SCAN_CAP = 50000;

	            function findMessage(messageId, folderPath) {
	              const opened = openFolder(folderPath);
	              if (opened.error) return opened;

	              const { folder, db } = opened;
	              let msgHdr = null;

	              const hasDirectLookup = typeof db.getMsgHdrForMessageID === "function";
	              if (hasDirectLookup) {
	                try {
	                  msgHdr = db.getMsgHdrForMessageID(messageId);
	                } catch {
	                  msgHdr = null;
	                }
	              }

	              if (!msgHdr) {
	                let scanned = 0;
	                let capped = false;
	                for (const hdr of db.enumerateMessages()) {
	                  scanned++;
	                  if (scanned > FIND_MESSAGE_SCAN_CAP) { capped = true; break; }
	                  if (hdr.messageId === messageId) {
	                    msgHdr = hdr;
	                    break;
	                  }
	                }
	                if (!msgHdr && capped) {
	                  return {
	                    error: `Message not found after scanning ${FIND_MESSAGE_SCAN_CAP} headers in this folder. Pass a more specific folderPath or use searchMessages first to locate the exact folder.`,
	                  };
	                }
	              }

	              if (!msgHdr) {
	                return { error: `Message not found: ${messageId}` };
	              }

	              return { msgHdr, folder, db };
	            }

            /**
             * Validate compose parameters without sending. Mirrors the
             * argument shape of composeMail and reports what WOULD happen:
             * resolved identity, recipient counts, attachment status per
             * entry (resolved size, sensitive-path verdict, oversize
             * verdict, success/failure), rendered subject after header
             * sanitization, and the skipReview-block pref state.
             *
             * Pure read-only -- nothing is queued, drafted, or sent. Useful
             * for an agent that wants to self-check a compose call before
             * committing to it.
             */

            /**
             * Snapshot of the agent's sandbox: what tools are available, what
             * accounts are reachable, what gates are on, how much rate-limit
             * budget remains. Designed for the LLM to call ONCE at the start
             * of a session to plan around constraints rather than discovering
             * them via failed calls.
             */
            function getServerCapabilities() {
              const allTools = tools.map(t => t.name);
              const enabled = allTools.filter(n => isToolEnabled(n));
              const disabled = allTools.filter(n => !isToolEnabled(n));
              const accessibleAccounts = getAccessibleAccounts().map(a => ({
                id: a.key,
                name: a.incomingServer && a.incomingServer.prettyName ? a.incomingServer.prettyName : a.key,
              }));
              const safeguards = {
                blockSkipReview: isSkipReviewBlocked(),
                blockFilterForwardReply: isFilterForwardReplyBlocked(),
                blockContactWrites: isContactWritesBlocked(),
              };
              return {
                serverVersion: getExtVersion(),
                tools: { enabled, disabled, total: allTools.length },
                accessibleAccounts,
                accessMode: getAllowedAccountIds().length === 0 ? "all" : "restricted",
                safeguards,
                rateLimits: inspectRateLimits(__rateLimiterState),
                auditLogPath: `<ProfD>/${AUDIT_LOG_SUBDIR}/${AUDIT_LOG_FILENAME}`,
              };
            }

            /**
             * Build a lookup from tool name to inputSchema for fast validation.
             */
            const toolSchemas = Object.create(null);
            for (const t of tools) {
              toolSchemas[t.name] = t.inputSchema;
            }

            /**
             * Validate tool arguments against the tool's inputSchema.
             * Checks required fields, types (string, number, boolean, array, object),
             * and rejects unknown properties.
             * Returns an array of error strings (empty = valid).
             */
            // validateAgainstSchema lives in security_helpers.js.

            // ── Domain module wiring ──
            //
            // The contacts, calendar, filters, mail, and compose tool
            // implementations live in mcp_server/domain/*.js. They are loaded
            // here with the same loadSubScript + module.exports shim used for
            // security_helpers.js, because a WebExtension Experiment API loads a
            // single parent script (this file) and supports neither a
            // multi-script list nor ES module import. Each module exports
            // register(ctx): it reads the shared bindings it needs from ctx and
            // assigns its tool functions back onto ctx. We load them at this
            // point -- after every shared constant, XPCOM binding, and infra
            // helper is in scope -- then pull the tool functions into local
            // scope so the dispatch switch below can call them by their original
            // bare names.
            //
            // The shared body/folder/message helpers (stripHtml,
            // extractBodyContent, extractPlainTextBody, openFolder, findMessage)
            // stay in api.js because BOTH mail and compose need them; they are
            // passed through ctx. The temp-attachment / timer / claimed-window
            // state objects also stay here because the shutdown cleanup below
            // (outside this block) still has to drain _tempAttachFiles.
            const __domainCtx = {
              Cc, Ci, Services, MailServices, NetUtil,
              cal, CalEvent, CalTodo, GlodaMsgSearcher,
              MAX_SEARCH_RESULTS_CAP, DEFAULT_MAX_RESULTS, AUDIT_LOG_SUBDIR,
              appendComposeAudit, findIdempotentEntry,
              isContactWritesBlocked, isSkipReviewBlocked,
              isFilterForwardReplyBlocked, isAccountAllowed,
              isFolderAccessible, isMailboxExportBlocked,
              getAccessibleFolder, getAccessibleAccounts,
              paginate, getUserTags,
              openFolder, findMessage,
              stripHtml, extractBodyContent, extractPlainTextBody,
              isSensitiveFilePath, sanitizeHeaderLine, countRecipients,
              isSafeImageSrc, isSafeMarkdownHref, isSystemPrincipalFetchAllowed,
              escapeMarkdownLinkText, renderMarkdownLink,
              wrapUntrustedBody, wrapUntrustedPreview,
              _attachTimers, _tempAttachFiles, _claimedComposeWindows,
            };
            for (const __mod of ["domain/contacts.js", "domain/calendar.js", "domain/filters.js", "domain/mail.js", "domain/compose.js"]) {
              const __scope = { module: { exports: {} } };
              Services.scriptloader.loadSubScript(
                "resource://thunderbird-mcp/mcp_server/" + __mod + "?" + Date.now(),
                __scope
              );
              __scope.module.exports(__domainCtx);
            }
            const {
              searchContacts, createContact, updateContact, deleteContact,
              listCalendars, createEvent, listEvents, updateEvent, deleteEvent,
              listCategories, createTask, listTasks, updateTask,
              listFilters, createFilter, updateFilter, deleteFilter,
              reorderFilters, applyFilters,
              searchMessages, getMessage, getMessageHeaders, getRecentMessages,
              batchGetMessageHeaders, searchByThread, searchAttachments,
              getSenderHistory, displayMessage, updateMessage, deleteMessages,
              createFolder, renameFolder, deleteFolder, moveFolder,
              emptyTrash, emptyJunk, refreshFolder, exportMailbox,
              composeMail, replyToMessage, forwardMessage, saveDraft,
              dryRunCompose, listTemplates, renderTemplate,
            } = __domainCtx;

            function validateToolArgs(name, args) {
              const schema = toolSchemas[name];
              if (!schema) return [`Unknown tool: ${name}`];

              const errors = [];
              const props = schema.properties || {};
              const required = schema.required || [];

              // Check required fields
              for (const key of required) {
                if (args[key] === undefined || args[key] === null) {
                  errors.push(`Missing required parameter: ${key}`);
                }
              }

              // Check types and reject unknown properties
              for (const [key, value] of Object.entries(args)) {
                // Use hasOwnProperty to prevent inherited properties like
                // 'constructor' or 'toString' from bypassing unknown-param checks.
                const propSchema = Object.prototype.hasOwnProperty.call(props, key) ? props[key] : undefined;
                if (!propSchema) {
                  errors.push(`Unknown parameter: ${key}`);
                  continue;
                }
                if (value === undefined || value === null) continue;

                validateAgainstSchema(value, propSchema, key, errors);
              }

              return errors;
            }

            /**
             * Coerce tool arguments to match expected schema types.
             * MCP clients may send "true"/"false" as strings for booleans,
             * "50" as strings for numbers, or JSON-encoded arrays as strings.
             * Mutates and returns the args object.
             */
            function coerceToolArgs(name, args) {
              const schema = toolSchemas[name];
              if (!schema) return args;
              const props = schema.properties || {};
              for (const [key, value] of Object.entries(args)) {
                if (value === undefined || value === null) continue;
                const propSchema = Object.prototype.hasOwnProperty.call(props, key) ? props[key] : undefined;
                if (!propSchema) continue;
                const expected = propSchema.type;
                if (expected === "boolean" && typeof value === "string") {
                  if (value === "true") args[key] = true;
                  else if (value === "false") args[key] = false;
                } else if (expected === "number" && typeof value === "string") {
                  // Reject blank/whitespace strings -- Number("") is 0 which
                  // would silently coerce empty input into a valid number.
                  if (value.trim() === "") continue;
                  const n = Number(value);
                  if (Number.isFinite(n)) args[key] = n;
                } else if (expected === "integer" && typeof value === "string") {
                  if (value.trim() === "") continue;
                  const n = Number(value);
                  if (Number.isFinite(n) && Number.isInteger(n)) args[key] = n;
                } else if (expected === "array" && typeof value === "string") {
                  try {
                    const parsed = JSON.parse(value);
                    if (Array.isArray(parsed)) args[key] = parsed;
                  } catch (e) {
                    // Leave value as-is so validator surfaces a typed error to the client.
                    console.warn(`thunderbird-mcp: coerceToolArgs JSON.parse failed for key=${key}:`, e.message);
                  }
                }
              }
              return args;
            }

            async function callTool(name, args) {
              switch (name) {
                case "listAccounts":
                  return listAccounts();
                case "listFolders":
                  return listFolders(args.accountId, args.folderPath);
                case "searchMessages":
                  return await searchMessages(args.query || "", args.folderPath, args.startDate, args.endDate, args.maxResults, args.offset, args.sortOrder, args.unreadOnly, args.flaggedOnly, args.tag, args.includeSubfolders, args.countOnly, args.searchBody);
                case "getMessage":
                  return await getMessage(args.messageId, args.folderPath, args.saveAttachments, args.bodyFormat, args.rawSource);
                case "getMessageHeaders":
                  return getMessageHeaders(args.messageId, args.folderPath);
                case "batchGetMessageHeaders":
                  return batchGetMessageHeaders(args.messageIds, args.folderPath);
                case "searchByThread":
                  return searchByThread(args.messageId, args.folderPath, args.maxResults);
                case "searchAttachments":
                  return await searchAttachments(args.nameContains, args.contentType, args.folderPath, args.maxResults, args.scanCap);
                case "getSenderHistory":
                  return getSenderHistory(args.email, args.maxResults, args.scanCap, args.sinceDays);
                case "exportMailbox":
                  return await exportMailbox(args.folderPath, args.maxMessages, args.includeBody, args.includeAttachmentMeta, args.scanCap, args.sortOrder);
                case "dryRunCompose":
                  return dryRunCompose(args.to, args.subject, args.body, args.cc, args.bcc, args.isHtml, args.from, args.attachments);
                case "getServerCapabilities":
                  return getServerCapabilities();
                case "listTemplates":
                  return listTemplates();
                case "renderTemplate":
                  return renderTemplate(args.name, args.vars);
                case "getAuditLog": {
                  const filter = {};
                  if (typeof args.tool === "string") filter.tool = args.tool;
                  if (typeof args.since === "string") filter.since = args.since;
                  if (typeof args.until === "string") filter.until = args.until;
                  const max = typeof args.maxEntries === "number" ? args.maxEntries : 200;
                  return readAuditLog(max, filter);
                }
                case "searchContacts":
                  return searchContacts(args.query || "", args.maxResults);
                case "createContact":
                  return createContact(args.email, args.displayName, args.firstName, args.lastName, args.addressBookId);
                case "updateContact":
                  return updateContact(args.contactId, args.email, args.displayName, args.firstName, args.lastName);
                case "deleteContact":
                  return deleteContact(args.contactId);
                case "listCalendars":
                  return listCalendars();
                case "createEvent":
                  return await createEvent(args.title, args.startDate, args.endDate, args.location, args.description, args.calendarId, args.allDay, args.skipReview, args.status);
                case "listEvents":
                  return await listEvents(args.calendarId, args.startDate, args.endDate, args.maxResults);
                case "updateEvent":
                  return await updateEvent(args.eventId, args.calendarId, args.title, args.startDate, args.endDate, args.location, args.description, args.status, args.recurringScope);
                case "deleteEvent":
                  return await deleteEvent(args.eventId, args.calendarId, args.recurringScope);
                case "listCategories":
                  return listCategories();
                case "createTask":
                  return await createTask(args.title, args.dueDate, args.calendarId, args.description, args.priority, args.categories, args.skipReview);
                case "listTasks":
                  return await listTasks(args.calendarId, args.completed, args.dueBefore, args.maxResults);
                case "updateTask":
                  return await updateTask(args.taskId, args.calendarId, args.title, args.dueDate, args.description, args.completed, args.percentComplete, args.priority);
                case "sendMail":
                  return await composeMail(args.to, args.subject, args.body, args.cc, args.bcc, args.isHtml, args.from, args.attachments, args.skipReview, args.idempotencyKey);
                case "saveDraft":
                  return await saveDraft(args.to, args.subject, args.body, args.cc, args.bcc, args.isHtml, args.from, args.attachments);
                case "replyToMessage":
                  return await replyToMessage(args.messageId, args.folderPath, args.body, args.replyAll, args.isHtml, args.to, args.cc, args.bcc, args.from, args.attachments, args.skipReview, args.idempotencyKey);
                case "forwardMessage":
                  return await forwardMessage(args.messageId, args.folderPath, args.to, args.body, args.isHtml, args.cc, args.bcc, args.from, args.attachments, args.skipReview, args.idempotencyKey);
                case "getRecentMessages":
                  return getRecentMessages(args.folderPath, args.daysBack, args.maxResults, args.offset, args.unreadOnly, args.flaggedOnly, args.includeSubfolders);
                case "refreshFolder":
                  return await refreshFolder(args.folderPath, args.timeoutMs);
                case "displayMessage":
                  return displayMessage(args.messageId, args.folderPath, args.displayMode);
                case "deleteMessages":
                  return deleteMessages(args.messageIds, args.folderPath);
                case "updateMessage":
                  return updateMessage(args.messageId, args.messageIds, args.folderPath, args.read, args.flagged, args.addTags, args.removeTags, args.moveTo, args.trash);
                case "createFolder":
                  return createFolder(args.parentFolderPath, args.name);
                case "renameFolder":
                  return renameFolder(args.folderPath, args.newName);
                case "deleteFolder":
                  return deleteFolder(args.folderPath);
                case "emptyTrash":
                  return emptyTrash(args.accountId);
                case "emptyJunk":
                  return emptyJunk(args.accountId);
                case "moveFolder":
                  return moveFolder(args.folderPath, args.newParentPath);
                case "listFilters":
                  return listFilters(args.accountId);
                case "createFilter":
                  return createFilter(args.accountId, args.name, args.enabled, args.type, args.conditions, args.actions, args.insertAtIndex);
                case "updateFilter":
                  return updateFilter(args.accountId, args.filterIndex, args.name, args.enabled, args.type, args.conditions, args.actions);
                case "deleteFilter":
                  return deleteFilter(args.accountId, args.filterIndex);
                case "reorderFilters":
                  return reorderFilters(args.accountId, args.fromIndex, args.toIndex);
                case "applyFilters":
                  return applyFilters(args.accountId, args.folderPath);
                case "getAccountAccess":
                  return getAccountAccess();
                default:
                  throw new Error(`Unknown tool: ${name}`);
              }
            }

            const server = new HttpServer();

            server.registerPathHandler("/", (req, res) => {
              res.processAsync();

              // Verify auth token on ALL requests (including non-POST) to
              // prevent unauthenticated probing of the server.
              let reqToken = "";
              try {
                reqToken = req.getHeader("Authorization") || "";
              } catch {
                // getHeader throws if header is missing in httpd.sys.mjs
              }
              if (!timingSafeEqual(reqToken, `Bearer ${authToken}`)) {
                res.setStatusLine("1.1", 403, "Forbidden");
                res.setHeader("Content-Type", "application/json; charset=utf-8", false);
                res.write(JSON.stringify({
                  jsonrpc: "2.0",
                  id: null,
                  error: { code: -32600, message: "Invalid or missing auth token" }
                }));
                res.finish();
                return;
              }

              if (req.method !== "POST") {
                res.setStatusLine("1.1", 405, "Method Not Allowed");
                res.setHeader("Allow", "POST", false);
                res.setHeader("Content-Type", "application/json; charset=utf-8", false);
                res.write(JSON.stringify({
                  jsonrpc: "2.0",
                  id: null,
                  error: { code: -32600, message: "Method not allowed" }
                }));
                res.finish();
                return;
              }

              // Reject oversized request bodies to prevent memory exhaustion
              let contentLength = 0;
              try {
                contentLength = parseInt(req.getHeader("Content-Length"), 10) || 0;
              } catch {
                // Header missing — will be 0
              }
              if (contentLength > MAX_REQUEST_BODY) {
                res.setStatusLine("1.1", 413, "Payload Too Large");
                res.setHeader("Content-Type", "application/json; charset=utf-8", false);
                res.write(JSON.stringify({
                  jsonrpc: "2.0",
                  id: null,
                  error: { code: -32600, message: "Request body too large" }
                }));
                res.finish();
                return;
              }

              let message;
              try {
                message = JSON.parse(readRequestBody(req));
              } catch {
                res.setStatusLine("1.1", 200, "OK");
                res.setHeader("Content-Type", "application/json; charset=utf-8", false);
                res.write(JSON.stringify({
                  jsonrpc: "2.0",
                  id: null,
                  error: { code: -32700, message: "Parse error" }
                }));
                res.finish();
                return;
              }

              if (!message || typeof message !== "object" || Array.isArray(message)) {
                res.setStatusLine("1.1", 200, "OK");
                res.setHeader("Content-Type", "application/json; charset=utf-8", false);
                res.write(JSON.stringify({
                  jsonrpc: "2.0",
                  id: null,
                  error: { code: -32600, message: "Invalid Request" }
                }));
                res.finish();
                return;
              }

              const { id, method, params } = message;

              // Notifications don't expect a response
              if (typeof method === "string" && method.startsWith("notifications/")) {
                res.setStatusLine("1.1", 204, "No Content");
                res.finish();
                return;
              }

              (async () => {
                try {
                  let result;
                  switch (method) {
                    case "initialize": {
                      // Per MCP lifecycle: respond with the requested version if
                      // we support it, otherwise the latest version we know.
                      // Behavior never depends on the version inside Thunderbird,
                      // so we accept any well-known protocol version.
                      const requested = params?.protocolVersion;
                      if (typeof requested !== "string") {
                        res.setStatusLine("1.1", 200, "OK");
                        res.setHeader("Content-Type", "application/json; charset=utf-8", false);
                        res.write(JSON.stringify({
                          jsonrpc: "2.0",
                          id: id ?? null,
                          error: { code: -32602, message: "Invalid params: protocolVersion must be a string" }
                        }));
                        res.finish();
                        return;
                      }
                      const negotiated = MCP_SUPPORTED_PROTOCOL_VERSIONS.has(requested)
                        ? requested
                        : MCP_LATEST_PROTOCOL_VERSION;
                      result = {
                        protocolVersion: negotiated,
                        capabilities: { tools: {} },
                        serverInfo: { name: "thunderbird-mcp", version: getExtVersion() }
                      };
                      break;
                    }
                    case "resources/list":
                      result = { resources: [] };
                      break;
                    case "prompts/list":
                      result = { prompts: [] };
                      break;
                    case "tools/list":
                      // Strip internal metadata (group, crud, title) — only expose MCP-spec fields
                      result = { tools: tools.filter(t => isToolEnabled(t.name)).map(({ name, description, inputSchema }) => ({ name, description, inputSchema })) };
                      break;
                    case "tools/call":
                      if (!params?.name) {
                        throw new Error("Missing tool name");
                      }
                      if (!isToolEnabled(params.name)) {
                        throw new Error(`Tool is disabled: ${params.name}`);
                      }
                      // Rate-limit check. Runs BEFORE coerce/validate so an
                      // agent stuck in a tight loop is throttled even if it
                      // is sending malformed arguments.
                      {
                        const rl = consumeRateLimit(__rateLimiterState, params.name);
                        if (!rl.allowed) {
                          throw new Error(
                            `Rate limit exceeded for '${params.name}': ` +
                            `${rl.limit} calls per ${Math.round(rl.windowMs / 1000)}s. ` +
                            `Retry in ${Math.ceil(rl.resetAfterMs / 1000)}s.`
                          );
                        }
                      }
                      {
                        const toolArgs = coerceToolArgs(params.name, params.arguments || {});
                        const validationErrors = validateToolArgs(params.name, toolArgs);
                        if (validationErrors.length > 0) {
                          throw new Error(`Invalid parameters for '${params.name}': ${validationErrors.join("; ")}`);
                        }
                        result = {
                          content: [{
                            type: "text",
                            text: JSON.stringify(await callTool(params.name, toolArgs), null, 2)
                          }]
                        };
                      }
                      break;
                    default:
                      res.setStatusLine("1.1", 200, "OK");
                      res.setHeader("Content-Type", "application/json; charset=utf-8", false);
                      res.write(JSON.stringify({
                        jsonrpc: "2.0",
                        id: id ?? null,
                        error: { code: -32601, message: "Method not found" }
                      }));
                      res.finish();
                      return;
                  }
                  res.setStatusLine("1.1", 200, "OK");
                  // charset=utf-8 is critical for proper emoji handling in responses
                  res.setHeader("Content-Type", "application/json; charset=utf-8", false);
                  res.write(JSON.stringify({ jsonrpc: "2.0", id: id ?? null, result }));
                } catch (e) {
                  res.setStatusLine("1.1", 200, "OK");
                  res.setHeader("Content-Type", "application/json; charset=utf-8", false);
                  res.write(JSON.stringify({
                    jsonrpc: "2.0",
                    id: id ?? null,
                    error: { code: -32000, message: e.toString() }
                  }));
                }
                res.finish();
              })().catch((e) => {
                console.error("thunderbird-mcp: unhandled dispatch error:", e);
                try { res.finish(); } catch {}
              });
            });

            // Try the default port first, then fall back to nearby ports
            let boundPort = null;
            for (let attempt = 0; attempt < MCP_MAX_PORT_ATTEMPTS; attempt++) {
              const tryPort = MCP_DEFAULT_PORT + attempt;
              try {
                server.start(tryPort);
                boundPort = tryPort;
                break;
              } catch (portErr) {
                if (attempt === MCP_MAX_PORT_ATTEMPTS - 1) {
                  throw new Error(`Could not bind to any port in range ${MCP_DEFAULT_PORT}-${tryPort}: ${portErr}`);
                }
                console.warn(`Port ${tryPort} in use, trying ${tryPort + 1}...`);
              }
            }

            globalThis.__tbMcpServer = server;
            let connFilePath;
            try {
              connFilePath = writeConnectionInfo(boundPort, authToken);
            } catch (writeErr) {
              // Connection file write failed -- stop the orphaned server
              try { server.stop(() => {}); } catch (e) { console.error("thunderbird-mcp: server.stop failed:", e); }
              globalThis.__tbMcpServer = null;
              throw writeErr;
            }
            console.log(`Thunderbird MCP server listening on port ${boundPort}`);
            console.log(`Connection info written to ${connFilePath}`);
            // Clear any prior start error now that we're fully up.
            globalThis.__tbMcpStartError = null;
            return { success: true, port: boundPort };
          } catch (e) {
            console.error("Failed to start MCP server:", e);
            // Persist the error so getServerInfo can surface it in the
            // options page; otherwise the user just sees "Running" forever
            // while the actual error sits in the Error Console.
            const errStr = e && e.toString ? e.toString() : String(e);
            const stack = e && e.stack ? e.stack : "";
            globalThis.__tbMcpStartError = errStr;
            // Also write to <TmpD>/thunderbird-mcp/start-error.log so the
            // error survives a TB restart and can be inspected from the
            // host shell even when the Error Console isn't open. Best-
            // effort; failures here must not mask the original problem.
            try {
              const tmpDir = Services.dirsvc.get("TmpD", Ci.nsIFile);
              tmpDir.append("thunderbird-mcp");
              if (!tmpDir.exists()) tmpDir.create(Ci.nsIFile.DIRECTORY_TYPE, 0o700);
              const f = tmpDir.clone();
              f.append("start-error.log");
              const out = Cc["@mozilla.org/network/file-output-stream;1"].createInstance(Ci.nsIFileOutputStream);
              out.init(f, 0x02 | 0x08 | 0x20, 0o600, 0);
              const conv = Cc["@mozilla.org/intl/converter-output-stream;1"].createInstance(Ci.nsIConverterOutputStream);
              conv.init(out, "UTF-8");
              conv.writeString(new Date().toISOString() + " " + errStr + "\n" + stack + "\n");
              conv.close();
            } catch { /* best-effort */ }
            // Stop server if it was started but something else failed
            if (globalThis.__tbMcpServer) {
              try { globalThis.__tbMcpServer.stop(() => {}); } catch (e) { console.error("thunderbird-mcp: server.stop failed:", e); }
              globalThis.__tbMcpServer = null;
            }
            // Clear cached promise so a retry can attempt to bind again
            globalThis.__tbMcpStartPromise = null;
            removeConnectionInfo();
            return { success: false, error: e.toString() };
          }
          })();
          // Set sentinel BEFORE awaiting to prevent race with concurrent start() calls
          globalThis.__tbMcpStartPromise = startPromise;
          return await startPromise;
        },

        getServerInfo: async function() {
          let port = null;
          let connectionFile = null;
          let buildVersion = null;
          let buildDate = null;

          // Read build info from bundled file via resource: protocol
          try {
            const uri = Services.io.newURI("resource://thunderbird-mcp/buildinfo.json");
            const channel = Services.io.newChannelFromURI(uri, null,
              Services.scriptSecurityManager.getSystemPrincipal(), null,
              Ci.nsILoadInfo.SEC_ALLOW_CROSS_ORIGIN_SEC_CONTEXT_IS_NULL,
              Ci.nsIContentPolicy.TYPE_OTHER);
            const sis = Cc["@mozilla.org/scriptableinputstream;1"]
              .createInstance(Ci.nsIScriptableInputStream);
            sis.init(channel.open());
            const text = sis.read(sis.available());
            sis.close();
            const bi = JSON.parse(text);
            buildVersion = bi.version || bi.commit || null;
            buildDate = bi.builtAt || null;
          } catch (e) {
            // buildinfo.json may legitimately be absent in dev builds; surface anything else.
            if (e?.name !== "NS_ERROR_FILE_NOT_FOUND") {
              console.warn("thunderbird-mcp: read buildinfo failed:", e);
            }
          }

          // Read connection info from temp file using XPCOM file I/O
          try {
            const connInfo = readConnectionInfo();
            connectionFile = connInfo.path;
            if (connInfo.data) {
              port = connInfo.data.port || null;
            }
          } catch (e) {
            // Connection file is absent before the server first binds; log other faults.
            if (e?.name !== "NS_ERROR_FILE_NOT_FOUND") {
              console.warn("thunderbird-mcp: read connection info failed:", e);
            }
          }

          // `running` is true only when the HTTP server has been instantiated
          // and the start promise resolved cleanly. Checking
          // `!!globalThis.__tbMcpStartPromise` alone was misleading because a
          // rejected promise is still truthy -- the page would say "Running"
          // while the server had silently failed to bind. Surface the actual
          // start error so the options page can show it.
          const startError = globalThis.__tbMcpStartError || null;
          const running = !!globalThis.__tbMcpServer && !startError;
          return {
            running,
            port,
            connectionFile,
            buildVersion,
            buildDate,
            startError,
          };
        },

        getCurrentAuthToken: async function() {
          let authToken = "";
          try {
            const connInfo = readConnectionInfo();
            if (connInfo.data && typeof connInfo.data.token === "string") {
              authToken = connInfo.data.token;
            }
          } catch (e) {
            if (e?.name !== "NS_ERROR_FILE_NOT_FOUND") {
              console.warn("thunderbird-mcp: read auth token failed:", e);
            }
          }
          return { authToken };
        },

        getAccountAccessConfig: async function() {
          const { MailServices } = ChromeUtils.importESModule(
            "resource:///modules/MailServices.sys.mjs"
          );
          let allowed = [];
          try {
            const pref = Services.prefs.getStringPref(PREF_ALLOWED_ACCOUNTS, "");
            if (pref) allowed = JSON.parse(pref);
          } catch (e) {
            // Falls back to "all accounts allowed"; surface the corruption.
            console.warn("thunderbird-mcp: account-access pref is not valid JSON:", e.message);
          }

          const accounts = [];
          for (const account of MailServices.accounts.accounts) {
            const server = account.incomingServer;
            accounts.push({
              id: account.key,
              name: server.prettyName,
              type: server.type,
              allowed: allowed.length === 0 || allowed.includes(account.key),
            });
          }
          return {
            mode: allowed.length === 0 ? "all" : "restricted",
            allowedAccountIds: allowed,
            accounts,
          };
        },

        getToolAccessConfig: async function() {
          // Use same fail-closed parsing as getDisabledTools() so the UI
          // accurately reflects the server's actual state on corrupt prefs
          let disabled = [];
          let corrupt = false;
          try {
            const pref = Services.prefs.getStringPref(PREF_DISABLED_TOOLS, "");
            if (pref) {
              const parsed = JSON.parse(pref);
              if (!Array.isArray(parsed)) {
                corrupt = true;
              } else {
                disabled = parsed;
              }
            }
          } catch {
            corrupt = true;
          }

          // Build tool list with group/crud metadata, sorted by group then CRUD order
          const toolList = tools
            .map(t => ({
              name: t.name,
              group: t.group,
              crud: t.crud,
              enabled: corrupt ? UNDISABLEABLE_TOOLS.has(t.name) : !disabled.includes(t.name),
              undisableable: UNDISABLEABLE_TOOLS.has(t.name),
            }))
            .sort((a, b) => {
              const gA = GROUP_ORDER[a.group] ?? 99;
              const gB = GROUP_ORDER[b.group] ?? 99;
              if (gA !== gB) return gA - gB;
              return (CRUD_ORDER[a.crud] ?? 99) - (CRUD_ORDER[b.crud] ?? 99);
            });
          const result = {
            mode: corrupt ? "error" : (disabled.length === 0 ? "all" : "restricted"),
            disabledTools: disabled,
            groups: GROUP_LABELS,
            tools: toolList,
          };
          if (corrupt) {
            result.error = "Disabled tools preference is corrupt. All non-infrastructure tools are blocked. Save to reset.";
          }
          return result;
        },

        setToolAccess: async function(disabledTools) {
          if (!Array.isArray(disabledTools)) {
            return { error: "disabledTools must be an array" };
          }
          // Validate types first, then semantic checks
          if (!disabledTools.every(t => typeof t === "string")) {
            return { error: "All tool names must be strings" };
          }
          // Reject internal sentinel values
          if (disabledTools.includes("__all__")) {
            return { error: "Invalid tool name: __all__" };
          }
          // Validate: can't disable undisableable tools
          const blocked = disabledTools.filter(t => UNDISABLEABLE_TOOLS.has(t));
          if (blocked.length > 0) {
            return { error: `Cannot disable infrastructure tools: ${blocked.join(", ")}` };
          }

          if (disabledTools.length === 0) {
            try { Services.prefs.clearUserPref(PREF_DISABLED_TOOLS); } catch { /* ignore */ }
          } else {
            Services.prefs.setStringPref(PREF_DISABLED_TOOLS, JSON.stringify(disabledTools));
          }
          return {
            success: true,
            mode: disabledTools.length === 0 ? "all" : "restricted",
            disabledTools,
          };
        },

        setAccountAccess: async function(allowedAccountIds) {
          if (!Array.isArray(allowedAccountIds)) {
            return { error: "allowedAccountIds must be an array" };
          }
          const { MailServices } = ChromeUtils.importESModule(
            "resource:///modules/MailServices.sys.mjs"
          );
          const validIds = new Set();
          for (const account of MailServices.accounts.accounts) {
            validIds.add(account.key);
          }
          const invalid = allowedAccountIds.filter(id => !validIds.has(id));
          if (invalid.length > 0) {
            return { error: `Unknown account IDs: ${invalid.join(", ")}` };
          }

          if (allowedAccountIds.length === 0) {
            try { Services.prefs.clearUserPref(PREF_ALLOWED_ACCOUNTS); } catch { /* ignore */ }
          } else {
            Services.prefs.setStringPref(PREF_ALLOWED_ACCOUNTS, JSON.stringify(allowedAccountIds));
          }
          return {
            success: true,
            mode: allowedAccountIds.length === 0 ? "all" : "restricted",
            allowedAccountIds,
          };
        },

        getBlockSkipReview: async function() {
          let blocked = true;
          try {
            blocked = Services.prefs.getBoolPref(PREF_BLOCK_SKIPREVIEW, true);
          } catch { /* ignore */ }
          return { blockSkipReview: blocked };
        },

        setBlockSkipReview: async function(blockSkipReview) {
          if (typeof blockSkipReview !== "boolean") {
            return { error: "blockSkipReview must be a boolean" };
          }
          // Default is true; persist the explicit value either way so the user's
          // choice survives independent of the default we ship.
          Services.prefs.setBoolPref(PREF_BLOCK_SKIPREVIEW, blockSkipReview);
          return { success: true, blockSkipReview };
        },

        getBlockFilterForwardReply: async function() {
          let blocked = true;
          try {
            blocked = Services.prefs.getBoolPref(PREF_BLOCK_FILTER_FORWARD_REPLY, true);
          } catch { /* ignore */ }
          return { blockFilterForwardReply: blocked };
        },

        setBlockFilterForwardReply: async function(blockFilterForwardReply) {
          if (typeof blockFilterForwardReply !== "boolean") {
            return { error: "blockFilterForwardReply must be a boolean" };
          }
          Services.prefs.setBoolPref(PREF_BLOCK_FILTER_FORWARD_REPLY, blockFilterForwardReply);
          return { success: true, blockFilterForwardReply };
        },

        getBlockContactWrites: async function() {
          let blocked = true;
          try {
            blocked = Services.prefs.getBoolPref(PREF_BLOCK_CONTACT_WRITES, true);
          } catch { /* ignore */ }
          return { blockContactWrites: blocked };
        },

        setBlockContactWrites: async function(blockContactWrites) {
          if (typeof blockContactWrites !== "boolean") {
            return { error: "blockContactWrites must be a boolean" };
          }
          Services.prefs.setBoolPref(PREF_BLOCK_CONTACT_WRITES, blockContactWrites);
          return { success: true, blockContactWrites };
        },

        getBlockMailboxExport: async function() {
          let blocked = true;
          try {
            blocked = Services.prefs.getBoolPref(PREF_BLOCK_MAILBOX_EXPORT, true);
          } catch { /* ignore */ }
          return { blockMailboxExport: blocked };
        },

        setBlockMailboxExport: async function(blockMailboxExport) {
          if (typeof blockMailboxExport !== "boolean") {
            return { error: "blockMailboxExport must be a boolean" };
          }
          Services.prefs.setBoolPref(PREF_BLOCK_MAILBOX_EXPORT, blockMailboxExport);
          return { success: true, blockMailboxExport };
        },

        readAuditLog: async function(maxEntries, filter) {
          // Defensive normalization of the filter object so a malformed call
          // from options.html cannot crash the experiment-API scope. Any
          // throw inside readAuditLog is reduced to a structured error so
          // the options page sees a useful message instead of the generic
          // "An unexpected error occurred" that Mozilla wraps thrown
          // experiment-API errors in.
          try {
            const safeFilter = (filter && typeof filter === "object" && !Array.isArray(filter)) ? filter : null;
            return readAuditLog(maxEntries, safeFilter);
          } catch (e) {
            return { entries: [], totalScanned: 0, truncated: false, errors: [{ reason: String(e) }] };
          }
        },

        clearAuditLog: async function() {
          try { return clearAuditLog(); }
          catch (e) { return { error: String(e) }; }
        },

        getStableAuthToken: async function() {
          return { stableAuthToken: getStableAuthTokenPref() };
        },

        setStableAuthToken: async function(stableAuthToken) {
          if (typeof stableAuthToken !== "string") {
            return { error: "stableAuthToken must be a string" };
          }
          stableAuthToken = stableAuthToken.trim();
          if (stableAuthToken && !AUTH_TOKEN_PATTERN.test(stableAuthToken)) {
            return { error: "stableAuthToken must be empty or 64 lowercase hex characters" };
          }
          if (stableAuthToken) {
            Services.prefs.setStringPref(PREF_STABLE_AUTH_TOKEN, stableAuthToken);
          } else {
            try { Services.prefs.clearUserPref(PREF_STABLE_AUTH_TOKEN); } catch { /* ignore */ }
          }
          return { success: true, stableAuthToken };
        },

        generateAuthToken: async function() {
          return { authToken: generateAuthToken() };
        },
      }
    };
  }

  onShutdown(isAppShutdown) {
    // Stop the HTTP server so the port is released
    if (globalThis.__tbMcpServer) {
      try { globalThis.__tbMcpServer.stop(() => {}); } catch { /* ignore */ }
      globalThis.__tbMcpServer = null;
    }
    // Clear the start promise so a fresh start can occur on reload
    globalThis.__tbMcpStartPromise = null;

    // Always clean up the connection info file so stale tokens don't linger
    // (Inlined here because removeConnectionInfo() is scoped inside start())
    try {
      const tmpDir = Services.dirsvc.get("TmpD", Ci.nsIFile);
      tmpDir.append("thunderbird-mcp");
      const connFile = tmpDir.clone();
      connFile.append("connection.json");
      if (connFile.exists()) {
        connFile.remove(false);
      }
    } catch {
      // Best-effort cleanup
    }

    // Always clean up temp attachment files (even on app shutdown) to avoid
    // leaving sensitive decoded attachments on disk.
    for (const tmpPath of _tempAttachFiles) {
      try {
        const f = Cc["@mozilla.org/file/local;1"].createInstance(Ci.nsIFile);
        f.initWithPath(tmpPath);
        if (f.exists()) f.remove(false);
      } catch {}
    }
    _tempAttachFiles.clear();
    if (isAppShutdown) return;
    resProto.setSubstitution("thunderbird-mcp", null);
    Services.obs.notifyObservers(null, "startupcache-invalidate");
  }
};
