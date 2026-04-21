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
// Keep references to active attach timers to prevent GC before they fire.
const _attachTimers = new Set();
// Track temp files created for inline base64 attachments (cleaned up on shutdown).
const _tempAttachFiles = new Set();
const MAX_BASE64_SIZE = 25 * 1024 * 1024; // 25 MB limit for inline base64 data (encoded)
// Must be large enough to carry MAX_BASE64_SIZE plus JSON-RPC framing overhead.
// The httpd.sys.mjs pre-buffer cap uses the same value.
const MAX_REQUEST_BODY = 32 * 1024 * 1024; // 32 MB limit for incoming HTTP request bodies
let _tempFileCounter = 0;
// Delay before injecting attachments into a newly opened compose window.
const COMPOSE_WINDOW_LOAD_DELAY_MS = 1500;
const DEFAULT_MAX_RESULTS = 50;
const PREF_ALLOWED_ACCOUNTS = "extensions.thunderbird-mcp.allowedAccounts";
const PREF_DISABLED_TOOLS = "extensions.thunderbird-mcp.disabledTools";
// Valid group and CRUD values for tool metadata validation
const VALID_GROUPS = ["messages", "folders", "contacts", "calendar", "filters", "system"];
const VALID_CRUD = ["create", "read", "update", "delete"];
// CRUD sort order: read first, then create, update, delete (safe → destructive)
const CRUD_ORDER = { read: 0, create: 1, update: 2, delete: 3 };
// Tools that cannot be disabled via the settings page (infrastructure tools)
const UNDISABLEABLE_TOOLS = new Set(["listAccounts", "listFolders", "getAccountAccess"]);
const MAX_SEARCH_RESULTS_CAP = 200;
const SEARCH_COLLECTION_CAP = 10000;
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
        description: "Update an existing calendar event's title, dates, location, or description",
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
          },
          required: ["eventId", "calendarId"],
        },
      },
      {
        name: "deleteEvent",
        group: "calendar", crud: "delete",
        title: "Delete Event",
        description: "Delete a calendar event",
        inputSchema: {
          type: "object",
          properties: {
            eventId: { type: "string", description: "The event ID (from listEvents results)" },
            calendarId: { type: "string", description: "The calendar ID containing the event (from listEvents results)" },
          },
          required: ["eventId", "calendarId"],
        },
      },
      {
        name: "createTask",
        group: "calendar", crud: "create",
        title: "Create Task",
        description: "Open a pre-filled task dialog in Thunderbird for user review before saving",
        inputSchema: {
          type: "object",
          properties: {
            title: { type: "string", description: "Task title" },
            dueDate: { type: "string", description: "Due date in ISO 8601 format (optional)" },
            calendarId: { type: "string", description: "Target calendar ID (from listCalendars, must have supportsTasks=true)" },
          },
          required: ["title"],
        },
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

            const authToken = generateAuthToken();

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

            /**
             * Get the list of allowed account IDs from preferences.
             * Returns an empty array if no restriction is set (all accounts allowed).
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
                // Fail closed: corrupt pref means block all accounts, not allow all
                console.error("thunderbird-mcp: failed to parse allowed accounts pref, blocking all accounts:", e);
                return ["__invalid__"];
              }
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
             * Get the list of disabled tool names from preferences.
             * Returns an empty array if no tools are disabled (all enabled).
             * Fails closed: corrupt pref disables all tools.
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

            /** Creates an nsIFile instance for the given path. */
            function createLocalFile(path) {
              const file = Cc["@mozilla.org/file/local;1"].createInstance(Ci.nsIFile);
              file.initWithPath(path);
              return file;
            }

            /** Returns user-visible tag keywords from a message header, filtering out internal IMAP flags. */
            function getUserTags(msgHdr) {
              return (msgHdr.getStringProperty("keywords") || "").split(/\s+/).filter(k => k && !INTERNAL_KEYWORDS.has(k.toLowerCase()));
            }

            /**
             * Converts attachment entries to attachment descriptors.
             * Each entry can be:
             *   - A string (file path) — resolved from disk
             *   - An object { name, contentType, base64 } — decoded and written
             *     to a temp file under <TmpD>/thunderbird-mcp/attachments/
             * Returns { descs: [{url, name, size, contentType?}], failed: string[] }
             */
            function filePathsToAttachDescs(filePaths) {
              const descs = [];
              const failed = [];
              if (!filePaths || !Array.isArray(filePaths)) return { descs, failed };
              for (const entry of filePaths) {
                try {
                  if (typeof entry === "string") {
                    // File path attachment
                    const file = createLocalFile(entry);
                    if (file.exists()) {
                      descs.push({ url: Services.io.newFileURI(file).spec, name: file.leafName, size: file.fileSize });
                    } else {
                      failed.push(entry);
                    }
                  } else if (entry && typeof entry === "object" && (entry.base64 || entry.content) && entry.name) {
                    // Inline base64 attachment — decode and write to temp file
                    const b64Data = entry.base64 || entry.content;
                    if (b64Data.length > MAX_BASE64_SIZE) {
                      failed.push(`${entry.name} (exceeds ${MAX_BASE64_SIZE / 1024 / 1024}MB size limit)`);
                      continue;
                    }
                    // Decode base64 to binary bytes
                    let bytes;
                    try {
                      // Use the global atob when available, otherwise fall back
                      const raw = typeof atob === "function" ? atob(b64Data) : ChromeUtils.base64Decode(b64Data);
                      bytes = new Uint8Array(raw.length);
                      for (let i = 0; i < raw.length; i++) bytes[i] = raw.charCodeAt(i);
                    } catch {
                      // Fallback: manual base64 decode (atob may not be available in XPCOM context)
                      try {
                        const lookup = new Uint8Array(256);
                        const chars = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/";
                        for (let i = 0; i < chars.length; i++) lookup[chars.charCodeAt(i)] = i;
                        const clean = b64Data.replace(/[^A-Za-z0-9+/]/g, "");
                        const len = clean.length;
                        const outLen = (len * 3) >> 2;
                        bytes = new Uint8Array(outLen);
                        let p = 0;
                        for (let i = 0; i < len; i += 4) {
                          const a = lookup[clean.charCodeAt(i)];
                          const b = lookup[clean.charCodeAt(i + 1)];
                          const c = lookup[clean.charCodeAt(i + 2)];
                          const d = lookup[clean.charCodeAt(i + 3)];
                          bytes[p++] = (a << 2) | (b >> 4);
                          if (i + 2 < len) bytes[p++] = ((b & 15) << 4) | (c >> 2);
                          if (i + 3 < len) bytes[p++] = ((c & 3) << 6) | d;
                        }
                        bytes = bytes.subarray(0, p);
                      } catch {
                        failed.push(`${entry.name} (invalid base64 data)`);
                        continue;
                      }
                    }
                    const tmpDir = Services.dirsvc.get("TmpD", Ci.nsIFile);
                    tmpDir.append("thunderbird-mcp");
                    tmpDir.append("attachments");
                    if (!tmpDir.exists()) {
                      tmpDir.create(Ci.nsIFile.DIRECTORY_TYPE, 0o700);
                    }
                    const tmpFile = tmpDir.clone();
                    let safeName = (entry.name || entry.filename || "attachment").replace(/[^a-zA-Z0-9._-]/g, "_");
                    if (!safeName || safeName === "." || safeName === "..") safeName = "attachment";
                    tmpFile.append(`${Date.now()}_${++_tempFileCounter}_${safeName}`);
                    // Write via XPCOM binary stream
                    const ostream = Cc["@mozilla.org/network/file-output-stream;1"]
                      .createInstance(Ci.nsIFileOutputStream);
                    ostream.init(tmpFile, 0x02 | 0x08 | 0x20, 0o600, 0);
                    const bstream = Cc["@mozilla.org/binaryoutputstream;1"]
                      .createInstance(Ci.nsIBinaryOutputStream);
                    bstream.setOutputStream(ostream);
                    bstream.writeByteArray(bytes, bytes.length);
                    bstream.close();
                    ostream.close();
                    _tempAttachFiles.add(tmpFile.path);
                    const desc = { url: Services.io.newFileURI(tmpFile).spec, name: entry.name || entry.filename, size: tmpFile.fileSize };
                    if (entry.contentType) desc.contentType = entry.contentType;
                    descs.push(desc);
                  } else {
                    failed.push(typeof entry === "object" ? JSON.stringify(entry) : String(entry));
                  }
                } catch (e) {
                  failed.push(typeof entry === "object" ? (entry.name || JSON.stringify(entry)) : String(entry));
                }
              }
              return { descs, failed };
            }

            /**
             * Injects attachment descriptors into the most recently opened compose window.
             * Uses nsITimer so the window has time to finish loading before injection.
             * Each call gets its own timer stored in _attachTimers to prevent GC.
             *
             * Known limitation: uses getMostRecentWindow("msgcompose") which is a race
             * if two compose operations happen within COMPOSE_WINDOW_LOAD_DELAY_MS --
             * attachments from the first may land on the second window.
             * OpenComposeWindowWithParams doesn't return a window handle, so there's
             * no reliable way to target a specific window. Injection failures are
             * silent (callers report success based on pre-validated descriptor counts).
             */
            /**
             * Converts attachment descriptors to nsIMsgAttachment objects.
             * Shared by injectAttachmentsAsync (compose window) and
             * sendMessageDirectly (headless send).
             */
            function descsToMsgAttachments(attachDescs) {
              const result = [];
              for (const desc of attachDescs) {
                try {
                  const att = Cc["@mozilla.org/messengercompose/attachment;1"]
                    .createInstance(Ci.nsIMsgAttachment);
                  att.url = desc.url;
                  att.name = desc.name;
                  if (desc.size != null) att.size = desc.size;
                  if (desc.contentType) att.contentType = desc.contentType;
                  result.push(att);
                } catch {}
              }
              return result;
            }

            function injectAttachmentsAsync(attachDescs) {
              if (!attachDescs || attachDescs.length === 0) return;
              const timer = Cc["@mozilla.org/timer;1"].createInstance(Ci.nsITimer);
              _attachTimers.add(timer);
              timer.initWithCallback({
                notify() {
                  _attachTimers.delete(timer);
                  try {
                    const composeWin = Services.wm.getMostRecentWindow("msgcompose");
                    if (!composeWin || typeof composeWin.AddAttachments !== "function") return;
                    const attachList = descsToMsgAttachments(attachDescs);
                    if (attachList.length > 0) {
                      composeWin.AddAttachments(attachList);
                    }
                  } catch {}
                }
              }, COMPOSE_WINDOW_LOAD_DELAY_MS, Ci.nsITimer.TYPE_ONE_SHOT);
            }

            /**
             * Sends a message directly via nsIMsgSend without opening a compose window.
             * Used by composeMail, replyToMessage, forwardMessage when skipReview=true.
             *
             * Handles two createAndSendMessage signatures:
             * - TB 102-127 (C++): 18 args, includes aAttachments + aPreloadedAttachments
             * - TB 128+   (JS):  16 args, attachments via composeFields only
             * Attachments are always added to composeFields (works in both).
             * We try the modern 16-arg call first; if TB throws
             * NS_ERROR_XPC_NOT_ENOUGH_ARGS, fall back to the legacy 18-arg call.
             */
            function sendMessageDirectly(composeFields, identity, attachDescs, originalMsgURI, compType) {
              if (!identity) {
                return Promise.resolve({ error: "No identity available for direct send" });
              }

              const SEND_TIMEOUT_MS = 120000; // 2 min safety timeout

              return new Promise((resolve) => {
                let settled = false;
                const settle = (result) => {
                  if (!settled) {
                    settled = true;
                    resolve(result);
                  }
                };

                // Safety timeout -- if neither listener callback nor error fires
                const timer = Cc["@mozilla.org/timer;1"].createInstance(Ci.nsITimer);
                timer.initWithCallback({
                  notify() { settle({ error: "Send timed out after " + (SEND_TIMEOUT_MS / 1000) + "s" }); }
                }, SEND_TIMEOUT_MS, Ci.nsITimer.TYPE_ONE_SHOT);

                try {
                  const msgSend = Cc["@mozilla.org/messengercompose/send;1"]
                    .createInstance(Ci.nsIMsgSend);

                  // Populate sender fields from identity (normally done by compose window)
                  if (identity.email) {
                    const name = identity.fullName || "";
                    composeFields.from = name
                      ? `"${name}" <${identity.email}>`
                      : identity.email;
                  }
                  if (identity.organization) {
                    composeFields.organization = identity.organization;
                  }

                  // Add attachments to composeFields (works in all TB versions)
                  for (const att of descsToMsgAttachments(attachDescs)) {
                    composeFields.addAttachment(att);
                  }

                  // Extract body -- createAndSendMessage takes it as a separate param
                  const body = composeFields.body || "";

                  // Resolve account key from identity
                  let accountKey = "";
                  try {
                    for (const account of MailServices.accounts.accounts) {
                      for (let i = 0; i < account.identities.length; i++) {
                        if (account.identities[i].key === identity.key) {
                          accountKey = account.key;
                          break;
                        }
                      }
                      if (accountKey) break;
                    }
                  } catch {}

                  const listener = {
                    QueryInterface: ChromeUtils.generateQI(["nsIMsgSendListener"]),
                    onStartSending() {},
                    onProgress() {},
                    onSendProgress() {},
                    onStatus() {},
                    onStopSending(msgID, status) {
                      timer.cancel();
                      if (Components.isSuccessCode(status)) {
                        settle({ success: true, message: "Message sent" });
                      } else {
                        settle({ error: `Send failed (status: 0x${status.toString(16)})` });
                      }
                    },
                    onGetDraftFolderURI() {},
                    onSendNotPerformed(msgID, status) {
                      timer.cancel();
                      settle({ error: "Send was not performed" });
                    },
                    onTransportSecurityError(msgID, status, secInfo, location) {
                      timer.cancel();
                      settle({ error: `Transport security error${location ? ": " + location : ""}` });
                    },
                  };

                  // Common args shared by both signatures (positions 1-10)
                  const commonArgs = [
                    null,                           // editor
                    identity,                       // identity
                    accountKey,                     // account key
                    composeFields,                  // fields
                    false,                          // isDigest
                    false,                          // dontDeliver
                    Ci.nsIMsgCompDeliverMode.Now,   // deliver mode
                    null,                           // msgToReplace
                    "text/html",                    // body type
                    body,                           // body
                  ];

                  // Tail args shared by both (parentWindow..compType)
                  const tailArgs = [
                    null,                           // parent window
                    null,                           // progress
                    listener,                       // listener
                    "",                             // password
                    originalMsgURI || "",           // original msg URI
                    compType,                       // compose type
                  ];

                  // Try modern 16-arg signature first (TB 128+).
                  // On TB 102-127, XPCOM throws NS_ERROR_XPC_NOT_ENOUGH_ARGS
                  // (0x80570001), so we fall back to legacy 18-arg with null
                  // attachment params (attachments already on composeFields).
                  // Modern TB may return a Promise -- catch async rejections.
                  let sendResult;
                  try {
                    sendResult = msgSend.createAndSendMessage(...commonArgs, ...tailArgs);
                  } catch (e) {
                    const isArgError = (e && e.result === 0x80570001) ||
                      String(e).includes("Not enough arguments");
                    if (isArgError) {
                      sendResult = msgSend.createAndSendMessage(...commonArgs, null, null, ...tailArgs);
                    } else {
                      throw e;
                    }
                  }
                  // Handle async Promise from modern TB (128+)
                  if (sendResult && typeof sendResult.catch === "function") {
                    sendResult.catch(e => {
                      timer.cancel();
                      settle({ error: e.toString() });
                    });
                  }
                } catch (e) {
                  timer.cancel();
                  settle({ error: e.toString() });
                }
              });
            }

            function escapeHtml(s) {
              return s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;').replace(/'/g, '&#39;');
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

            /**
             * Converts HTML to markdown using DOMParser for structure-preserving
             * body extraction. Handles headings, links, bold/italic, lists,
             * blockquotes, code blocks, images, and horizontal rules. Email
             * tables (usually layout, not data) are flattened to text.
             * Falls back to stripHtml if DOMParser is unavailable.
             */
            function htmlToMarkdown(html) {
              if (!html) return "";
              try {
                const doc = new DOMParser().parseFromString(html, "text/html");

                function walkChildren(node) {
                  return Array.from(node.childNodes).map(walk).join("");
                }

                function walk(node) {
                  if (node.nodeType === 3) { // Text
                    return node.textContent.replace(/[ \t]+/g, " ");
                  }
                  if (node.nodeType !== 1) return "";
                  const tag = node.tagName.toLowerCase();
                  const inner = () => walkChildren(node);

                  switch (tag) {
                    case "script": case "style": case "head": return "";
                    case "br": return "\n";
                    case "hr": return "\n\n---\n\n";
                    case "p": case "div": case "section": case "article":
                      return "\n\n" + inner().trim() + "\n\n";
                    case "h1": return "\n\n# " + inner().trim() + "\n\n";
                    case "h2": return "\n\n## " + inner().trim() + "\n\n";
                    case "h3": return "\n\n### " + inner().trim() + "\n\n";
                    case "h4": return "\n\n#### " + inner().trim() + "\n\n";
                    case "h5": return "\n\n##### " + inner().trim() + "\n\n";
                    case "h6": return "\n\n###### " + inner().trim() + "\n\n";
                    case "strong": case "b": {
                      const t = inner().trim();
                      return t ? "**" + t + "**" : "";
                    }
                    case "em": case "i": {
                      const t = inner().trim();
                      return t ? "*" + t + "*" : "";
                    }
                    case "a": {
                      const href = node.getAttribute("href") || "";
                      const text = inner().trim();
                      // Skip empty/anchor-only links and mailto: without text
                      if (!text && !href) return "";
                      if (href && text && text !== href) return `[${text}](${href})`;
                      return text || href;
                    }
                    case "img": {
                      const alt = node.getAttribute("alt") || "";
                      const src = node.getAttribute("src") || "";
                      // Skip tracking pixels (1x1, tiny, or data: without alt)
                      const w = parseInt(node.getAttribute("width")) || 0;
                      const h = parseInt(node.getAttribute("height")) || 0;
                      if ((w > 0 && w <= 3) || (h > 0 && h <= 3)) return "";
                      if (src.startsWith("data:") && !alt) return "";
                      if (src) return `![${alt}](${src})`;
                      return alt;
                    }
                    case "code": return "`" + node.textContent + "`";
                    case "pre": return "\n\n```\n" + node.textContent.trim() + "\n```\n\n";
                    case "blockquote": {
                      const text = inner().trim();
                      return "\n\n" + text.split("\n").map(l => "> " + l).join("\n") + "\n\n";
                    }
                    case "ul": case "ol": return "\n" + inner() + "\n";
                    case "li": {
                      const parent = node.parentElement;
                      const isOl = parent && parent.tagName.toLowerCase() === "ol";
                      return (isOl ? "1. " : "- ") + inner().trim() + "\n";
                    }
                    // Tables: extract text with spacing (email tables are usually layout)
                    case "table": return "\n\n" + inner().trim() + "\n\n";
                    case "tr": return inner().trim() + "\n";
                    case "td": case "th": return inner().trim() + " ";
                    case "thead": case "tbody": case "tfoot": return inner();
                    default: return inner();
                  }
                }

                const body = doc.body || doc.documentElement;
                let result = walk(body);
                // Collapse excessive newlines, trim
                result = result.replace(/\n{3,}/g, "\n\n").trim();
                return result;
              } catch {
                // DOMParser unavailable or parse failure -- fall back to stripHtml
                return stripHtml(html);
              }
            }

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

            /**
             * Extracts body from a MIME message in the requested format.
             * For "text": uses coerceBodyToPlaintext fast path (original behavior).
             * For "markdown"/"html": walks MIME tree to find raw HTML content.
             */
            function extractFormattedBody(aMimeMsg, bodyFormat) {
              if (bodyFormat === "text") {
                return { body: extractPlainTextBody(aMimeMsg), bodyIsHtml: false };
              }
              // For markdown/html: need raw MIME content, not coerced text
              const { text, isHtml } = extractBodyContent(aMimeMsg);
              if (!text) {
                // MIME tree empty -- try coerce as last resort
                const fallback = extractPlainTextBody(aMimeMsg);
                return { body: fallback, bodyIsHtml: false };
              }
              if (!isHtml) return { body: text, bodyIsHtml: false };
              if (bodyFormat === "html") return { body: text, bodyIsHtml: true };
              // Default: markdown
              return { body: htmlToMarkdown(text), bodyIsHtml: false };
            }

            /**
             * Converts body text to HTML for compose fields.
             * Handles both HTML input (entity-encodes non-ASCII) and plain text.
             */
            function formatBodyHtml(body, isHtml) {
              if (isHtml) {
                let text = (body || "").replace(/\n/g, '');
                text = [...text].map(c => c.codePointAt(0) > 127 ? `&#${c.codePointAt(0)};` : c).join('');
                return text;
              }
              return escapeHtml(body || "").replace(/\n/g, '<br>');
            }

            /**
             * Sets compose identity from `from` param or falls back to default.
             * Returns "" on success, or { error } if `from` was explicitly
             * provided but not found / restricted.  Fallback to default only
             * applies when `from` is omitted.
             */
            function setComposeIdentity(msgComposeParams, from, fallbackServer) {
              if (from) {
                // Explicit `from` -- must resolve or fail, never silently substitute
                const identity = findIdentity(from);
                if (identity) {
                  // findIdentity searches accessible accounts, so this is safe
                  msgComposeParams.identity = identity;
                  return "";
                }
                // Not found in accessible accounts -- check ALL accounts to
                // distinguish "restricted" from "genuinely unknown"
                if (findIdentityIn(MailServices.accounts.accounts, from)) {
                  return { error: `identity ${from} belongs to a restricted account` };
                }
                return { error: `unknown identity: ${from} -- no matching account configured in Thunderbird` };
              }
              // No explicit `from` -- fall back to contextual default
              if (fallbackServer) {
                const account = MailServices.accounts.findAccountForServer(fallbackServer);
                if (account && isAccountAllowed(account.key)) {
                  msgComposeParams.identity = account.defaultIdentity;
                }
              } else {
                const defaultAccount = MailServices.accounts.defaultAccount;
                if (defaultAccount && isAccountAllowed(defaultAccount.key)) {
                  msgComposeParams.identity = defaultAccount.defaultIdentity;
                }
              }
              // If no identity was set (all fallbacks restricted), explicitly set
              // the first accessible identity. Without this, Thunderbird's
              // OpenComposeWindowWithParams fills identity from defaultAccount
              // internally, bypassing account restrictions.
              if (!msgComposeParams.identity) {
                for (const account of getAccessibleAccounts()) {
                  if (account.defaultIdentity) {
                    msgComposeParams.identity = account.defaultIdentity;
                    break;
                  }
                }
                if (!msgComposeParams.identity) {
                  return { error: "No accessible identity found -- all accounts are restricted" };
                }
              }
              return "";
            }

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

	            /**
	             * Finds a single message header by messageId within a folderPath.
	             * Returns { msgHdr, folder, db } or { error }.
	             */
            function findTrashFolder(folder) {
              const TRASH_FLAG = 0x00000100;
              let account = null;
              try {
                account = MailServices.accounts.findAccountForServer(folder.server);
              } catch {
                return null;
              }
              const root = account?.incomingServer?.rootFolder;
              if (!root) return null;

              let fallback = null;
              const TRASH_NAMES = ["trash", "deleted items"];
              const stack = [root];
              while (stack.length > 0) {
                const current = stack.pop();
                try {
                  if (current && typeof current.getFlag === "function" && current.getFlag(TRASH_FLAG)) {
                    return current;
                  }
                } catch {}
                if (!fallback && current?.prettyName && TRASH_NAMES.includes(current.prettyName.toLowerCase())) {
                  fallback = current;
                }
                try {
                  if (current?.hasSubFolders) {
                    for (const sf of current.subFolders) stack.push(sf);
                  }
                } catch {}
              }
              return fallback;
            }

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
	                for (const hdr of db.enumerateMessages()) {
	                  if (hdr.messageId === messageId) {
	                    msgHdr = hdr;
	                    break;
	                  }
	                }
	              }

	              if (!msgHdr) {
	                return { error: `Message not found: ${messageId}` };
	              }

	              return { msgHdr, folder, db };
	            }

            /**
             * Full-text body search using Thunderbird's Gloda index via
             * GlodaMsgSearcher. Searches subject, body, and attachment
             * names. Returns a Promise resolving to the same format as
             * searchMessages. IMAP accounts need offline sync for body
             * indexing; without it only headers are searched.
             */
            function glodaBodySearch(query, folderPath, startDate, endDate, maxResults, offset, sortOrder, unreadOnly, flaggedOnly, tag, countOnly) {
              const requestedLimit = Number(maxResults);
              const effectiveLimit = Math.min(
                Number.isFinite(requestedLimit) && requestedLimit > 0 ? Math.floor(requestedLimit) : DEFAULT_MAX_RESULTS,
                MAX_SEARCH_RESULTS_CAP
              );
              const normalizedSortOrder = sortOrder === "asc" ? "asc" : "desc";
              const parsedStartDate = startDate ? new Date(startDate).getTime() : null;
              const parsedEndDate = endDate ? new Date(endDate).getTime() : null;
              if (parsedStartDate !== null && isNaN(parsedStartDate)) return { error: `Invalid startDate: ${startDate}` };
              if (parsedEndDate !== null && isNaN(parsedEndDate)) return { error: `Invalid endDate: ${endDate}` };
              // Match the regular search path: expand date-only endDate to end of day
              const isDateOnly = endDate && /^\d{4}-\d{2}-\d{2}$/.test(endDate.trim());
              const endDateOffset = isDateOnly ? 86400000 : 0;
              const endDateTs = parsedEndDate !== null ? (parsedEndDate + endDateOffset) * 1000 : null;
              const startDateTs = parsedStartDate !== null ? parsedStartDate * 1000 : null;

              // Resolve folder filter upfront -- match by URI prefix for subfolder inclusion
              let folderFilterURI = null;
              if (folderPath) {
                const result = getAccessibleFolder(folderPath);
                if (result.error) return result;
                folderFilterURI = result.folder.URI;
              }

              return new Promise((resolve) => {
                try {
                  const listener = {
                    onItemsAdded() {},
                    onItemsModified() {},
                    onItemsRemoved() {},
                    onQueryCompleted(collection) {
                      try {
                        const results = [];
                        for (const glodaMsg of collection.items) {
                          if (results.length >= SEARCH_COLLECTION_CAP) break;
                          // Get the underlying msgHdr
                          let msgHdr;
                          try {
                            msgHdr = glodaMsg.folderMessage;
                          } catch { continue; }
                          if (!msgHdr) continue;

                          // Account access control
                          const folder = msgHdr.folder;
                          if (!folder) continue;
                          if (!isFolderAccessible(folder)) continue;

                          // Folder filter (URI prefix match includes subfolders)
                          if (folderFilterURI && !folder.URI.startsWith(folderFilterURI)) continue;

                          // Date filters (timestamps in microseconds)
                          const msgDateTs = msgHdr.date || 0;
                          if (startDateTs !== null && msgDateTs < startDateTs) continue;
                          if (endDateTs !== null && msgDateTs > endDateTs) continue;

                          // Boolean filters
                          if (unreadOnly && msgHdr.isRead) continue;
                          if (flaggedOnly && !msgHdr.isFlagged) continue;
                          if (tag) {
                            const keywords = (msgHdr.getStringProperty("keywords") || "").split(/\s+/);
                            if (!keywords.includes(tag)) continue;
                          }

                          const msgTags = getUserTags(msgHdr);
                          const preview = msgHdr.getStringProperty("preview") || "";
                          const result = {
                            id: msgHdr.messageId,
                            threadId: msgHdr.threadId,
                            subject: msgHdr.mime2DecodedSubject || msgHdr.subject,
                            author: msgHdr.mime2DecodedAuthor || msgHdr.author,
                            recipients: msgHdr.mime2DecodedRecipients || msgHdr.recipients,
                            ccList: msgHdr.ccList,
                            date: msgHdr.date ? new Date(msgHdr.date / 1000).toISOString() : null,
                            folder: folder.prettyName,
                            folderPath: folder.URI,
                            read: msgHdr.isRead,
                            flagged: msgHdr.isFlagged,
                            tags: msgTags,
                            _dateTs: msgDateTs
                          };
                          if (preview) result.preview = preview;
                          results.push(result);
                        }

                        if (countOnly) {
                          resolve({ count: results.length });
                          return;
                        }
                        results.sort((a, b) => normalizedSortOrder === "asc" ? a._dateTs - b._dateTs : b._dateTs - a._dateTs);
                        resolve(paginate(results, offset, effectiveLimit));
                      } catch (e) {
                        resolve({ error: e.toString() });
                      }
                    }
                  };
                  const searcher = new GlodaMsgSearcher(listener, query);
                  searcher.getCollection();
                } catch (e) {
                  resolve({ error: e.toString() });
                }
              });
            }

	            function searchMessages(query, folderPath, startDate, endDate, maxResults, offset, sortOrder, unreadOnly, flaggedOnly, tag, includeSubfolders, countOnly, searchBody) {
	              // Gloda full-body search path (async)
	              if (searchBody) {
	                if (!GlodaMsgSearcher) return { error: "Gloda full-text index is not available" };
	                if (!query) return { error: "searchBody requires a non-empty query" };
	                return glodaBodySearch(query, folderPath, startDate, endDate, maxResults, offset, sortOrder, unreadOnly, flaggedOnly, tag, countOnly);
	              }
	              const results = [];
	              const lowerQuery = (query || "").toLowerCase();
	              const hasQuery = !!lowerQuery;
	              // Parse optional field-operator prefix and split into AND tokens.
	              // Supports: from:Name, subject:Text, to:Email, cc:Email
	              // Without an operator, every token must appear somewhere across all fields.
	              const OPERATOR_RE = /^(from|subject|to|cc):\s*/;
	              let fieldTarget = null;
	              let queryTokens = [];
	              if (hasQuery) {
	                const opMatch = lowerQuery.match(OPERATOR_RE);
	                if (opMatch) {
	                  const opMap = { from: 'author', subject: 'subject', to: 'recipients', cc: 'ccList' };
	                  fieldTarget = opMap[opMatch[1]];
	                  queryTokens = lowerQuery.slice(opMatch[0].length).trim().split(/\s+/).filter(Boolean);
	                } else {
	                  queryTokens = lowerQuery.split(/\s+/).filter(Boolean);
	                }
	              }
	              // Treat whitespace-only queries and bare field operators (e.g. "from:"
	              // with nothing after) as failed queries that match nothing, rather
	              // than silently matching every message. The documented way to match
	              // all messages is to pass an empty string, which keeps hasQuery=false.
	              const failedQuery = hasQuery && queryTokens.length === 0;
	              const parsedStartDate = startDate ? new Date(startDate).getTime() : NaN;
              const parsedEndDate = endDate ? new Date(endDate).getTime() : NaN;
              const startDateTs = Number.isFinite(parsedStartDate) ? parsedStartDate * 1000 : null;
              // Add 24h only for date-only strings (e.g. "2024-01-15") to include the full day.
              // Use regex to detect ISO date-only format rather than checking for "T" which
              // would match arbitrary strings like "Totally invalid".
              const isDateOnly = endDate && /^\d{4}-\d{2}-\d{2}$/.test(endDate.trim());
              const endDateOffset = isDateOnly ? 86400000 : 0;
              const endDateTs = Number.isFinite(parsedEndDate) ? (parsedEndDate + endDateOffset) * 1000 : null;
              const requestedLimit = Number(maxResults);
              const effectiveLimit = Math.min(
                Number.isFinite(requestedLimit) && requestedLimit > 0 ? Math.floor(requestedLimit) : DEFAULT_MAX_RESULTS,
                MAX_SEARCH_RESULTS_CAP
              );
              const normalizedSortOrder = sortOrder === "asc" ? "asc" : "desc";

              function searchFolder(folder) {
                if (results.length >= SEARCH_COLLECTION_CAP) return;

                try {
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
                  if (!db) return;

                  for (const msgHdr of db.enumerateMessages()) {
                    if (results.length >= SEARCH_COLLECTION_CAP) break;

                    // Check cheap numeric/boolean filters before string work
                    const msgDateTs = msgHdr.date || 0;
                    if (startDateTs !== null && msgDateTs < startDateTs) continue;
                    if (endDateTs !== null && msgDateTs > endDateTs) continue;
                    if (unreadOnly && msgHdr.isRead) continue;
                    if (flaggedOnly && !msgHdr.isFlagged) continue;
                    if (tag) {
                      const keywords = (msgHdr.getStringProperty("keywords") || "").split(/\s+/);
                      if (!keywords.includes(tag)) continue;
                    }

                    // IMPORTANT: Use mime2Decoded* properties for searching.
                    // Raw headers contain MIME encoding like "=?UTF-8?Q?...?="
                    // which won't match plain text searches.
                    const preview = msgHdr.getStringProperty("preview") || "";
                    if (failedQuery) continue;
                    if (hasQuery) {
                      const subject = (msgHdr.mime2DecodedSubject || msgHdr.subject || "").toLowerCase();
                      const author = (msgHdr.mime2DecodedAuthor || msgHdr.author || "").toLowerCase();
                      const recipients = (msgHdr.mime2DecodedRecipients || msgHdr.recipients || "").toLowerCase();
                      const ccList = (msgHdr.ccList || "").toLowerCase();
                      // AND-of-tokens: every token must appear somewhere across the fields.
                      // If a field operator (from:, subject:, to:, cc:) was given,
                      // restrict matching to that specific field only.
                      const fieldValues = { subject, author, recipients, ccList };
                      const matches = fieldTarget
                        ? queryTokens.every(t => (fieldValues[fieldTarget] || "").includes(t))
                        : queryTokens.every(t =>
                            subject.includes(t) ||
                            author.includes(t) ||
                            recipients.includes(t) ||
                            ccList.includes(t) ||
                            preview.toLowerCase().includes(t)
                          );
                      if (!matches) continue;
                    }

                    const msgTags = getUserTags(msgHdr);
                    const result = {
                      id: msgHdr.messageId,
                      threadId: msgHdr.threadId, // folder-local, use with folderPath for grouping
                      subject: msgHdr.mime2DecodedSubject || msgHdr.subject,
                      author: msgHdr.mime2DecodedAuthor || msgHdr.author,
                      recipients: msgHdr.mime2DecodedRecipients || msgHdr.recipients,
                      ccList: msgHdr.ccList,
                      date: msgHdr.date ? new Date(msgHdr.date / 1000).toISOString() : null,
                      folder: folder.prettyName,
                      folderPath: folder.URI,
                      read: msgHdr.isRead,
                      flagged: msgHdr.isFlagged,
                      tags: msgTags,
                      _dateTs: msgDateTs
                    };
                    if (preview) result.preview = preview;
                    results.push(result);
                  }
                } catch {
                  // Skip inaccessible folders
                }

                const recurse = includeSubfolders !== false; // default true
                if (recurse && folder.hasSubFolders) {
                  for (const subfolder of folder.subFolders) {
                    if (results.length >= SEARCH_COLLECTION_CAP) break;
                    searchFolder(subfolder);
                  }
                }
              }

              if (folderPath) {
                const result = getAccessibleFolder(folderPath);
                if (result.error) return result;
                searchFolder(result.folder);
              } else {
                for (const account of getAccessibleAccounts()) {
                  if (results.length >= SEARCH_COLLECTION_CAP) break;
                  searchFolder(account.incomingServer.rootFolder);
                }
              }

              if (countOnly) {
                return { count: results.length };
              }

              results.sort((a, b) => normalizedSortOrder === "asc" ? a._dateTs - b._dateTs : b._dateTs - a._dateTs);

              return paginate(results, offset, effectiveLimit);
            }

            function searchContacts(query, maxResults) {
              const results = [];
              const lowerQuery = query.toLowerCase();
              const requestedLimit = Number(maxResults);
              const limit = Number.isFinite(requestedLimit) && requestedLimit > 0
                ? Math.min(Math.floor(requestedLimit), MAX_SEARCH_RESULTS_CAP)
                : DEFAULT_MAX_RESULTS;
              let truncated = false;

              for (const book of MailServices.ab.directories) {
                for (const card of book.childCards) {
                  if (card.isMailList) continue;

                  const email = (card.primaryEmail || "").toLowerCase();
                  const displayName = (card.displayName || "").toLowerCase();
                  const firstName = (card.firstName || "").toLowerCase();
                  const lastName = (card.lastName || "").toLowerCase();

                  if (email.includes(lowerQuery) ||
                      displayName.includes(lowerQuery) ||
                      firstName.includes(lowerQuery) ||
                      lastName.includes(lowerQuery)) {
                    results.push({
                      id: card.UID,
                      displayName: card.displayName,
                      email: card.primaryEmail,
                      firstName: card.firstName,
                      lastName: card.lastName,
                      addressBook: book.dirName,
                      addressBookId: book.URI,
                    });
                  }

                  if (results.length >= limit) { truncated = true; break; }
                }
                if (truncated) break;
              }

              if (truncated) {
                return { contacts: results, hasMore: true, message: `Results limited to ${limit}. Refine your query to see more.` };
              }
              return results;
            }

            /**
             * Find a contact card by UID across all address books.
             * Returns { card, book } or { error }.
             */
            function findContactByUID(contactId) {
              for (const book of MailServices.ab.directories) {
                for (const card of book.childCards) {
                  if (card.UID === contactId) {
                    return { card, book };
                  }
                }
              }
              return { error: `Contact not found: ${contactId}` };
            }

            function createContact(email, displayName, firstName, lastName, addressBookId) {
              try {
                if (typeof email !== "string" || !email) {
                  return { error: "email must be a non-empty string" };
                }

                // Find the target address book
                let targetBook = null;
                if (addressBookId) {
                  for (const book of MailServices.ab.directories) {
                    if (book.dirPrefId === addressBookId || book.UID === addressBookId || book.URI === addressBookId) {
                      targetBook = book;
                      break;
                    }
                  }
                  if (!targetBook) {
                    return { error: `Address book not found: ${addressBookId}` };
                  }
                } else {
                  // Use the first writable address book
                  for (const book of MailServices.ab.directories) {
                    if (!book.readOnly) {
                      targetBook = book;
                      break;
                    }
                  }
                  if (!targetBook) {
                    return { error: "No writable address book found" };
                  }
                }

                const card = Cc["@mozilla.org/addressbook/cardproperty;1"]
                  .createInstance(Ci.nsIAbCard);
                card.primaryEmail = email;
                if (displayName) card.displayName = displayName;
                if (firstName) card.firstName = firstName;
                if (lastName) card.lastName = lastName;

                const newCard = targetBook.addCard(card);
                return {
                  success: true,
                  id: newCard.UID,
                  email: newCard.primaryEmail,
                  displayName: newCard.displayName,
                  addressBook: targetBook.dirName,
                };
              } catch (e) {
                return { error: e.toString() };
              }
            }

            function updateContact(contactId, email, displayName, firstName, lastName) {
              try {
                if (typeof contactId !== "string" || !contactId) {
                  return { error: "contactId must be a non-empty string" };
                }

                const found = findContactByUID(contactId);
                if (found.error) return found;
                const { card, book } = found;

                if (email !== undefined) card.primaryEmail = email;
                if (displayName !== undefined) card.displayName = displayName;
                if (firstName !== undefined) card.firstName = firstName;
                if (lastName !== undefined) card.lastName = lastName;

                book.modifyCard(card);
                return {
                  success: true,
                  id: card.UID,
                  email: card.primaryEmail,
                  displayName: card.displayName,
                  firstName: card.firstName,
                  lastName: card.lastName,
                };
              } catch (e) {
                return { error: e.toString() };
              }
            }

            function deleteContact(contactId) {
              try {
                if (typeof contactId !== "string" || !contactId) {
                  return { error: "contactId must be a non-empty string" };
                }

                const found = findContactByUID(contactId);
                if (found.error) return found;
                const { card, book } = found;

                book.deleteCards([card]);
                return {
                  success: true,
                  message: `Contact "${card.displayName || card.primaryEmail}" deleted`,
                };
              } catch (e) {
                return { error: e.toString() };
              }
            }

            function listCalendars() {
              if (!cal) {
                return { error: "Calendar not available" };
              }
              try {
                return cal.manager.getCalendars().map(c => ({
                  id: c.id,
                  name: c.name,
                  type: c.type,
                  readOnly: c.readOnly,
                  supportsEvents: c.getProperty("capabilities.events.supported") !== false,
                  supportsTasks: c.getProperty("capabilities.tasks.supported") !== false,
                }));
              } catch (e) {
                return { error: e.toString() };
              }
            }

            async function createEvent(title, startDate, endDate, location, description, calendarId, allDay, skipReview) {
              if (!cal || !CalEvent) {
                return { error: "Calendar module not available" };
              }
              try {
                const win = Services.wm.getMostRecentWindow("mail:3pane");
                if (!win && !skipReview) {
                  return { error: "No Thunderbird window found" };
                }

                const startJs = new Date(startDate);
                if (isNaN(startJs.getTime())) {
                  return { error: `Invalid startDate: ${startDate}` };
                }

                let endJs = endDate ? new Date(endDate) : null;
                if (endDate && (!endJs || isNaN(endJs.getTime()))) {
                  return { error: `Invalid endDate: ${endDate}` };
                }

                if (endJs) {
                  if (allDay) {
                    const startDay = new Date(startJs.getFullYear(), startJs.getMonth(), startJs.getDate());
                    const endDay = new Date(endJs.getFullYear(), endJs.getMonth(), endJs.getDate());
                    if (endDay.getTime() < startDay.getTime()) {
                      return { error: "endDate must not be before startDate" };
                    }
                  } else if (endJs.getTime() <= startJs.getTime()) {
                    return { error: "endDate must be after startDate" };
                  }
                }

                const event = new CalEvent();
                event.title = title;

                if (allDay) {
                  const startDt = cal.createDateTime();
                  startDt.resetTo(startJs.getFullYear(), startJs.getMonth(), startJs.getDate(), 0, 0, 0, cal.dtz.floating);
                  startDt.isDate = true;
                  event.startDate = startDt;

                  const endDt = cal.createDateTime();
                  if (endJs) {
                    endDt.resetTo(endJs.getFullYear(), endJs.getMonth(), endJs.getDate(), 0, 0, 0, cal.dtz.floating);
                    endDt.isDate = true;
                    // iCal DTEND is exclusive — bump if same as start
                    if (endDt.compare(startDt) <= 0) {
                      const bumpedEnd = new Date(endJs.getFullYear(), endJs.getMonth(), endJs.getDate());
                      bumpedEnd.setDate(bumpedEnd.getDate() + 1);
                      endDt.resetTo(
                        bumpedEnd.getFullYear(),
                        bumpedEnd.getMonth(),
                        bumpedEnd.getDate(),
                        0,
                        0,
                        0,
                        cal.dtz.floating
                      );
                      endDt.isDate = true;
                    }
                  } else {
                    const defaultEnd = new Date(startJs.getTime());
                    defaultEnd.setDate(defaultEnd.getDate() + 1);
                    endDt.resetTo(
                      defaultEnd.getFullYear(),
                      defaultEnd.getMonth(),
                      defaultEnd.getDate(),
                      0,
                      0,
                      0,
                      cal.dtz.floating
                    );
                    endDt.isDate = true;
                  }
                  event.endDate = endDt;
                } else {
                  event.startDate = cal.dtz.jsDateToDateTime(startJs, cal.dtz.defaultTimezone);
                  if (endJs) {
                    event.endDate = cal.dtz.jsDateToDateTime(endJs, cal.dtz.defaultTimezone);
                  } else {
                    const defaultEnd = new Date(startJs.getTime() + 3600000);
                    event.endDate = cal.dtz.jsDateToDateTime(defaultEnd, cal.dtz.defaultTimezone);
                  }
                }

                if (location) event.setProperty("LOCATION", location);
                if (description) event.setProperty("DESCRIPTION", description);

                // Find target calendar
                const calendars = cal.manager.getCalendars();
                let targetCalendar = null;
                if (calendarId) {
                  targetCalendar = calendars.find(c => c.id === calendarId);
                  if (!targetCalendar) {
                    return { error: `Calendar not found: ${calendarId}` };
                  }
                  if (targetCalendar.readOnly) {
                    return { error: `Calendar is read-only: ${targetCalendar.name}` };
                  }
                } else {
                  targetCalendar = calendars.find(c => !c.readOnly);
                  if (!targetCalendar) {
                    return { error: "No writable calendar found" };
                  }
                }

                event.calendar = targetCalendar;

                if (skipReview) {
                  await targetCalendar.addItem(event);
                  return { success: true, message: `Event "${title}" added to calendar "${targetCalendar.name}"` };
                }

                const args = {
                  calendarEvent: event,
                  calendar: targetCalendar,
                  mode: "new",
                  inTab: false,
                  onOk(item, calendar) {
                    calendar.addItem(item);
                  },
                };

                win.openDialog(
                  "chrome://calendar/content/calendar-event-dialog.xhtml",
                  "_blank",
                  "centerscreen,chrome,titlebar,toolbar,resizable",
                  args
                );

                return { success: true, message: `Event dialog opened for "${title}" on calendar "${targetCalendar.name}"` };
              } catch (e) {
                return { error: e.toString() };
              }
            }

            async function getCalendarItems(calendar, rangeStart, rangeEnd) {
              const FILTER_EVENT = 1 << 3;
              if (typeof calendar.getItemsAsArray === "function") {
                return await calendar.getItemsAsArray(FILTER_EVENT, 0, rangeStart, rangeEnd);
              }
              // Fallback for older Thunderbird versions using ReadableStream
              const items = [];
              const stream = cal.iterate.streamValues(calendar.getItems(FILTER_EVENT, 0, rangeStart, rangeEnd));
              for await (const chunk of stream) {
                for (const i of chunk) items.push(i);
              }
              return items;
            }

            function calDateToISO(dt) {
              if (!dt) return null;
              try { return new Date(dt.nativeTime / 1000).toISOString(); }
              catch { return dt.icalString || null; }
            }

            function formatEvent(item, calendar) {
              const allDay = item.startDate ? item.startDate.isDate : false;
              // For all-day events, iCal DTEND is exclusive. Convert to inclusive
              // (last day of event) so the API is intuitive and round-trips correctly.
              let endDateISO = calDateToISO(item.endDate);
              if (allDay && item.endDate) {
                try {
                  const raw = new Date(item.endDate.nativeTime / 1000);
                  raw.setDate(raw.getDate() - 1);
                  endDateISO = raw.toISOString();
                } catch { /* keep raw value */ }
              }
              const result = {
                id: item.id,
                calendarId: calendar.id,
                calendarName: calendar.name,
                title: item.title || "",
                startDate: calDateToISO(item.startDate),
                endDate: endDateISO,
                location: item.getProperty("LOCATION") || "",
                description: item.getProperty("DESCRIPTION") || "",
                allDay,
                isRecurring: !!item.recurrenceInfo,
              };
              // Occurrences of recurring events share the parent's id.
              // Include recurrenceId so callers can distinguish them.
              if (item.recurrenceId) {
                result.recurrenceId = calDateToISO(item.recurrenceId);
              }
              return result;
            }

            function formatTask(item, calendar) {
              const completed = item.isCompleted || (item.percentComplete === 100);
              const priority = item.priority || 0; // 0=undefined, 1=high, 5=normal, 9=low per iCal
              return {
                id: item.id,
                calendarId: calendar.id,
                calendarName: calendar.name,
                title: item.title || "",
                dueDate: calDateToISO(item.dueDate),
                startDate: calDateToISO(item.entryDate),
                completedDate: calDateToISO(item.completedDate),
                completed,
                percentComplete: item.percentComplete || 0,
                priority,
                description: item.getProperty("DESCRIPTION") || "",
              };
            }

            async function updateTask(taskId, calendarId, title, dueDate, description, completed, percentComplete, priority) {
              if (!cal) return { error: "Calendar not available" };
              try {
                if (!taskId) return { error: "taskId is required" };
                if (!calendarId) return { error: "calendarId is required" };

                const calendar = cal.manager.getCalendars().find(c => c.id === calendarId);
                if (!calendar) return { error: `Calendar not found: ${calendarId}` };
                if (calendar.readOnly) return { error: `Calendar is read-only: ${calendar.name}` };
                if (calendar.getProperty("capabilities.tasks.supported") === false) {
                  return { error: `Calendar "${calendar.name}" does not support tasks. Use listCalendars to find one with supportsTasks=true.` };
                }

                // Try direct lookup first, then fall back to scanning all tasks
                let oldItem = null;
                if (typeof calendar.getItem === "function") {
                  try { oldItem = await calendar.getItem(taskId); } catch {}
                }
                if (!oldItem) {
                  const FILTER_TODO = 1 << 2;
                  const COMPLETED_YES = 1 << 0;
                  const COMPLETED_NO = 1 << 1;
                  let items;
                  if (typeof calendar.getItemsAsArray === "function") {
                    items = await calendar.getItemsAsArray(FILTER_TODO | COMPLETED_YES | COMPLETED_NO, 0, null, null);
                  } else {
                    items = [];
                    const stream = cal.iterate.streamValues(calendar.getItems(FILTER_TODO | COMPLETED_YES | COMPLETED_NO, 0, null, null));
                    for await (const chunk of stream) {
                      for (const i of chunk) items.push(i);
                    }
                  }
                  oldItem = items.find(i => i.id === taskId) || null;
                }
                if (!oldItem) return { error: `Task not found: ${taskId}` };

                const newItem = oldItem.clone();
                const changes = [];

                if (title !== undefined) { newItem.title = title; changes.push("title"); }
                if (description !== undefined) { newItem.setProperty("DESCRIPTION", description); changes.push("description"); }
                if (priority !== undefined) { newItem.priority = priority; changes.push("priority"); }

                if (dueDate !== undefined) {
                  // Explicit null or empty string clears the due date.
                  // Without this, `new Date(null).getTime() === 0` would
                  // silently write Unix epoch (1970-01-01) instead.
                  if (dueDate === null || dueDate === "") {
                    newItem.dueDate = null;
                  } else {
                    const js = new Date(dueDate);
                    if (isNaN(js.getTime())) return { error: `Invalid dueDate: ${dueDate}` };
                    if (/^\d{4}-\d{2}-\d{2}$/.test(dueDate.trim())) {
                      const dt = cal.createDateTime();
                      dt.resetTo(js.getFullYear(), js.getMonth(), js.getDate(), 0, 0, 0, cal.dtz.floating);
                      dt.isDate = true;
                      newItem.dueDate = dt;
                    } else {
                      newItem.dueDate = cal.dtz.jsDateToDateTime(js, cal.dtz.defaultTimezone);
                    }
                  }
                  changes.push("dueDate");
                }

                // 'completed' and 'percentComplete' both control completion state.
                // Reject ambiguous input rather than guessing precedence.
                if (completed !== undefined && percentComplete !== undefined) {
                  return { error: "Specify either 'completed' or 'percentComplete', not both" };
                }

                // Apply completion state keeping STATUS, PERCENT-COMPLETE, and
                // COMPLETED consistent per iCal RFC 5545 VTODO rules -- so
                // Thunderbird's UI and other consumers see a valid task state.
                function applyCompletionState(pct) {
                  const clamped = Math.min(100, Math.max(0, pct));
                  newItem.percentComplete = clamped;
                  if (clamped === 100) {
                    newItem.setProperty("STATUS", "COMPLETED");
                    newItem.completedDate = cal.dtz.jsDateToDateTime(new Date(), cal.dtz.defaultTimezone);
                  } else if (clamped === 0) {
                    newItem.setProperty("STATUS", "NEEDS-ACTION");
                    newItem.completedDate = null;
                  } else {
                    newItem.setProperty("STATUS", "IN-PROCESS");
                    newItem.completedDate = null;
                  }
                }

                if (percentComplete !== undefined) {
                  applyCompletionState(percentComplete);
                  changes.push("percentComplete");
                }

                if (completed !== undefined) {
                  applyCompletionState(completed ? 100 : 0);
                  changes.push("completed");
                }

                if (changes.length === 0) return { error: "No changes specified" };

                await calendar.modifyItem(newItem, oldItem);
                const result = { success: true, updated: changes, task: formatTask(newItem, calendar) };
                if (newItem.recurrenceInfo) {
                  result.warning = "This is a recurring task -- changes apply to the entire series.";
                }
                return result;
              } catch (e) {
                return { error: e.toString() };
              }
            }

            async function listEvents(calendarId, startDate, endDate, maxResults) {
              if (!cal) {
                return { error: "Calendar not available" };
              }
              try {
                const calendars = cal.manager.getCalendars();
                let targets = calendars;
                if (calendarId) {
                  const found = calendars.find(c => c.id === calendarId);
                  if (!found) return { error: `Calendar not found: ${calendarId}` };
                  targets = [found];
                }

                const startJs = startDate ? new Date(startDate) : new Date();
                if (isNaN(startJs.getTime())) return { error: `Invalid startDate: ${startDate}` };
                const endJs = endDate ? new Date(endDate) : new Date(startJs.getTime() + 30 * 86400000);
                if (isNaN(endJs.getTime())) return { error: `Invalid endDate: ${endDate}` };

                const rangeStart = cal.dtz.jsDateToDateTime(startJs, cal.dtz.defaultTimezone);
                const rangeEnd = cal.dtz.jsDateToDateTime(endJs, cal.dtz.defaultTimezone);
                const limit = Math.min(Math.max(maxResults || 100, 1), 500);

                // Query with date range and occurrence expansion
                const FILTER_EVENT = 1 << 3;
                const FILTER_OCCURRENCES = 1 << 4;
                const results = [];
                for (const calendar of targets) {
                  let items;
                  try {
                    // Try with occurrence expansion first (works on most providers)
                    if (typeof calendar.getItemsAsArray === "function") {
                      items = await calendar.getItemsAsArray(FILTER_EVENT | FILTER_OCCURRENCES, 0, rangeStart, rangeEnd);
                    } else {
                      items = [];
                      const stream = cal.iterate.streamValues(calendar.getItems(FILTER_EVENT | FILTER_OCCURRENCES, 0, rangeStart, rangeEnd));
                      for await (const chunk of stream) {
                        for (const i of chunk) items.push(i);
                      }
                    }
                  } catch {
                    // Fallback: fetch without occurrence expansion, expand manually
                    items = await getCalendarItems(calendar, rangeStart, rangeEnd);
                  }

                  const EVENT_COLLECTION_CAP = limit * 10; // Safety cap for recurring event expansion
                  for (const item of items) {
                    if (results.length >= EVENT_COLLECTION_CAP) break;
                    // If we got base recurring events (fallback path), expand them
                    if (item.recurrenceInfo) {
                      try {
                        const occurrences = item.recurrenceInfo.getOccurrences(rangeStart, rangeEnd, 0);
                        for (const occ of occurrences) {
                          if (results.length >= EVENT_COLLECTION_CAP) break;
                          results.push(formatEvent(occ, calendar));
                        }
                      } catch {
                        results.push(formatEvent(item, calendar));
                      }
                    } else {
                      results.push(formatEvent(item, calendar));
                    }
                  }
                }

                results.sort((a, b) => new Date(a.startDate) - new Date(b.startDate));
                return results.slice(0, limit);
              } catch (e) {
                return { error: e.toString() };
              }
            }

            async function listTasks(calendarId, completed, dueBefore, maxResults) {
              if (!cal) return { error: "Calendar not available" };
              try {
                const calendars = cal.manager.getCalendars();
                let targets = calendars.filter(c =>
                  c.getProperty("capabilities.tasks.supported") !== false
                );
                if (calendarId) {
                  const found = calendars.find(c => c.id === calendarId);
                  if (!found) return { error: `Calendar not found: ${calendarId}` };
                  if (found.getProperty("capabilities.tasks.supported") === false) {
                    return { error: `Calendar "${found.name}" does not support tasks` };
                  }
                  targets = [found];
                }

                let dueBeforeDt = null;
                if (dueBefore) {
                  const js = new Date(dueBefore);
                  if (isNaN(js.getTime())) return { error: `Invalid dueBefore: ${dueBefore}` };
                  dueBeforeDt = js;
                }

                const limit = Math.min(Math.max(maxResults || 100, 1), 500);
                // Thunderbird calICalendar filter bits:
                // TYPE_TODO = 1<<2, COMPLETED_YES = 1<<0, COMPLETED_NO = 1<<1
                const FILTER_TODO = 1 << 2;
                const COMPLETED_YES = 1 << 0;
                const COMPLETED_NO = 1 << 1;
                let itemFilter = FILTER_TODO;
                if (completed === true) {
                  itemFilter |= COMPLETED_YES;
                } else if (completed === false) {
                  itemFilter |= COMPLETED_NO;
                } else {
                  itemFilter |= COMPLETED_YES | COMPLETED_NO;
                }
                const TASK_COLLECTION_CAP = limit * 10;
                const results = [];

                for (const calendar of targets) {
                  let items;
                  try {
                    if (typeof calendar.getItemsAsArray === "function") {
                      items = await calendar.getItemsAsArray(itemFilter, 0, null, null);
                    } else {
                      items = [];
                      const stream = cal.iterate.streamValues(calendar.getItems(itemFilter, 0, null, null));
                      for await (const chunk of stream) {
                        for (const i of chunk) items.push(i);
                      }
                    }
                  } catch {
                    continue; // Skip calendars that fail to query
                  }

                  for (const item of items) {
                    if (results.length >= TASK_COLLECTION_CAP) break;
                    // Filter by due date -- exclude undated tasks when dueBefore is set
                    if (dueBeforeDt) {
                      if (!item.dueDate) continue;
                      try {
                        const due = new Date(item.dueDate.nativeTime / 1000);
                        if (due >= dueBeforeDt) continue;
                      } catch { /* include if we can't parse */ }
                    }
                    results.push(formatTask(item, calendar));
                  }
                }

                // Sort by dueDate (nulls last), then title
                results.sort((a, b) => {
                  if (a.dueDate && b.dueDate) return new Date(a.dueDate) - new Date(b.dueDate);
                  if (a.dueDate) return -1;
                  if (b.dueDate) return 1;
                  return (a.title || "").localeCompare(b.title || "");
                });
                return results.slice(0, limit);
              } catch (e) {
                return { error: e.toString() };
              }
            }

            async function updateEvent(eventId, calendarId, title, startDate, endDate, location, description) {
              if (!cal) return { error: "Calendar not available" };
              try {
                if (!eventId) return { error: "eventId is required" };
                if (!calendarId) return { error: "calendarId is required" };

                const calendar = cal.manager.getCalendars().find(c => c.id === calendarId);
                if (!calendar) return { error: `Calendar not found: ${calendarId}` };
                if (calendar.readOnly) return { error: `Calendar is read-only: ${calendar.name}` };

                // Use getItem API if available, else scan
                let oldItem = null;
                if (typeof calendar.getItem === "function") {
                  try { oldItem = await calendar.getItem(eventId); } catch {}
                }
                if (!oldItem) {
                  // Fallback: scan all events
                  const all = await getCalendarItems(calendar, null, null);
                  oldItem = all.find(i => i.id === eventId) || null;
                }
                if (!oldItem) return { error: `Event not found: ${eventId}` };

                const newItem = oldItem.clone();
                const changes = [];

                if (title !== undefined) { newItem.title = title; changes.push("title"); }

                if (startDate !== undefined) {
                  const js = new Date(startDate);
                  if (isNaN(js.getTime())) return { error: `Invalid startDate: ${startDate}` };
                  if (newItem.startDate && newItem.startDate.isDate) {
                    const dt = cal.createDateTime();
                    dt.resetTo(js.getFullYear(), js.getMonth(), js.getDate(), 0, 0, 0, cal.dtz.floating);
                    dt.isDate = true;
                    newItem.startDate = dt;
                  } else {
                    newItem.startDate = cal.dtz.jsDateToDateTime(js, cal.dtz.defaultTimezone);
                  }
                  changes.push("startDate");
                }

                if (endDate !== undefined) {
                  const js = new Date(endDate);
                  if (isNaN(js.getTime())) return { error: `Invalid endDate: ${endDate}` };
                  if (newItem.endDate && newItem.endDate.isDate) {
                    const dt = cal.createDateTime();
                    // iCal DTEND is exclusive for all-day -- bump by 1 day
                    const next = new Date(js.getFullYear(), js.getMonth(), js.getDate());
                    next.setDate(next.getDate() + 1);
                    dt.resetTo(next.getFullYear(), next.getMonth(), next.getDate(), 0, 0, 0, cal.dtz.floating);
                    dt.isDate = true;
                    newItem.endDate = dt;
                  } else {
                    newItem.endDate = cal.dtz.jsDateToDateTime(js, cal.dtz.defaultTimezone);
                  }
                  changes.push("endDate");
                }

                if (location !== undefined) { newItem.setProperty("LOCATION", location); changes.push("location"); }
                if (description !== undefined) { newItem.setProperty("DESCRIPTION", description); changes.push("description"); }

                if (changes.length === 0) return { error: "No changes specified" };

                // Validate end > start after all changes
                if (newItem.startDate && newItem.endDate && newItem.endDate.compare(newItem.startDate) <= 0) {
                  return { error: "endDate must be after startDate" };
                }

                await calendar.modifyItem(newItem, oldItem);
                const result = { success: true, updated: changes };
                if (oldItem.recurrenceInfo) {
                  result.warning = "This is a recurring event -- changes apply to the entire series.";
                }
                return result;
              } catch (e) {
                return { error: e.toString() };
              }
            }

            async function deleteEvent(eventId, calendarId) {
              if (!cal) return { error: "Calendar not available" };
              try {
                if (!eventId) return { error: "eventId is required" };
                if (!calendarId) return { error: "calendarId is required" };

                const calendar = cal.manager.getCalendars().find(c => c.id === calendarId);
                if (!calendar) return { error: `Calendar not found: ${calendarId}` };
                if (calendar.readOnly) return { error: `Calendar is read-only: ${calendar.name}` };

                let item = null;
                if (typeof calendar.getItem === "function") {
                  try { item = await calendar.getItem(eventId); } catch {}
                }
                if (!item) {
                  const all = await getCalendarItems(calendar, null, null);
                  item = all.find(i => i.id === eventId) || null;
                }
                if (!item) return { error: `Event not found: ${eventId}` };

                const isRecurring = !!item.recurrenceInfo;
                await calendar.deleteItem(item);
                const result = { success: true, deleted: eventId };
                if (isRecurring) {
                  result.warning = "This was a recurring event -- the entire series was deleted.";
                }
                return result;
              } catch (e) {
                return { error: e.toString() };
              }
            }

            function createTask(title, dueDate, calendarId) {
              if (!cal || !CalTodo) return { error: "Calendar module not available" };
              try {
                const win = Services.wm.getMostRecentWindow("mail:3pane");
                if (!win) return { error: "No Thunderbird window found" };

                let dueDt = null;
                if (dueDate) {
                  const js = new Date(dueDate);
                  if (isNaN(js.getTime())) return { error: `Invalid dueDate: ${dueDate}` };
                  // Date-only string (YYYY-MM-DD) means all-day
                  if (/^\d{4}-\d{2}-\d{2}$/.test(dueDate.trim())) {
                    dueDt = cal.createDateTime();
                    dueDt.resetTo(js.getFullYear(), js.getMonth(), js.getDate(), 0, 0, 0, cal.dtz.floating);
                    dueDt.isDate = true;
                  } else {
                    dueDt = cal.dtz.jsDateToDateTime(js, cal.dtz.defaultTimezone);
                  }
                }

                // Find target calendar (must support tasks)
                let targetCalendar = null;
                if (calendarId) {
                  targetCalendar = cal.manager.getCalendars().find(c => c.id === calendarId);
                  if (!targetCalendar) return { error: `Calendar not found: ${calendarId}` };
                  if (targetCalendar.readOnly) return { error: `Calendar is read-only: ${targetCalendar.name}` };
                  if (targetCalendar.getProperty("capabilities.tasks.supported") === false) {
                    return { error: `Calendar "${targetCalendar.name}" does not support tasks. Use listCalendars to find one with supportsTasks=true.` };
                  }
                }

                // Cross-context CalTodo objects cause silent save failure in dialog.
                // Pass title as summary param; TB creates its own CalTodo internally.
                win.createTodoWithDialog(targetCalendar, dueDt, title, null);

                return { success: true, message: `Task dialog opened for "${title}"` };
              } catch (e) {
                return { error: e.toString() };
              }
            }

	            function getMessage(messageId, folderPath, saveAttachments, bodyFormat, rawSource) {
	              return new Promise((resolve) => {
	                try {
	                  const found = findMessage(messageId, folderPath);
	                  if (found.error) {
	                    resolve({ error: found.error });
	                    return;
	                  }
	                  const { msgHdr } = found;

	                  // Raw source mode: return full RFC 2822 message
	                  if (rawSource) {
	                    let stream = null;
	                    try {
	                      const folder = msgHdr.folder;
	                      stream = folder.getMsgInputStream(msgHdr, {});
	                      let messageSize = folder.hasMsgOffline(msgHdr.messageKey)
	                        ? msgHdr.offlineMessageSize
	                        : msgHdr.messageSize;
	                      // For local folders (mbox), messageSize can be 0 or
	                      // inaccurate for imported messages. Fall back to reading
	                      // whatever is available in the stream.
	                      if (!messageSize || messageSize <= 0) {
	                        messageSize = stream.available();
	                      }
	                      if (!messageSize || messageSize <= 0) {
	                        resolve({ error: "Message has zero size - cannot read raw source" });
	                        return;
	                      }
	                      // No charset specified -- defaults to Latin-1 which
	                      // preserves raw bytes. UTF-8 would corrupt messages
	                      // with 8-bit content.
	                      const raw = NetUtil.readInputStreamToString(stream, messageSize);
	                      resolve({
	                        id: msgHdr.messageId,
	                        subject: msgHdr.mime2DecodedSubject || msgHdr.subject,
	                        rawSource: raw,
	                      });
	                    } catch (e) {
	                      resolve({ error: `Failed to read raw source: ${e}` });
	                    } finally {
	                      if (stream) try { stream.close(); } catch { /* ignore */ }
	                    }
	                    return;
	                  }

	                  const { MsgHdrToMimeMessage } = ChromeUtils.importESModule(
	                    "resource:///modules/gloda/MimeMessage.sys.mjs"
	                  );

                  MsgHdrToMimeMessage(msgHdr, null, (aMsgHdr, aMimeMsg) => {
                    if (!aMimeMsg) {
                      resolve({ error: "Could not parse message" });
                      return;
                    }

                    const requestedBodyFormat = bodyFormat || "markdown";
                    const fmt = extractFormattedBody(aMimeMsg, requestedBodyFormat);
                    let body = fmt.body;
                    let bodyIsHtml = fmt.bodyIsHtml;
                    // If structured MIME extraction failed, try raw stream
                    // fallback for local mbox folders where MsgHdrToMimeMessage
                    // returns empty body parts.
                    if (!body) {
                      const fallbackContext = `thunderbird-mcp: raw singlepart body fallback (${msgHdr.messageId})`;
                      let rawStream = null;
                      try {
                        const rawFolder = msgHdr.folder;
                        rawStream = rawFolder.getMsgInputStream(msgHdr, {});
                        let rawSize = rawFolder.hasMsgOffline(msgHdr.messageKey)
                          ? msgHdr.offlineMessageSize
                          : msgHdr.messageSize;
                        if (!rawSize || rawSize <= 0) rawSize = rawStream.available();
                        if (!rawSize || rawSize <= 0) {
                          console.error(`${fallbackContext}: message stream has zero size`);
                        } else {
                          // No charset specified -- defaults to Latin-1 which
                          // preserves raw bytes for later transfer decoding.
                          const rawContent = NetUtil.readInputStreamToString(rawStream, rawSize);
                          let headerEnd = rawContent.indexOf("\r\n\r\n");
                          let bodyStart = headerEnd >= 0 ? headerEnd + 4 : -1;
                          if (headerEnd < 0) {
                            headerEnd = rawContent.indexOf("\n\n");
                            bodyStart = headerEnd >= 0 ? headerEnd + 2 : -1;
                          }
                          if (bodyStart < 0) {
                            console.error(`${fallbackContext}: could not find header/body boundary`);
                          } else {
                            const headerBlock = rawContent.slice(0, headerEnd);
                            const rawBody = rawContent.slice(bodyStart);
                            const unfoldedHeaders = headerBlock
                              .replace(/\r\n[ \t]+/g, " ")
                              .replace(/\n[ \t]+/g, " ");
                            let contentTypeHeader = "";
                            let transferEncodingHeader = "";
                            for (const line of unfoldedHeaders.split(/\r?\n/)) {
                              const colonIdx = line.indexOf(":");
                              if (colonIdx < 0) continue;
                              const headerName = line.slice(0, colonIdx).trim().toLowerCase();
                              const headerValue = line.slice(colonIdx + 1).trim();
                              if (headerName === "content-type" && !contentTypeHeader) {
                                contentTypeHeader = headerValue;
                              } else if (headerName === "content-transfer-encoding" && !transferEncodingHeader) {
                                transferEncodingHeader = headerValue;
                              }
                            }
                            const contentTypeValue = contentTypeHeader || "text/plain";
                            const contentType = (contentTypeValue.split(";")[0] || "text/plain").trim().toLowerCase();
                            if (contentType.startsWith("multipart/")) {
                              console.error(`${fallbackContext}: multipart top-level content-type not supported (${contentType})`);
                            } else if (contentType !== "text/plain" && contentType !== "text/html") {
                              console.error(`${fallbackContext}: unsupported top-level content-type "${contentType || "(missing)"}"`);
                            } else {
                              const charsetMatch = contentTypeValue.match(/(?:^|;)\s*charset\s*=\s*(?:"([^"]+)"|'([^']+)'|([^;\s]+))/i);
                              const charset = (charsetMatch?.[1] || charsetMatch?.[2] || charsetMatch?.[3] || "utf-8").trim();
                              const transferEncoding = ((transferEncodingHeader.split(";")[0] || "7bit").trim().toLowerCase() || "7bit");
                              let bodyBytes = null;

                              if (transferEncoding === "quoted-printable") {
                                const qpBody = rawBody.replace(/=\r\n/g, "").replace(/=\n/g, "");
                                const decodedBytes = [];
                                for (let i = 0; i < qpBody.length; i++) {
                                  if (qpBody[i] === "=" && i + 2 < qpBody.length) {
                                    const hex = qpBody.slice(i + 1, i + 3);
                                    if (/^[0-9A-Fa-f]{2}$/.test(hex)) {
                                      decodedBytes.push(parseInt(hex, 16));
                                      i += 2;
                                      continue;
                                    }
                                  }
                                  decodedBytes.push(qpBody.charCodeAt(i) & 0xFF);
                                }
                                bodyBytes = new Uint8Array(decodedBytes);
                              } else if (transferEncoding === "base64") {
                                try {
                                  const binary = atob(rawBody.replace(/\s/g, ""));
                                  bodyBytes = new Uint8Array(binary.length);
                                  for (let i = 0; i < binary.length; i++) {
                                    bodyBytes[i] = binary.charCodeAt(i) & 0xFF;
                                  }
                                } catch (e) {
                                  console.error(`${fallbackContext}: invalid base64 body`, e);
                                }
                              } else if (transferEncoding === "7bit" || transferEncoding === "8bit" || transferEncoding === "binary") {
                                bodyBytes = new Uint8Array(rawBody.length);
                                for (let i = 0; i < rawBody.length; i++) {
                                  bodyBytes[i] = rawBody.charCodeAt(i) & 0xFF;
                                }
                              } else {
                                console.error(`${fallbackContext}: unsupported content-transfer-encoding "${transferEncoding}"`);
                              }

                              if (bodyBytes) {
                                let decodedBody;
                                try {
                                  decodedBody = new TextDecoder(charset, { fatal: false }).decode(bodyBytes);
                                } catch (e) {
                                  if (!(e instanceof RangeError) && e?.name !== "RangeError") throw e;
                                  console.error(`${fallbackContext}: unknown charset "${charset}", retrying with utf-8`);
                                  decodedBody = new TextDecoder("utf-8", { fatal: false }).decode(bodyBytes);
                                }

                                if (contentType === "text/html") {
                                  if (requestedBodyFormat === "html") {
                                    body = decodedBody;
                                    bodyIsHtml = true;
                                  } else if (requestedBodyFormat === "markdown") {
                                    body = htmlToMarkdown(decodedBody);
                                    bodyIsHtml = false;
                                  } else {
                                    body = stripHtml(decodedBody);
                                    bodyIsHtml = false;
                                  }
                                } else {
                                  body = decodedBody;
                                  bodyIsHtml = false;
                                }
                              }
                            }
                          }
                        }
                      } catch (e) {
                        console.error(`${fallbackContext}: failed`, e);
                      } finally {
                        if (rawStream) try { rawStream.close(); } catch (e) {
                          console.error(`${fallbackContext}: failed to close stream`, e);
                        }
                      }
                    }

                    // Always collect attachment metadata
                    const attachments = [];
                    const attachmentSources = [];
                    if (aMimeMsg && aMimeMsg.allUserAttachments) {
                      for (const att of aMimeMsg.allUserAttachments) {
                        const info = {
                          name: att?.name || "",
                          contentType: att?.contentType || "",
                          size: typeof att?.size === "number" ? att.size : null,
                          isInline: false,
                        };
                        attachments.push(info);
                        attachmentSources.push({
                          info,
                          url: att?.url || "",
                          size: typeof att?.size === "number" ? att.size : null
                        });
                      }
                    }

                    // Find inline CID images not included in allUserAttachments.
                    // Gloda's MimeMessage strips content-id headers, so we identify
                    // inline images by: image/* parts inside multipart/related that
                    // aren't already in allUserAttachments. URLs are resolved via
                    // MailServices.messageServiceFromURI (imap-message:// isn't
                    // directly fetchable by NetUtil).
                    if (aMimeMsg) {
                      const existingPartNames = new Set(attachments.map(a => a.partName).filter(Boolean));
                      function collectInlineImages(part, insideRelated, results) {
                        const ct = ((part.contentType || "").split(";")[0] || "").trim().toLowerCase();
                        // Skip nested messages -- their inline images are not ours
                        if (ct === "message/rfc822") return;
                        if (ct === "multipart/related") insideRelated = true;
                        if (insideRelated && ct.startsWith("image/") && part.partName) {
                          // Deduplicate by partName (stable ID), not filename (can collide)
                          if (existingPartNames.has(part.partName)) return;
                          existingPartNames.add(part.partName);
                          // Extract filename from headers (contentType field lacks params)
                          const ctHeader = part.headers?.["content-type"]?.[0] || "";
                          const nameMatch = ctHeader.match(/name\s*=\s*"?([^";]+)"?/i);
                          const name = nameMatch ? nameMatch[1] : `inline_${part.partName}`;
                          results.push({ part, name, ct });
                        }
                        if (part.parts) {
                          for (const sub of part.parts) collectInlineImages(sub, insideRelated, results);
                        }
                      }
                      const inlineImages = [];
                      collectInlineImages(aMimeMsg, false, inlineImages);
                      if (inlineImages.length > 0) {
                        const msgUri = msgHdr.folder.getUriForMsg(msgHdr);
                        for (const { part, name, ct } of inlineImages) {
                          // Resolve to a fetchable URL via the message service
                          let partUrl = "";
                          try {
                            const svc = MailServices.messageServiceFromURI(msgUri);
                            const baseUri = svc.getUrlForUri(msgUri);
                            // Append part parameter to the resolved fetchable URL
                            const sep = baseUri.spec.includes("?") ? "&" : "?";
                            partUrl = `${baseUri.spec}${sep}part=${part.partName}`;
                          } catch {
                            partUrl = "";
                          }
                          const info = {
                            name,
                            contentType: ct,
                            size: typeof part.size === "number" && part.size > 0 ? part.size : null,
                            partName: part.partName,
                            isInline: true,
                          };
                          attachments.push(info);
                          if (partUrl) {
                            attachmentSources.push({ info, url: partUrl, size: info.size });
                          }
                        }
                      }
                    }

                    const msgTags = getUserTags(msgHdr);
                    const baseResponse = {
                      id: msgHdr.messageId,
                      subject: msgHdr.mime2DecodedSubject || msgHdr.subject,
                      author: msgHdr.mime2DecodedAuthor || msgHdr.author,
                      recipients: msgHdr.mime2DecodedRecipients || msgHdr.recipients,
                      ccList: msgHdr.ccList,
                      date: msgHdr.date ? new Date(msgHdr.date / 1000).toISOString() : null,
                      tags: msgTags,
                      body,
                      bodyIsHtml,
                      attachments
                    };

                    if (!saveAttachments || attachmentSources.length === 0) {
                      resolve(baseResponse);
                      return;
                    }

                    function sanitizePathSegment(s) {
                      const sanitized = String(s || "").replace(/[^a-zA-Z0-9]/g, "_");
                      return sanitized || "message";
                    }

                    function sanitizeFilename(s) {
                      let name = String(s || "").trim();
                      if (!name) name = "attachment";
                      name = name.replace(/[^a-zA-Z0-9._-]/g, "_");
                      name = name.replace(/^_+/, "").replace(/_+$/, "");
                      return name || "attachment";
                    }

                    function ensureAttachmentDir(sanitizedId) {
                      const root = Services.dirsvc.get("TmpD", Ci.nsIFile);
                      root.append("thunderbird-mcp");
                      try {
                        root.create(Ci.nsIFile.DIRECTORY_TYPE, 0o700);
                      } catch (e) {
                        if (!root.exists() || !root.isDirectory()) throw e;
                        // already exists, fine
                      }
                      const dir = root.clone();
                      dir.append(sanitizedId);
                      try {
                        dir.create(Ci.nsIFile.DIRECTORY_TYPE, 0o700);
                      } catch (e) {
                        if (!dir.exists() || !dir.isDirectory()) throw e;
                        // already exists, fine
                      }
                      return dir;
                    }

                    const sanitizedId = sanitizePathSegment(messageId);
                    let dir;
                    try {
                      dir = ensureAttachmentDir(sanitizedId);
                    } catch (e) {
                      for (const { info } of attachmentSources) {
                        info.error = `Failed to create attachment directory: ${e}`;
                      }
                      resolve(baseResponse);
                      return;
                    }

                    const MAX_ATTACHMENT_BYTES = 50 * 1024 * 1024;

                    const saveOne = ({ info, url, size }, index) =>
                      new Promise((done) => {
                        try {
                          if (!url) {
                            info.error = "Missing attachment URL";
                            done();
                            return;
                          }

                          const knownSize = typeof size === "number" ? size : null;
                          if (knownSize !== null && knownSize > MAX_ATTACHMENT_BYTES) {
                            info.error = `Attachment too large (${knownSize} bytes, limit ${MAX_ATTACHMENT_BYTES})`;
                            done();
                            return;
                          }

                          const idx = typeof index === "number" && Number.isFinite(index) ? index : 0;
                          let safeName = sanitizeFilename(info.name);
                          if (!safeName || safeName === "." || safeName === "..") {
                            safeName = `attachment_${idx}`;
                          }
                          const file = dir.clone();
                          file.append(safeName);

                          try {
                            file.createUnique(Ci.nsIFile.NORMAL_FILE_TYPE, 0o600);
                          } catch (e) {
                            info.error = `Failed to create file: ${e}`;
                            done();
                            return;
                          }

                          const channel = NetUtil.newChannel({
                            uri: url,
                            loadUsingSystemPrincipal: true
                          });

                          NetUtil.asyncFetch(channel, (inputStream, status, request) => {
                            try {
                              if (status && status !== 0) {
                                try { inputStream?.close(); } catch {}
                                info.error = `Fetch failed: ${status}`;
                                try { file.remove(false); } catch {}
                                done();
                                return;
                              }
                              if (!inputStream) {
                                info.error = "Fetch returned no data";
                                try { file.remove(false); } catch {}
                                done();
                                return;
                              }

                              try {
                                const reqLen = request && typeof request.contentLength === "number" ? request.contentLength : -1;
                                if (reqLen >= 0 && reqLen > MAX_ATTACHMENT_BYTES) {
                                  try { inputStream.close(); } catch {}
                                  info.error = `Attachment too large (${reqLen} bytes, limit ${MAX_ATTACHMENT_BYTES})`;
                                  try { file.remove(false); } catch {}
                                  done();
                                  return;
                                }
                              } catch {
                                // ignore contentLength failures
                              }

                              const ostream = Cc["@mozilla.org/network/file-output-stream;1"]
                                .createInstance(Ci.nsIFileOutputStream);
                              ostream.init(file, -1, -1, 0);

                              NetUtil.asyncCopy(inputStream, ostream, (copyStatus) => {
                                try {
                                  if (copyStatus && copyStatus !== 0) {
                                    info.error = `Write failed: ${copyStatus}`;
                                    try { file.remove(false); } catch {}
                                    done();
                                    return;
                                  }

                                  try {
                                    if (file.fileSize > MAX_ATTACHMENT_BYTES) {
                                      info.error = `Attachment too large (${file.fileSize} bytes, limit ${MAX_ATTACHMENT_BYTES})`;
                                      try { file.remove(false); } catch {}
                                      done();
                                      return;
                                    }
                                  } catch {
                                    // ignore fileSize failures
                                  }

                                  info.filePath = file.path;
                                  done();
                                } catch (e) {
                                  info.error = `Write failed: ${e}`;
                                  try { file.remove(false); } catch {}
                                  done();
                                }
                              });
                            } catch (e) {
                              info.error = `Fetch failed: ${e}`;
                              try { file.remove(false); } catch {}
                              done();
                            }
                          });
                        } catch (e) {
                          info.error = String(e);
                          done();
                        }
                      });

                    (async () => {
                      try {
                        await Promise.all(attachmentSources.map((src, i) => saveOne(src, i)));
                      } catch (e) {
                        // Per-attachment errors are handled; this is just a safeguard.
                        for (const { info } of attachmentSources) {
                          if (!info.error) info.error = `Unexpected save error: ${e}`;
                        }
                      }
                      resolve(baseResponse);
                    })();
                  }, true, { examineEncryptedParts: true });

	                } catch (e) {
	                  resolve({ error: e.toString() });
	                }
	              });
	            }

            /**
             * Composes a new email. Opens a compose window for review, or sends
             * directly when skipReview is true.
             *
             * HTML body handling quirks:
             * 1. Strip newlines from HTML - Thunderbird adds <br> for each \n
             * 2. Encode non-ASCII as HTML entities - compose window has charset issues
             *    with emojis/unicode even with <meta charset="UTF-8">
             */
            function composeMail(to, subject, body, cc, bcc, isHtml, from, attachments, skipReview) {
              try {
                const msgComposeParams = Cc["@mozilla.org/messengercompose/composeparams;1"]
                  .createInstance(Ci.nsIMsgComposeParams);

                const composeFields = Cc["@mozilla.org/messengercompose/composefields;1"]
                  .createInstance(Ci.nsIMsgCompFields);

                composeFields.to = to || "";
                composeFields.cc = cc || "";
                composeFields.bcc = bcc || "";
                composeFields.subject = subject || "";

                const formatted = formatBodyHtml(body, isHtml);
                if (isHtml && formatted.includes('<html')) {
                  composeFields.body = formatted;
                } else {
                  composeFields.body = `<html><head><meta charset="UTF-8"></head><body>${formatted}</body></html>`;
                }

                msgComposeParams.type = Ci.nsIMsgCompType.New;
                msgComposeParams.format = Ci.nsIMsgCompFormat.HTML;
                msgComposeParams.composeFields = composeFields;

                const identityResult = setComposeIdentity(msgComposeParams, from, null);
                if (identityResult && identityResult.error) return identityResult;

                const { descs: fileDescs, failed: failedPaths } = filePathsToAttachDescs(attachments);

                if (skipReview) {
                  return sendMessageDirectly(composeFields, msgComposeParams.identity, fileDescs, null, Ci.nsIMsgCompType.New).then(result => {
                    if (result.success) {
                      let msg = "Message sent";
                      if (failedPaths.length > 0) msg += ` (failed to attach: ${failedPaths.join(", ")})`;
                      result.message = msg;
                    }
                    return result;
                  });
                }

                const msgComposeService = Cc["@mozilla.org/messengercompose;1"]
                  .getService(Ci.nsIMsgComposeService);
                msgComposeService.OpenComposeWindowWithParams(null, msgComposeParams);

                injectAttachmentsAsync(fileDescs);

                let msg = "Compose window opened";
                if (failedPaths.length > 0) msg += ` (failed to attach: ${failedPaths.join(", ")})`;
                return { success: true, message: msg };
              } catch (e) {
                return { error: e.toString() };
              }
            }

            /**
             * Replies to a message with quoted original. Opens a compose window
             * for review, or sends directly when skipReview is true.
             *
             * Uses nsIMsgCompType.New to preserve our body content, then manually
             * builds the quoted original message text. Threading is maintained
             * via the References and In-Reply-To headers.
             */
	            function replyToMessage(messageId, folderPath, body, replyAll, isHtml, to, cc, bcc, from, attachments, skipReview) {
	              return new Promise((resolve) => {
	                try {
	                  const found = findMessage(messageId, folderPath);
	                  if (found.error) {
	                    resolve({ error: found.error });
	                    return;
	                  }
	                  const { msgHdr, folder } = found;

	                  // Fetch original message body for quoting
	                  const { MsgHdrToMimeMessage } = ChromeUtils.importESModule(
	                    "resource:///modules/gloda/MimeMessage.sys.mjs"
                  );

                  MsgHdrToMimeMessage(msgHdr, null, (aMsgHdr, aMimeMsg) => {
                    try {
                      const originalBody = extractPlainTextBody(aMimeMsg);

                      const msgComposeParams = Cc["@mozilla.org/messengercompose/composeparams;1"]
                        .createInstance(Ci.nsIMsgComposeParams);

                      const composeFields = Cc["@mozilla.org/messengercompose/composefields;1"]
                        .createInstance(Ci.nsIMsgCompFields);

                      if (replyAll) {
                        composeFields.to = to || msgHdr.author;
                        // Combine original recipients and CC, filter out own address
                        // Split on commas not inside quotes to handle "Last, First" <email>
                        const splitAddresses = (s) => (s || "").match(/(?:[^,"]|"[^"]*")+/g) || [];
                        const extractEmail = (s) => (s.match(/<([^>]+)>/)?.[1] || s.trim()).toLowerCase();
                        // Get all own emails (default + aliases) for accurate self-filtering
                        const ownAccount = MailServices.accounts.findAccountForServer(folder.server);
                        const ownEmails = new Set();
                        if (ownAccount) {
                          for (const identity of ownAccount.identities) {
                            if (identity.email) ownEmails.add(identity.email.toLowerCase());
                          }
                        }
                        const allRecipients = [
                          ...splitAddresses(msgHdr.recipients),
                          ...splitAddresses(msgHdr.ccList)
                        ]
                          .map(r => r.trim())
                          .filter(r => r && (ownEmails.size === 0 || !ownEmails.has(extractEmail(r))));
                        // Deduplicate by email address
                        const seen = new Set();
                        const uniqueRecipients = allRecipients.filter(r => {
                          const email = extractEmail(r);
                          if (seen.has(email)) return false;
                          seen.add(email);
                          return true;
                        });
                        if (cc) {
                          composeFields.cc = cc;
                        } else if (uniqueRecipients.length > 0) {
                          composeFields.cc = uniqueRecipients.join(", ");
                        }
                      } else {
                        composeFields.to = to || msgHdr.author;
                        if (cc) composeFields.cc = cc;
                      }

                      composeFields.bcc = bcc || "";

                      const origSubject = msgHdr.mime2DecodedSubject || msgHdr.subject || "";
                      composeFields.subject = /^re:/i.test(origSubject) ? origSubject : `Re: ${origSubject}`;

                      // Threading headers
                      composeFields.references = `<${messageId}>`;
                      composeFields.setHeader("In-Reply-To", `<${messageId}>`);

                      // Build quoted text block
                      const dateStr = msgHdr.date ? new Date(msgHdr.date / 1000).toLocaleString() : "";
                      const author = msgHdr.mime2DecodedAuthor || msgHdr.author || "";
                      const quotedLines = originalBody.split('\n').map(line =>
                        `&gt; ${escapeHtml(line)}`
                      ).join('<br>');
                      const quoteBlock = `<br><br>On ${dateStr}, ${escapeHtml(author)} wrote:<br>${quotedLines}`;

                      composeFields.body = `<html><head><meta charset="UTF-8"></head><body>${formatBodyHtml(body, isHtml)}${quoteBlock}</body></html>`;

                      const { descs: fileDescs, failed: failedPaths } = filePathsToAttachDescs(attachments);

                      msgComposeParams.type = Ci.nsIMsgCompType.New;
                      msgComposeParams.format = Ci.nsIMsgCompFormat.HTML;
                      msgComposeParams.composeFields = composeFields;

                      const identityResult = setComposeIdentity(msgComposeParams, from, folder.server);
                      if (identityResult && identityResult.error) { resolve(identityResult); return; }

                      if (skipReview) {
                        const msgURI = folder.getUriForMsg(msgHdr);
                        const compType = replyAll ? Ci.nsIMsgCompType.ReplyAll : Ci.nsIMsgCompType.Reply;
                        sendMessageDirectly(composeFields, msgComposeParams.identity, fileDescs, msgURI, compType).then(result => {
                          if (result.success) {
                            let msg = "Reply sent";
                            if (failedPaths.length > 0) msg += ` (failed to attach: ${failedPaths.join(", ")})`;
                            result.message = msg;
                          }
                          resolve(result);
                        });
                        return;
                      }

                      const msgComposeService = Cc["@mozilla.org/messengercompose;1"]
                        .getService(Ci.nsIMsgComposeService);
                      msgComposeService.OpenComposeWindowWithParams(null, msgComposeParams);

                      injectAttachmentsAsync(fileDescs);

                      let msg = "Reply window opened";
                      if (failedPaths.length > 0) msg += ` (failed to attach: ${failedPaths.join(", ")})`;
                      resolve({ success: true, message: msg });
                    } catch (e) {
                      resolve({ error: e.toString() });
                    }
                  }, true, { examineEncryptedParts: true });

                } catch (e) {
                  resolve({ error: e.toString() });
                }
              });
            }

            /**
             * Forwards a message with original content and attachments. Opens a
             * compose window for review, or sends directly when skipReview is true.
             * Uses New type with manual forward quote to preserve both intro body and forwarded content.
             */
	            function forwardMessage(messageId, folderPath, to, body, isHtml, cc, bcc, from, attachments, skipReview) {
	              return new Promise((resolve) => {
	                try {
	                  const found = findMessage(messageId, folderPath);
	                  if (found.error) {
	                    resolve({ error: found.error });
	                    return;
	                  }
	                  const { msgHdr, folder } = found;

	                  // Get attachments and body from original message
	                  const { MsgHdrToMimeMessage } = ChromeUtils.importESModule(
	                    "resource:///modules/gloda/MimeMessage.sys.mjs"
                  );

                  MsgHdrToMimeMessage(msgHdr, null, (aMsgHdr, aMimeMsg) => {
                    try {
                      const msgComposeParams = Cc["@mozilla.org/messengercompose/composeparams;1"]
                        .createInstance(Ci.nsIMsgComposeParams);

                      const composeFields = Cc["@mozilla.org/messengercompose/composefields;1"]
                        .createInstance(Ci.nsIMsgCompFields);

                      composeFields.to = to;
                      composeFields.cc = cc || "";
                      composeFields.bcc = bcc || "";

                      const origSubject = msgHdr.mime2DecodedSubject || msgHdr.subject || "";
                      composeFields.subject = /^fwd:/i.test(origSubject) ? origSubject : `Fwd: ${origSubject}`;

                      // Get original body
                      const originalBody = extractPlainTextBody(aMimeMsg);

                      // Build forward header block
                      const dateStr = msgHdr.date ? new Date(msgHdr.date / 1000).toLocaleString() : "";
                      const fwdAuthor = msgHdr.mime2DecodedAuthor || msgHdr.author || "";
                      const fwdSubject = msgHdr.mime2DecodedSubject || msgHdr.subject || "";
                      const fwdRecipients = msgHdr.mime2DecodedRecipients || msgHdr.recipients || "";
                      const escapedBody = escapeHtml(originalBody).replace(/\n/g, '<br>');

                      const forwardBlock = `-------- Forwarded Message --------<br>` +
                        `Subject: ${escapeHtml(fwdSubject)}<br>` +
                        `Date: ${dateStr}<br>` +
                        `From: ${escapeHtml(fwdAuthor)}<br>` +
                        `To: ${escapeHtml(fwdRecipients)}<br><br>` +
                        escapedBody;

                      // Combine intro body + forward block
                      const introHtml = body ? formatBodyHtml(body, isHtml) + '<br><br>' : "";

                      composeFields.body = `<html><head><meta charset="UTF-8"></head><body>${introHtml}${forwardBlock}</body></html>`;

                      // Collect original message attachments as descriptors
                      const origDescs = [];
                      if (aMimeMsg && aMimeMsg.allUserAttachments) {
                        for (const att of aMimeMsg.allUserAttachments) {
                          try {
                            origDescs.push({ url: att.url, name: att.name, contentType: att.contentType });
                          } catch {
                            // Skip unreadable original attachments
                          }
                        }
                      }

                      // Validate user-specified file attachments
                      const { descs: fileDescs, failed: failedPaths } = filePathsToAttachDescs(attachments);

                      // Use New type - we build forward quote manually
                      msgComposeParams.type = Ci.nsIMsgCompType.New;
                      msgComposeParams.format = Ci.nsIMsgCompFormat.HTML;
                      msgComposeParams.composeFields = composeFields;

                      const identityResult = setComposeIdentity(msgComposeParams, from, folder.server);
                      if (identityResult && identityResult.error) { resolve(identityResult); return; }

                      const allDescs = [...origDescs, ...fileDescs];

                      if (skipReview) {
                        const msgURI = folder.getUriForMsg(msgHdr);
                        sendMessageDirectly(composeFields, msgComposeParams.identity, allDescs, msgURI, Ci.nsIMsgCompType.ForwardInline).then(result => {
                          if (result.success) {
                            let msg = `Forward sent with ${allDescs.length} attachment(s)`;
                            if (failedPaths.length > 0) msg += ` (failed to attach: ${failedPaths.join(", ")})`;
                            result.message = msg;
                          }
                          resolve(result);
                        });
                        return;
                      }

                      const msgComposeService = Cc["@mozilla.org/messengercompose;1"]
                        .getService(Ci.nsIMsgComposeService);
                      msgComposeService.OpenComposeWindowWithParams(null, msgComposeParams);

                      injectAttachmentsAsync(allDescs);

                      let msg = `Forward window opened with ${allDescs.length} attachment(s)`;
                      if (failedPaths.length > 0) msg += ` (failed to attach: ${failedPaths.join(", ")})`;
                      resolve({ success: true, message: msg });
                    } catch (e) {
                      resolve({ error: e.toString() });
                    }
                  }, true, { examineEncryptedParts: true });

                } catch (e) {
                  resolve({ error: e.toString() });
                }
              });
            }

            function displayMessage(messageId, folderPath, displayMode) {
              const found = findMessage(messageId, folderPath);
              if (found.error) return found;
              const { msgHdr } = found;

              const VALID_DISPLAY_MODES = ["3pane", "tab", "window"];
              const mode = displayMode || "3pane";
              if (!VALID_DISPLAY_MODES.includes(mode)) {
                return { error: `Invalid displayMode: "${mode}". Must be one of: ${VALID_DISPLAY_MODES.join(", ")}` };
              }

              try {
                const { MailUtils } = ChromeUtils.importESModule(
                  "resource:///modules/MailUtils.sys.mjs"
                );

                switch (mode) {
                  case "tab": {
                    const win = Services.wm.getMostRecentWindow("mail:3pane");
                    if (!win) return { error: "No Thunderbird mail window found (required for tab mode)" };
                    const msgUri = msgHdr.folder.getUriForMsg(msgHdr);
                    const tabmail = win.document.getElementById("tabmail");
                    if (!tabmail) return { error: "Could not access tabmail interface" };
                    tabmail.openTab("mailMessageTab", { messageURI: msgUri, msgHdr });
                    break;
                  }
                  case "window":
                    // openMessageInNewWindow doesn't need an existing 3pane window
                    MailUtils.openMessageInNewWindow(msgHdr);
                    break;
                  case "3pane":
                    MailUtils.displayMessageInFolderTab(msgHdr);
                    break;
                }
              } catch (e) {
                return { error: `Failed to display message: ${e.message || e}` };
              }

              return { success: true, displayMode: mode, subject: msgHdr.mime2DecodedSubject || msgHdr.subject || "" };
            }

            function getRecentMessages(folderPath, daysBack, maxResults, offset, unreadOnly, flaggedOnly, includeSubfolders) {
              const results = [];
              const days = Number.isFinite(Number(daysBack)) && Number(daysBack) > 0 ? Math.floor(Number(daysBack)) : 7;
              const cutoffTs = (Date.now() - days * 86400000) * 1000; // Thunderbird uses microseconds
              const requestedLimit = Number(maxResults);
              const effectiveLimit = Math.min(
                Number.isFinite(requestedLimit) && requestedLimit > 0 ? Math.floor(requestedLimit) : DEFAULT_MAX_RESULTS,
                MAX_SEARCH_RESULTS_CAP
              );

              function collectFromFolder(folder) {
                if (results.length >= SEARCH_COLLECTION_CAP) return;

                try {
                  const db = folder.msgDatabase;
                  if (!db) return;

                  for (const msgHdr of db.enumerateMessages()) {
                    if (results.length >= SEARCH_COLLECTION_CAP) break;

                    const msgDateTs = msgHdr.date || 0;
                    if (msgDateTs < cutoffTs) continue;
                    if (unreadOnly && msgHdr.isRead) continue;
                    if (flaggedOnly && !msgHdr.isFlagged) continue;

                    const msgTags = getUserTags(msgHdr);
                    const preview = msgHdr.getStringProperty("preview") || "";
                    const result = {
                      id: msgHdr.messageId,
                      threadId: msgHdr.threadId, // folder-local, use with folderPath for grouping
                      subject: msgHdr.mime2DecodedSubject || msgHdr.subject,
                      author: msgHdr.mime2DecodedAuthor || msgHdr.author,
                      recipients: msgHdr.mime2DecodedRecipients || msgHdr.recipients,
                      date: msgHdr.date ? new Date(msgHdr.date / 1000).toISOString() : null,
                      folder: folder.prettyName,
                      folderPath: folder.URI,
                      read: msgHdr.isRead,
                      flagged: msgHdr.isFlagged,
                      tags: msgTags,
                      _dateTs: msgDateTs
                    };
                    if (preview) result.preview = preview;
                    results.push(result);
                  }
                } catch {
                  // Skip inaccessible folders
                }

                const recurse = includeSubfolders !== false; // default true
                if (recurse && folder.hasSubFolders) {
                  for (const subfolder of folder.subFolders) {
                    if (results.length >= SEARCH_COLLECTION_CAP) break;
                    collectFromFolder(subfolder);
                  }
                }
              }

              if (folderPath) {
                // Specific folder
                const opened = openFolder(folderPath);
                if (opened.error) return { error: opened.error };
                collectFromFolder(opened.folder);
              } else {
                // All folders across accessible accounts
                for (const account of getAccessibleAccounts()) {
                  if (results.length >= SEARCH_COLLECTION_CAP) break;
                  try {
                    const root = account.incomingServer.rootFolder;
                    collectFromFolder(root);
                  } catch {
                    // Skip inaccessible accounts
                  }
                }
              }

              results.sort((a, b) => b._dateTs - a._dateTs);

              return paginate(results, offset, effectiveLimit);
            }

            function deleteMessages(messageIds, folderPath) {
              try {
                // MCP clients may send arrays as JSON strings
                if (typeof messageIds === "string") {
                  try { messageIds = JSON.parse(messageIds); } catch { /* leave as-is */ }
                }
                if (!Array.isArray(messageIds) || messageIds.length === 0) {
                  return { error: "messageIds must be a non-empty array of strings" };
                }
                if (typeof folderPath !== "string" || !folderPath) {
                  return { error: "folderPath must be a non-empty string" };
                }

                const opened = openFolder(folderPath);
                if (opened.error) return { error: opened.error };
                const { folder, db } = opened;

                // Find all requested message headers
                const found = [];
                const notFound = [];
                for (const msgId of messageIds) {
                  if (typeof msgId !== "string" || !msgId) {
                    notFound.push(msgId);
                    continue;
                  }
                  let hdr = null;
                  const hasDirectLookup = typeof db.getMsgHdrForMessageID === "function";
                  if (hasDirectLookup) {
                    try { hdr = db.getMsgHdrForMessageID(msgId); } catch { hdr = null; }
                  }
                  if (!hdr) {
                    for (const h of db.enumerateMessages()) {
                      if (h.messageId === msgId) { hdr = h; break; }
                    }
                  }
                  if (hdr) {
                    found.push(hdr);
                  } else {
                    notFound.push(msgId);
                  }
                }

                if (found.length === 0) {
                  return { error: "No matching messages found" };
                }

                // Drafts get moved to Trash instead of hard-deleted
                const DRAFTS_FLAG = 0x00000400;
                const isDrafts = typeof folder.getFlag === "function" && folder.getFlag(DRAFTS_FLAG);
                let trashFolder = null;

                if (isDrafts) {
                  trashFolder = findTrashFolder(folder);

                  if (trashFolder) {
                    MailServices.copy.copyMessages(folder, found, trashFolder, true, null, null, false);
                  } else {
                    // No trash found, fall back to regular delete
                    folder.deleteMessages(found, null, false, true, null, false);
                  }
                } else {
                  folder.deleteMessages(found, null, false, true, null, false);
                }

                let result = { success: true, deleted: found.length };
                if (isDrafts && trashFolder) result.movedToTrash = true;
                if (notFound.length > 0) result.notFound = notFound;
                return result;
              } catch (e) {
                return { error: e.toString() };
              }
            }

            function updateMessage(messageId, messageIds, folderPath, read, flagged, addTags, removeTags, moveTo, trash) {
              try {
                // Normalize to an array of IDs
                if (typeof messageIds === "string") {
                  try { messageIds = JSON.parse(messageIds); } catch { /* leave as-is */ }
                }
                if (messageId && messageIds) {
                  return { error: "Specify messageId or messageIds, not both" };
                }
                if (messageId) {
                  messageIds = [messageId];
                }
                if (!Array.isArray(messageIds) || messageIds.length === 0) {
                  return { error: "messageId or messageIds is required" };
                }
                if (typeof folderPath !== "string" || !folderPath) {
                  return { error: "folderPath must be a non-empty string" };
                }

                // Coerce boolean params (MCP clients may send strings)
                if (read !== undefined) read = read === true || read === "true";
                if (flagged !== undefined) flagged = flagged === true || flagged === "true";
                if (trash !== undefined) trash = trash === true || trash === "true";
                if (moveTo !== undefined && (typeof moveTo !== "string" || !moveTo)) {
                  return { error: "moveTo must be a non-empty string" };
                }
                // Coerce tag arrays (MCP clients may send JSON strings)
                if (typeof addTags === "string") {
                  try { addTags = JSON.parse(addTags); } catch { /* leave as-is */ }
                }
                if (typeof removeTags === "string") {
                  try { removeTags = JSON.parse(removeTags); } catch { /* leave as-is */ }
                }
                if (addTags !== undefined && !Array.isArray(addTags)) {
                  return { error: "addTags must be an array of tag keyword strings" };
                }
                if (removeTags !== undefined && !Array.isArray(removeTags)) {
                  return { error: "removeTags must be an array of tag keyword strings" };
                }

                if (moveTo && trash === true) {
                  return { error: "Cannot specify both moveTo and trash" };
                }

                // Find all requested message headers
                const opened = openFolder(folderPath);
                if (opened.error) return { error: opened.error };
                const { folder, db } = opened;

                const foundHdrs = [];
                const notFound = [];
                for (const msgId of messageIds) {
                  if (typeof msgId !== "string" || !msgId) {
                    notFound.push(msgId);
                    continue;
                  }
                  let hdr = null;
                  const hasDirectLookup = typeof db.getMsgHdrForMessageID === "function";
                  if (hasDirectLookup) {
                    try { hdr = db.getMsgHdrForMessageID(msgId); } catch { hdr = null; }
                  }
                  if (!hdr) {
                    for (const h of db.enumerateMessages()) {
                      if (h.messageId === msgId) { hdr = h; break; }
                    }
                  }
                  if (hdr) {
                    foundHdrs.push(hdr);
                  } else {
                    notFound.push(msgId);
                  }
                }

                if (foundHdrs.length === 0) {
                  return { error: "No matching messages found" };
                }

                const actions = [];

                if (read !== undefined) {
                  // Use folder-level API for proper IMAP sync (hdr.markRead
                  // only updates the local DB, doesn't queue IMAP commands)
                  folder.markMessagesRead(foundHdrs, read);
                  actions.push({ type: "read", value: read });
                }

                if (flagged !== undefined) {
                  folder.markMessagesFlagged(foundHdrs, flagged);
                  actions.push({ type: "flagged", value: flagged });
                }

                if (addTags || removeTags) {
                  // Validate: allow IMAP atom chars per RFC 3501 plus & for modified UTF-7
                  // tag keys that Thunderbird generates for non-ASCII labels.
                  // Blocks whitespace, null bytes, parens, braces, wildcards, quotes, backslash.
                  const VALID_TAG = /^[a-zA-Z0-9_$.\-&+!']+$/;
                  const tagsToAdd = (addTags || []).filter(t => typeof t === "string" && VALID_TAG.test(t));
                  const tagsToRemove = (removeTags || []).filter(t => typeof t === "string" && VALID_TAG.test(t));
                  // Use folder-level keyword APIs for proper IMAP sync
                  if (tagsToAdd.length > 0) {
                    folder.addKeywordsToMessages(foundHdrs, tagsToAdd.join(" "));
                    actions.push({ type: "addTags", value: tagsToAdd });
                  }
                  if (tagsToRemove.length > 0) {
                    folder.removeKeywordsFromMessages(foundHdrs, tagsToRemove.join(" "));
                    actions.push({ type: "removeTags", value: tagsToRemove });
                  }
                }

                let targetFolder = null;

                if (trash === true) {
                  targetFolder = findTrashFolder(folder);
                  if (!targetFolder) {
                    return { error: "Trash folder not found" };
                  }
                } else if (moveTo) {
                  const moveResult = getAccessibleFolder(moveTo);
                  if (moveResult.error) return moveResult;
                  targetFolder = moveResult.folder;
                }

                if (targetFolder) {
                  // Note: on IMAP, tags/flags set above may not transfer to the
                  // moved copy. If both tags and move are needed, consider making
                  // two separate updateMessage calls (tags first, then move).
                  MailServices.copy.copyMessages(folder, foundHdrs, targetFolder, true, null, null, false);
                  actions.push({ type: "move", to: targetFolder.URI });
                }

                const result = { success: true, updated: foundHdrs.length, actions };
                if (targetFolder && (addTags || removeTags)) {
                  result.warning = "Tags were applied before move; on IMAP accounts, tags may not transfer to the moved copy. Consider separate calls if tags are missing.";
                }
                if (notFound.length > 0) result.notFound = notFound;
                return result;
              } catch (e) {
                return { error: e.toString() };
              }
            }

            function createFolder(parentFolderPath, name) {
              try {
                if (typeof parentFolderPath !== "string" || !parentFolderPath) {
                  return { error: "parentFolderPath must be a non-empty string" };
                }
                if (typeof name !== "string" || !name) {
                  return { error: "name must be a non-empty string" };
                }

                const parentResult = getAccessibleFolder(parentFolderPath);
                if (parentResult.error) return parentResult;
                const parent = parentResult.folder;

                parent.createSubfolder(name, null);

                // Try to return the new folder's URI
                let newPath = null;
                try {
                  if (parent.hasSubFolders) {
                    for (const sub of parent.subFolders) {
                      if (sub.prettyName === name || sub.name === name) {
                        newPath = sub.URI;
                        break;
                      }
                    }
                  }
                } catch {
                  // Folder may not be immediately visible (IMAP)
                }

                return {
                  success: true,
                  message: `Folder "${name}" created`,
                  path: newPath
                };
              } catch (e) {
                const msg = e.toString();
                if (msg.includes("NS_MSG_FOLDER_EXISTS")) {
                  return { error: `Folder "${name}" already exists under this parent` };
                }
                return { error: msg };
              }
            }

            function renameFolder(folderPath, newName) {
              try {
                if (typeof folderPath !== "string" || !folderPath) {
                  return { error: "folderPath must be a non-empty string" };
                }
                if (typeof newName !== "string" || !newName) {
                  return { error: "newName must be a non-empty string" };
                }

                const renameResult = getAccessibleFolder(folderPath);
                if (renameResult.error) return renameResult;
                const folder = renameResult.folder;

                folder.rename(newName, null);
                return {
                  success: true,
                  message: `Folder renamed to "${newName}"`,
                  oldPath: folderPath,
                };
              } catch (e) {
                return { error: e.toString() };
              }
            }

            function deleteFolder(folderPath) {
              try {
                if (typeof folderPath !== "string" || !folderPath) {
                  return { error: "folderPath must be a non-empty string" };
                }

                const delResult = getAccessibleFolder(folderPath);
                if (delResult.error) return delResult;
                const folder = delResult.folder;
                const folderName = folder.prettyName || folder.name || folderPath;

                const parent = folder.parent;
                if (!parent) {
                  return { error: "Cannot delete a root folder" };
                }

                // Check if folder is already in Trash — if so, permanently delete
                const TRASH_FLAG = 0x00000100;
                let inTrash = false;
                let ancestor = folder;
                while (ancestor) {
                  try {
                    if (ancestor.getFlag && ancestor.getFlag(TRASH_FLAG)) {
                      inTrash = true;
                      break;
                    }
                  } catch { /* ignore */ }
                  ancestor = ancestor.parent;
                }

                if (inTrash) {
                  // Permanently delete — deleteSelf requires a msgWindow
                  const win = Services.wm.getMostRecentWindow("mail:3pane");
                  folder.deleteSelf(win?.msgWindow ?? null);
                  return { success: true, message: `Folder "${folderName}" permanently deleted` };
                } else {
                  // Move to trash
                  const trashFolder = findTrashFolder(folder);
                  if (!trashFolder) {
                    return { error: "Trash folder not found" };
                  }
                  MailServices.copy.copyFolder(folder, trashFolder, true, null, null);
                  return { success: true, message: `Folder "${folderName}" moved to Trash` };
                }
              } catch (e) {
                return { error: e.toString() };
              }
            }

            /**
             * Find a special folder by flag bit, searching the account's folder tree.
             */
            function findSpecialFolder(root, flagBit) {
              const search = (folder) => {
                try {
                  if (folder.getFlag && folder.getFlag(flagBit)) return folder;
                } catch {}
                if (folder.hasSubFolders) {
                  for (const sub of folder.subFolders) {
                    const found = search(sub);
                    if (found) return found;
                  }
                }
                return null;
              };
              return search(root);
            }

            /**
             * Recursively delete all messages in a folder and its subfolders.
             * Returns total count of messages deleted.
             */
            function deleteAllMessagesRecursive(folder) {
              let count = 0;
              try {
                const db = folder.msgDatabase;
                if (db) {
                  const hdrs = [];
                  for (const hdr of db.enumerateMessages()) hdrs.push(hdr);
                  if (hdrs.length > 0) {
                    folder.deleteMessages(hdrs, null, true, false, null, false);
                    count += hdrs.length;
                  }
                }
              } catch {}
              if (folder.hasSubFolders) {
                for (const sub of folder.subFolders) {
                  count += deleteAllMessagesRecursive(sub);
                }
              }
              return count;
            }

            function emptyTrash(accountId) {
              try {
                const TRASH_FLAG = 0x00000100;
                const accounts = accountId
                  ? [MailServices.accounts.getAccount(accountId)].filter(Boolean)
                  : Array.from(getAccessibleAccounts());
                if (accountId && accounts.length === 0) {
                  return { error: `Account not found: ${accountId}` };
                }
                if (accountId && !isAccountAllowed(accountId)) {
                  return { error: `Account not accessible: ${accountId}` };
                }

                const results = [];
                for (const account of accounts) {
                  const root = account.incomingServer?.rootFolder;
                  if (!root) continue;
                  const trash = findSpecialFolder(root, TRASH_FLAG);
                  if (!trash) {
                    results.push({ account: account.key, status: "no Trash folder found" });
                    continue;
                  }
                  // Use Thunderbird's native emptyTrash when available (handles
                  // IMAP expunge, subfolders, and compaction correctly)
                  if (typeof trash.emptyTrash === "function") {
                    const win = Services.wm.getMostRecentWindow("mail:3pane");
                    trash.emptyTrash(win?.msgWindow ?? null, null);
                    results.push({ account: account.key, folder: trash.URI, status: "emptied" });
                  } else {
                    // Fallback: manually delete messages in folder + subfolders
                    const deleted = deleteAllMessagesRecursive(trash);
                    results.push({ account: account.key, folder: trash.URI, deleted });
                  }
                }
                return { success: true, results };
              } catch (e) {
                return { error: e.toString() };
              }
            }

            function emptyJunk(accountId) {
              try {
                const JUNK_FLAG = 0x40000000;
                const accounts = accountId
                  ? [MailServices.accounts.getAccount(accountId)].filter(Boolean)
                  : Array.from(getAccessibleAccounts());
                if (accountId && accounts.length === 0) {
                  return { error: `Account not found: ${accountId}` };
                }
                if (accountId && !isAccountAllowed(accountId)) {
                  return { error: `Account not accessible: ${accountId}` };
                }

                const results = [];
                for (const account of accounts) {
                  const root = account.incomingServer?.rootFolder;
                  if (!root) continue;
                  const junk = findSpecialFolder(root, JUNK_FLAG);
                  if (!junk) {
                    results.push({ account: account.key, status: "no Junk folder found" });
                    continue;
                  }
                  // No native emptyJunk in Thunderbird, delete recursively
                  const deleted = deleteAllMessagesRecursive(junk);
                  results.push({ account: account.key, folder: junk.URI, deleted });
                }
                return { success: true, results };
              } catch (e) {
                return { error: e.toString() };
              }
            }

            function moveFolder(folderPath, newParentPath) {
              try {
                if (typeof folderPath !== "string" || !folderPath) {
                  return { error: "folderPath must be a non-empty string" };
                }
                if (typeof newParentPath !== "string" || !newParentPath) {
                  return { error: "newParentPath must be a non-empty string" };
                }

                const srcResult = getAccessibleFolder(folderPath);
                if (srcResult.error) return srcResult;
                const folder = srcResult.folder;
                const folderName = folder.prettyName || folder.name || folderPath;

                const destResult = getAccessibleFolder(newParentPath);
                if (destResult.error) return destResult;
                const newParent = destResult.folder;
                const parentName = newParent.prettyName || newParent.name || newParentPath;

                if (folder.parent && folder.parent.URI === newParentPath) {
                  return { error: "Folder is already under this parent" };
                }

                MailServices.copy.copyFolder(folder, newParent, true, null, null);
                return {
                  success: true,
                  message: `Folder "${folderName}" moved to "${parentName}"`,
                };
              } catch (e) {
                return { error: e.toString() };
              }
            }

            // ── Filter constant maps ──

            const ATTRIB_MAP = {
              subject: 0, from: 1, body: 2, date: 3, priority: 4,
              status: 5, to: 6, cc: 7, toOrCc: 8, allAddresses: 9,
              ageInDays: 10, size: 11, tag: 12, hasAttachment: 13,
              junkStatus: 14, junkPercent: 15, otherHeader: 16,
            };
            const ATTRIB_NAMES = Object.fromEntries(Object.entries(ATTRIB_MAP).map(([k, v]) => [v, k]));

            const OP_MAP = {
              contains: 0, doesntContain: 1, is: 2, isnt: 3, isEmpty: 4,
              isBefore: 5, isAfter: 6, isHigherThan: 7, isLowerThan: 8,
              beginsWith: 9, endsWith: 10, isInAB: 11, isntInAB: 12,
              isGreaterThan: 13, isLessThan: 14, matches: 15, doesntMatch: 16,
            };
            const OP_NAMES = Object.fromEntries(Object.entries(OP_MAP).map(([k, v]) => [v, k]));

            const ACTION_MAP = {
              moveToFolder: 0x01, copyToFolder: 0x02, changePriority: 0x03,
              delete: 0x04, markRead: 0x05, killThread: 0x06,
              watchThread: 0x07, markFlagged: 0x08, label: 0x09,
              reply: 0x0A, forward: 0x0B, stopExecution: 0x0C,
              deleteFromServer: 0x0D, leaveOnServer: 0x0E, junkScore: 0x0F,
              fetchBody: 0x10, addTag: 0x11, deleteBody: 0x12,
              markUnread: 0x14, custom: 0x15,
            };
            const ACTION_NAMES = Object.fromEntries(Object.entries(ACTION_MAP).map(([k, v]) => [v, k]));

            function getFilterListForAccount(accountId) {
              if (!isAccountAllowed(accountId)) {
                return { error: `Account not accessible: ${accountId}` };
              }
              const account = MailServices.accounts.getAccount(accountId);
              if (!account) return { error: `Account not found: ${accountId}` };
              const server = account.incomingServer;
              if (!server) return { error: "Account has no server" };
              if (server.canHaveFilters === false) return { error: "Account does not support filters" };
              const filterList = server.getFilterList(null);
              if (!filterList) return { error: "Could not access filter list" };
              return { account, server, filterList };
            }

            function serializeFilter(filter, index) {
              const terms = [];
              try {
                for (const term of filter.searchTerms) {
                  const t = {
                    attrib: ATTRIB_NAMES[term.attrib] || String(term.attrib),
                    op: OP_NAMES[term.op] || String(term.op),
                    booleanAnd: term.booleanAnd,
                  };
                  try {
                    if (term.attrib === 3 || term.attrib === 10) {
                      // Date or AgeInDays: try date first, then str
                      try {
                        const d = term.value.date;
                        t.value = d ? new Date(d / 1000).toISOString() : (term.value.str || "");
                      } catch { t.value = term.value.str || ""; }
                    } else {
                      t.value = term.value.str || "";
                    }
                  } catch { t.value = ""; }
                  if (term.arbitraryHeader) t.header = term.arbitraryHeader;
                  terms.push(t);
                }
              } catch {
                // searchTerms iteration may fail on some TB versions
                // Try indexed access via termAsString as fallback
              }

              const actions = [];
              for (let a = 0; a < filter.actionCount; a++) {
                try {
                  const action = filter.getActionAt(a);
                  const act = { type: ACTION_NAMES[action.type] || String(action.type) };
                  if (action.type === 0x01 || action.type === 0x02) {
                    act.value = action.targetFolderUri || "";
                  } else if (action.type === 0x03) {
                    act.value = String(action.priority);
                  } else if (action.type === 0x0F) {
                    act.value = String(action.junkScore);
                  } else {
                    try { if (action.strValue) act.value = action.strValue; } catch {}
                  }
                  actions.push(act);
                } catch {
                  // Skip unreadable actions
                }
              }

              return {
                index,
                name: filter.filterName,
                enabled: filter.enabled,
                type: filter.filterType,
                temporary: filter.temporary,
                terms,
                actions,
              };
            }

            function buildTerms(filter, conditions) {
              for (const cond of conditions) {
                const term = filter.createTerm();
                const attribNum = ATTRIB_MAP[cond.attrib] ?? parseInt(cond.attrib);
                if (isNaN(attribNum)) throw new Error(`Unknown attribute: ${cond.attrib}`);
                term.attrib = attribNum;

                const opNum = OP_MAP[cond.op] ?? parseInt(cond.op);
                if (isNaN(opNum)) throw new Error(`Unknown operator: ${cond.op}`);
                term.op = opNum;

                const value = term.value;
                value.attrib = term.attrib;
                value.str = cond.value || "";
                term.value = value;

                term.booleanAnd = cond.booleanAnd !== false;
                if (cond.header) term.arbitraryHeader = cond.header;
                filter.appendTerm(term);
              }
            }

            function buildActions(filter, actions) {
              for (const act of actions) {
                const action = filter.createAction();
                const typeNum = ACTION_MAP[act.type] ?? parseInt(act.type);
                if (isNaN(typeNum)) throw new Error(`Unknown action type: ${act.type}`);
                action.type = typeNum;

                if (act.value) {
                  if (typeNum === 0x01 || typeNum === 0x02) {
                    // Move/Copy to folder -- verify target is accessible
                    const targetCheck = getAccessibleFolder(act.value);
                    if (targetCheck.error) throw new Error(`Filter target folder not accessible: ${act.value}`);
                    action.targetFolderUri = act.value;
                  } else if (typeNum === 0x03) {
                    action.priority = parseInt(act.value);
                  } else if (typeNum === 0x0F) {
                    action.junkScore = parseInt(act.value);
                  } else {
                    action.strValue = act.value;
                  }
                }
                filter.appendAction(action);
              }
            }

            // ── Filter tool handlers ──

            function listFilters(accountId) {
              try {
                const results = [];
                let accounts;
                if (accountId) {
                  if (!isAccountAllowed(accountId)) {
                    return { error: `Account not accessible: ${accountId}` };
                  }
                  const account = MailServices.accounts.getAccount(accountId);
                  if (!account) return { error: `Account not found: ${accountId}` };
                  accounts = [account];
                } else {
                  accounts = Array.from(getAccessibleAccounts());
                }

                for (const account of accounts) {
                  if (!account) continue;
                  try {
                    const server = account.incomingServer;
                    if (!server || server.canHaveFilters === false) continue;

                    const filterList = server.getFilterList(null);
                    if (!filterList) continue;

                    const filters = [];
                    for (let i = 0; i < filterList.filterCount; i++) {
                      try {
                        filters.push(serializeFilter(filterList.getFilterAt(i), i));
                      } catch {
                        // Skip unreadable filters
                      }
                    }

                    results.push({
                      accountId: account.key,
                      accountName: server.prettyName,
                      filterCount: filterList.filterCount,
                      loggingEnabled: filterList.loggingEnabled,
                      filters,
                    });
                  } catch {
                    // Skip inaccessible accounts
                  }
                }

                return results;
              } catch (e) {
                return { error: e.toString() };
              }
            }

            function createFilter(accountId, name, enabled, type, conditions, actions, insertAtIndex) {
              try {
                // Coerce arrays from MCP client string serialization
                if (typeof conditions === "string") {
                  try { conditions = JSON.parse(conditions); } catch { /* leave as-is */ }
                }
                if (typeof actions === "string") {
                  try { actions = JSON.parse(actions); } catch { /* leave as-is */ }
                }
                if (typeof enabled === "string") enabled = enabled === "true";
                if (typeof type === "string") type = parseInt(type);
                if (typeof insertAtIndex === "string") insertAtIndex = parseInt(insertAtIndex);

                if (!Array.isArray(conditions) || conditions.length === 0) {
                  return { error: "conditions must be a non-empty array" };
                }
                if (!Array.isArray(actions) || actions.length === 0) {
                  return { error: "actions must be a non-empty array" };
                }

                const fl = getFilterListForAccount(accountId);
                if (fl.error) return fl;
                const { filterList } = fl;

                const filter = filterList.createFilter(name);
                filter.enabled = enabled !== false;
                filter.filterType = (Number.isFinite(type) && type > 0) ? type : 17; // inbox + manual

                buildTerms(filter, conditions);
                buildActions(filter, actions);

                const idx = (insertAtIndex != null && insertAtIndex >= 0)
                  ? Math.min(insertAtIndex, filterList.filterCount)
                  : filterList.filterCount;
                filterList.insertFilterAt(idx, filter);
                filterList.saveToDefaultFile();

                return {
                  success: true,
                  name: filter.filterName,
                  index: idx,
                  filterCount: filterList.filterCount,
                };
              } catch (e) {
                return { error: e.toString() };
              }
            }

            function updateFilter(accountId, filterIndex, name, enabled, type, conditions, actions) {
              try {
                // Coerce from MCP client
                if (typeof filterIndex === "string") filterIndex = parseInt(filterIndex);
                if (!Number.isInteger(filterIndex)) return { error: "filterIndex must be an integer" };
                if (typeof enabled === "string") enabled = enabled === "true";
                if (typeof type === "string") type = parseInt(type);
                if (typeof conditions === "string") {
                  try { conditions = JSON.parse(conditions); } catch {
                    return { error: "conditions must be a valid JSON array" };
                  }
                }
                if (typeof actions === "string") {
                  try { actions = JSON.parse(actions); } catch {
                    return { error: "actions must be a valid JSON array" };
                  }
                }

                const fl = getFilterListForAccount(accountId);
                if (fl.error) return fl;
                const { filterList } = fl;

                if (filterIndex < 0 || filterIndex >= filterList.filterCount) {
                  return { error: `Invalid filter index: ${filterIndex}` };
                }

                const filter = filterList.getFilterAt(filterIndex);
                const changes = [];

                if (name !== undefined) {
                  filter.filterName = name;
                  changes.push("name");
                }
                if (enabled !== undefined) {
                  filter.enabled = enabled;
                  changes.push("enabled");
                }
                if (type !== undefined) {
                  filter.filterType = type;
                  changes.push("type");
                }

                const replaceConditions = Array.isArray(conditions) && conditions.length > 0;
                const replaceActions = Array.isArray(actions) && actions.length > 0;

                if (replaceConditions || replaceActions) {
                  // No clearTerms/clearActions API -- rebuild filter via remove+insert
                  const newFilter = filterList.createFilter(filter.filterName);
                  newFilter.enabled = filter.enabled;
                  newFilter.filterType = filter.filterType;

                  // Build or copy conditions
                  if (replaceConditions) {
                    buildTerms(newFilter, conditions);
                    changes.push("conditions");
                  } else {
                    // Copy existing terms -- abort on failure to prevent data loss
                    let termsCopied = 0;
                    try {
                      for (const term of filter.searchTerms) {
                        const newTerm = newFilter.createTerm();
                        newTerm.attrib = term.attrib;
                        newTerm.op = term.op;
                        const val = newTerm.value;
                        val.attrib = term.attrib;
                        try { val.str = term.value.str || ""; } catch {}
                        try { if (term.attrib === 3) val.date = term.value.date; } catch {}
                        newTerm.value = val;
                        newTerm.booleanAnd = term.booleanAnd;
                        try { newTerm.beginsGrouping = term.beginsGrouping; } catch {}
                        try { newTerm.endsGrouping = term.endsGrouping; } catch {}
                        try { if (term.arbitraryHeader) newTerm.arbitraryHeader = term.arbitraryHeader; } catch {}
                        newFilter.appendTerm(newTerm);
                        termsCopied++;
                      }
                    } catch (e) {
                      return { error: `Failed to copy existing conditions: ${e.toString()}` };
                    }
                    if (termsCopied === 0) {
                      return { error: "Cannot update: failed to read existing filter conditions" };
                    }
                  }

                  // Build or copy actions
                  if (replaceActions) {
                    buildActions(newFilter, actions);
                    changes.push("actions");
                  } else {
                    for (let a = 0; a < filter.actionCount; a++) {
                      try {
                        const origAction = filter.getActionAt(a);
                        const newAction = newFilter.createAction();
                        newAction.type = origAction.type;
                        try { newAction.targetFolderUri = origAction.targetFolderUri; } catch {}
                        try { newAction.priority = origAction.priority; } catch {}
                        try { newAction.strValue = origAction.strValue; } catch {}
                        try { newAction.junkScore = origAction.junkScore; } catch {}
                        newFilter.appendAction(newAction);
                      } catch {}
                    }
                  }

                  filterList.removeFilterAt(filterIndex);
                  filterList.insertFilterAt(filterIndex, newFilter);
                }

                filterList.saveToDefaultFile();

                return {
                  success: true,
                  changes,
                  filter: serializeFilter(filterList.getFilterAt(filterIndex), filterIndex),
                };
              } catch (e) {
                return { error: e.toString() };
              }
            }

            function deleteFilter(accountId, filterIndex) {
              try {
                if (typeof filterIndex === "string") filterIndex = parseInt(filterIndex);
                if (!Number.isInteger(filterIndex)) return { error: "filterIndex must be an integer" };

                const fl = getFilterListForAccount(accountId);
                if (fl.error) return fl;
                const { filterList } = fl;

                if (filterIndex < 0 || filterIndex >= filterList.filterCount) {
                  return { error: `Invalid filter index: ${filterIndex}` };
                }

                const filter = filterList.getFilterAt(filterIndex);
                const filterName = filter.filterName;
                filterList.removeFilterAt(filterIndex);
                filterList.saveToDefaultFile();

                return { success: true, deleted: filterName, remainingCount: filterList.filterCount };
              } catch (e) {
                return { error: e.toString() };
              }
            }

            function reorderFilters(accountId, fromIndex, toIndex) {
              try {
                if (typeof fromIndex === "string") fromIndex = parseInt(fromIndex);
                if (typeof toIndex === "string") toIndex = parseInt(toIndex);
                if (!Number.isInteger(fromIndex)) return { error: "fromIndex must be an integer" };
                if (!Number.isInteger(toIndex)) return { error: "toIndex must be an integer" };

                const fl = getFilterListForAccount(accountId);
                if (fl.error) return fl;
                const { filterList } = fl;

                if (fromIndex < 0 || fromIndex >= filterList.filterCount) {
                  return { error: `Invalid source index: ${fromIndex}` };
                }
                if (toIndex < 0 || toIndex >= filterList.filterCount) {
                  return { error: `Invalid target index: ${toIndex}` };
                }

                // moveFilterAt is unreliable — use remove + insert instead
                // Adjust toIndex after removal: if moving down, indices shift
                const filter = filterList.getFilterAt(fromIndex);
                filterList.removeFilterAt(fromIndex);
                const adjustedTo = (fromIndex < toIndex) ? toIndex - 1 : toIndex;
                filterList.insertFilterAt(adjustedTo, filter);
                filterList.saveToDefaultFile();

                return { success: true, name: filter.filterName, fromIndex, toIndex };
              } catch (e) {
                return { error: e.toString() };
              }
            }

            function applyFilters(accountId, folderPath) {
              try {
                const fl = getFilterListForAccount(accountId);
                if (fl.error) return fl;
                const { filterList } = fl;

                const afResult = getAccessibleFolder(folderPath);
                if (afResult.error) return afResult;
                const folder = afResult.folder;

                // Try MailServices.filters first, fall back to XPCOM contract ID
                let filterService;
                try {
                  filterService = MailServices.filters;
                } catch {}
                if (!filterService) {
                  try {
                    filterService = Cc["@mozilla.org/messenger/filter-service;1"]
                      .getService(Ci.nsIMsgFilterService);
                  } catch {}
                }
                if (!filterService) {
                  return { error: "Filter service not available in this Thunderbird version" };
                }
                filterService.applyFiltersToFolders(filterList, [folder], null);

                // applyFiltersToFolders is async — returns immediately
                return {
                  success: true,
                  message: "Filters applied (processing may take a moment)",
                  folder: folderPath,
                  enabledFilters: (() => {
                    let count = 0;
                    for (let i = 0; i < filterList.filterCount; i++) {
                      if (filterList.getFilterAt(i).enabled) count++;
                    }
                    return count;
                  })(),
                };
              } catch (e) {
                return { error: e.toString() };
              }
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

                const expectedType = propSchema.type;
                if (expectedType === "array") {
                  if (!Array.isArray(value)) {
                    errors.push(`Parameter '${key}' must be an array, got ${typeof value}`);
                  }
                } else if (expectedType === "object") {
                  if (typeof value !== "object" || Array.isArray(value)) {
                    errors.push(`Parameter '${key}' must be an object, got ${Array.isArray(value) ? "array" : typeof value}`);
                  }
                } else if (expectedType === "integer") {
                  // JSON Schema "integer" is a whole number. typeof reports
                  // "number" for both integers and floats, so check explicitly.
                  if (typeof value !== "number" || !Number.isInteger(value)) {
                    errors.push(`Parameter '${key}' must be an integer, got ${typeof value === "number" ? "non-integer number" : typeof value}`);
                  }
                } else if (expectedType && typeof value !== expectedType) {
                  errors.push(`Parameter '${key}' must be ${expectedType}, got ${typeof value}`);
                }
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
                  } catch { /* leave as-is for validation to catch */ }
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
                  return await createEvent(args.title, args.startDate, args.endDate, args.location, args.description, args.calendarId, args.allDay, args.skipReview);
                case "listEvents":
                  return await listEvents(args.calendarId, args.startDate, args.endDate, args.maxResults);
                case "updateEvent":
                  return await updateEvent(args.eventId, args.calendarId, args.title, args.startDate, args.endDate, args.location, args.description);
                case "deleteEvent":
                  return await deleteEvent(args.eventId, args.calendarId);
                case "createTask":
                  return createTask(args.title, args.dueDate, args.calendarId);
                case "listTasks":
                  return await listTasks(args.calendarId, args.completed, args.dueBefore, args.maxResults);
                case "updateTask":
                  return await updateTask(args.taskId, args.calendarId, args.title, args.dueDate, args.description, args.completed, args.percentComplete, args.priority);
                case "sendMail":
                  return await composeMail(args.to, args.subject, args.body, args.cc, args.bcc, args.isHtml, args.from, args.attachments, args.skipReview);
                case "replyToMessage":
                  return await replyToMessage(args.messageId, args.folderPath, args.body, args.replyAll, args.isHtml, args.to, args.cc, args.bcc, args.from, args.attachments, args.skipReview);
                case "forwardMessage":
                  return await forwardMessage(args.messageId, args.folderPath, args.to, args.body, args.isHtml, args.cc, args.bcc, args.from, args.attachments, args.skipReview);
                case "getRecentMessages":
                  return getRecentMessages(args.folderPath, args.daysBack, args.maxResults, args.offset, args.unreadOnly, args.flaggedOnly, args.includeSubfolders);
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
                    case "initialize":
                      result = {
                        protocolVersion: "2024-11-05",
                        capabilities: { tools: {} },
                        serverInfo: { name: "thunderbird-mcp", version: "0.1.0" }
                      };
                      break;
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
              })();
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
              try { server.stop(() => {}); } catch {}
              globalThis.__tbMcpServer = null;
              throw writeErr;
            }
            console.log(`Thunderbird MCP server listening on port ${boundPort}`);
            console.log(`Connection info written to ${connFilePath}`);
            return { success: true, port: boundPort };
          } catch (e) {
            console.error("Failed to start MCP server:", e);
            // Stop server if it was started but something else failed
            if (globalThis.__tbMcpServer) {
              try { globalThis.__tbMcpServer.stop(() => {}); } catch {}
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
          } catch { /* build info not available */ }

          // Read connection info from temp file using XPCOM file I/O
          try {
            const tmpDir = Services.dirsvc.get("TmpD", Ci.nsIFile);
            tmpDir.append("thunderbird-mcp");
            const connFile = tmpDir.clone();
            connFile.append("connection.json");
            connectionFile = connFile.path;
            if (connFile.exists()) {
              const fis = Cc["@mozilla.org/network/file-input-stream;1"]
                .createInstance(Ci.nsIFileInputStream);
              fis.init(connFile, 0x01, 0, 0);
              const sis = Cc["@mozilla.org/scriptableinputstream;1"]
                .createInstance(Ci.nsIScriptableInputStream);
              sis.init(fis);
              const text = sis.read(sis.available());
              sis.close();
              const data = JSON.parse(text);
              port = data.port || null;
            }
          } catch { /* ignore */ }

          return {
            running: !!globalThis.__tbMcpStartPromise,
            port,
            connectionFile,
            buildVersion,
            buildDate,
          };
        },

        getAccountAccessConfig: async function() {
          const { MailServices } = ChromeUtils.importESModule(
            "resource:///modules/MailServices.sys.mjs"
          );
          let allowed = [];
          try {
            const pref = Services.prefs.getStringPref(PREF_ALLOWED_ACCOUNTS, "");
            if (pref) allowed = JSON.parse(pref);
          } catch { /* ignore */ }

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
