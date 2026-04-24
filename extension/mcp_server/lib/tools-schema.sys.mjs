/**
 * MCP tool schemas for the Thunderbird MCP server.
 *
 * Pure data -- no XPCOM, no closures, no behavior. Each tool entry has
 * `name` (the MCP tool name), `group` + `crud` (used by the settings UI
 * to organize tools), `title` (human-readable), `description` (sent to
 * MCP clients verbatim), and `inputSchema` (a JSON-Schema object).
 *
 * api.js validates that every entry has valid group + crud values
 * before exposing the list. New tools added here MUST also have a
 * matching case in callTool() in api.js.
 */

export const tools = [
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
    status: { type: "string", description: "New VEVENT STATUS: 'tentative', 'confirmed', or 'cancelled' (optional)" },
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
    description: "Run enabled filters on a folder and report per-filter results. With dry_run=true (default false) the filters are not executed; instead each filter is evaluated against every message in the folder and the report shows which filters WOULD hit how many messages, so you can verify before mutating anything. Always returns a detailed { filtersRun, messagesProcessed, byFilter: [{ name, hit, actions: [...] }] } summary.",
    inputSchema: {
      type: "object",
      properties: {
        accountId: { type: "string", description: "Account ID (uses its filters)" },
        folderPath: { type: "string", description: "Folder URI to apply filters to (from listFolders)" },
        dry_run: { type: "boolean", description: "If true, evaluate each filter against messages and report hit counts without executing any actions" }
      },
      required: ["accountId", "folderPath"]
    }
  },
  {
    name: "getAccountAccess",
    group: "system", crud: "read",
    title: "Get Account Access",
    description: "Get the current account access control list. Shows which accounts the MCP server can access. Account access is configured by the user in the extension settings page (Tools > Add-ons > Thunderbird MCP > Options) and cannot be changed via MCP tools.",
    inputSchema: { type: "object", properties: {}, required: [] },
  },
  {
    name: "inbox_inventory",
    group: "messages", crud: "read",
    title: "Inventory Messages",
    description: "Aggregate messages by a key (from_domain, from_address, subject_prefix, tag, day) and return grouped counts + sample IDs + sample subjects instead of raw headers. Same filter syntax as searchMessages (query, folderPath, date range, unreadOnly, flaggedOnly, tag, includeSubfolders). Use this for \"what's in the inbox?\" questions on large folders -- a 25k-message inbox becomes ~50 groups and fits in <2k tokens instead of 80k+. Set collectAllIds=true to get the full ID lists for use with bulk_move_by_query.",
    inputSchema: {
      type: "object",
      properties: {
        query: { type: "string", description: "Text to search. Same tokens and prefixes as searchMessages. Empty string = all messages." },
        folderPath: { type: "string", description: "Optional folder URI to scope the scan" },
        groupBy: { type: "string", enum: ["from_domain", "from_address", "subject_prefix", "tag", "day"], description: "Aggregation key (default: from_domain). subject_prefix strips Re:/Fwd: and caps at 30 chars. day buckets by ISO date (YYYY-MM-DD)." },
        startDate: { type: "string", description: "Only count messages on or after this ISO date" },
        endDate: { type: "string", description: "Only count messages on or before this ISO date. Date-only strings include the full day." },
        unreadOnly: { type: "boolean" },
        flaggedOnly: { type: "boolean" },
        tag: { type: "string", description: "Filter by tag keyword (e.g. $label1)" },
        includeSubfolders: { type: "boolean", description: "Default true" },
        maxGroups: { type: "integer", description: "Cap the number of groups returned (default 50, max 500). Groups sorted by count descending." },
        samplesPerGroup: { type: "integer", description: "Number of sample message IDs + subjects per group (default 3, max 20)" },
        collectAllIds: { type: "boolean", description: "If true, include the FULL message-id list in each group (to hand off to bulk_move_by_query). Omit by default to keep responses small. Capped at 5000 ids per group." }
      },
      required: []
    }
  },
  {
    name: "bulk_move_by_query",
    group: "messages", crud: "update",
    title: "Bulk Move by Query",
    description: "Search + move in one call, with dry-run safety. Finds all messages in folderPath matching the query and moves them to the target folder. dry_run defaults to TRUE -- returns { matched, sample } without executing so you can verify before acting. Set dry_run=false to actually move. Optional read/flagged/tag mutations applied during the move. Limit caps total messages per call (default 1000, max 10000). Halves both token cost and typo risk vs searchMessages + updateMessage.",
    inputSchema: {
      type: "object",
      properties: {
        query: { type: "string", description: "Same syntax as searchMessages (from:, subject:, to:, cc:, tokens)" },
        folderPath: { type: "string", description: "Source folder URI (from listFolders)" },
        to: { type: "string", description: "Destination folder URI. Required when dry_run is false. Omit to just count." },
        dry_run: { type: "boolean", description: "If true (default), do not execute -- return matched/sample only" },
        markRead: { type: "boolean", description: "Mark moved messages as read" },
        flagged: { type: "boolean", description: "Set flagged state on moved messages" },
        addTags: { type: "array", items: { type: "string" }, description: "Tag keywords to add" },
        removeTags: { type: "array", items: { type: "string" }, description: "Tag keywords to remove" },
        limit: { type: "integer", description: "Max messages to process in one call (default 1000, max 10000)" },
        startDate: { type: "string" },
        endDate: { type: "string" },
        unreadOnly: { type: "boolean" },
        flaggedOnly: { type: "boolean" },
        includeSubfolders: { type: "boolean", description: "Default false -- bulk ops should be folder-scoped unless explicitly widened" }
      },
      required: ["query", "folderPath"]
    }
  },
  {
    name: "createDrafts",
    group: "messages", crud: "create",
    title: "Create Drafts (batch)",
    description: "Create multiple draft messages in the identity's Drafts folder in a single call, without opening compose windows. Designed for parallel correspondence (e.g. 8 agency replies at once). Returns an array with one result per draft: { success: true, draftFolder, messageId } or { error }. Each draft spec mirrors sendMail's signature. No messages are sent -- everything lands in Drafts for review.",
    inputSchema: {
      type: "object",
      properties: {
        drafts: {
          type: "array",
          description: "Array of draft specs",
          items: {
            type: "object",
            properties: {
              to: { type: "string", description: "Recipient email address" },
              subject: { type: "string" },
              body: { type: "string" },
              cc: { type: "string" },
              bcc: { type: "string" },
              isHtml: { type: "boolean" },
              from: { type: "string", description: "Sender identity (email address or identity ID)" },
              attachments: {
                type: "array",
                description: "File paths (strings) or inline objects ({name, contentType, base64})",
                items: { oneOf: [ { type: "string" }, { type: "object" } ] }
              }
            },
            required: ["to", "subject", "body"]
          }
        }
      },
      required: ["drafts"]
    }
  }
];
