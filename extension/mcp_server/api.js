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

const MCP_PORT = 8765;
const DEFAULT_MAX_RESULTS = 50;
const MAX_SEARCH_RESULTS_CAP = 200;
const SEARCH_COLLECTION_CAP = 1000;

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
        title: "List Accounts",
        description: "List all email accounts and their identities",
        inputSchema: { type: "object", properties: {}, required: [] },
      },
      {
        name: "listFolders",
        title: "List Folders",
        description: "List all mail folders with URIs and message counts",
        inputSchema: {
          type: "object",
          properties: {
            accountId: { type: "string", description: "Optional account ID (from listAccounts) to limit results to a single account" },
            folderPath: { type: "string", description: "Optional folder URI to list only that folder and its subfolders" },
          },
          required: [],
        },
      },
      {
        name: "searchMessages",
        title: "Search Mail",
        description: "Search message headers and return IDs/folder paths you can use with getMessage to read full email content",
        inputSchema: {
          type: "object",
          properties: {
            query: { type: "string", description: "Text to search in subject, author, or recipients (use empty string to match all)" },
            folderPath: { type: "string", description: "Optional folder URI to limit search to that folder and its subfolders" },
            startDate: { type: "string", description: "Filter messages on or after this ISO 8601 date" },
            endDate: { type: "string", description: "Filter messages on or before this ISO 8601 date" },
            maxResults: { type: "number", description: "Maximum number of results to return (default 50, max 200)" },
            sortOrder: { type: "string", description: "Date sort order: asc (oldest first) or desc (newest first, default)" }
          },
          required: ["query"],
        },
      },
      {
        name: "getMessage",
        title: "Get Message",
        description: "Read the full content of an email message by its ID",
        inputSchema: {
          type: "object",
          properties: {
            messageId: { type: "string", description: "The message ID (from searchMessages results)" },
            folderPath: { type: "string", description: "The folder URI path (from searchMessages results)" },
            saveAttachments: { type: "boolean", description: "If true, save attachments to /tmp/thunderbird-mcp/<messageId>/ and include filePath in response (default: false)" }
          },
          required: ["messageId", "folderPath"],
        },
      },
      {
        name: "sendMail",
        title: "Compose Mail",
        description: "Open a compose window with pre-filled recipient, subject, and body for user review before sending",
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
            attachments: { type: "array", items: { type: "string" }, description: "Array of file paths to attach" },
          },
          required: ["to", "subject", "body"],
        },
      },
      {
        name: "listCalendars",
        title: "List Calendars",
        description: "Return the user's calendars",
        inputSchema: { type: "object", properties: {}, required: [] },
      },
      {
        name: "createEvent",
        title: "Create Event",
        description: "Open a pre-filled event dialog in Thunderbird for user review before saving",
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
          },
          required: ["title", "startDate"],
        },
      },
      {
        name: "searchContacts",
        title: "Search Contacts",
        description: "Find contacts the user interacted with",
        inputSchema: {
          type: "object",
          properties: {
            query: { type: "string", description: "Email address or name to search for" }
          },
          required: ["query"],
        },
      },
      {
        name: "replyToMessage",
        title: "Reply to Message",
        description: "Open a reply compose window for a specific message with proper threading",
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
            attachments: { type: "array", items: { type: "string" }, description: "Array of file paths to attach" },
          },
          required: ["messageId", "folderPath", "body"],
        },
      },
      {
        name: "forwardMessage",
        title: "Forward Message",
        description: "Open a forward compose window for a message with attachments preserved",
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
            attachments: { type: "array", items: { type: "string" }, description: "Array of additional file paths to attach" },
          },
          required: ["messageId", "folderPath", "to"],
        },
      },
      {
        name: "getRecentMessages",
        title: "Get Recent Messages",
        description: "Get recent messages from a specific folder or all folders, with date and unread filtering",
        inputSchema: {
          type: "object",
          properties: {
            folderPath: { type: "string", description: "Folder URI to list messages from (defaults to all Inboxes)" },
            daysBack: { type: "number", description: "Only return messages from the last N days (default: 7)" },
            maxResults: { type: "number", description: "Maximum number of results (default: 50, max: 200)" },
            unreadOnly: { type: "boolean", description: "Only return unread messages (default: false)" },
          },
          required: [],
        },
      },
      {
        name: "deleteMessages",
        title: "Delete Messages",
        description: "Delete messages from a folder. Drafts are moved to Trash instead of permanently deleted.",
        inputSchema: {
          type: "object",
          properties: {
            messageIds: { type: "array", items: { type: "string" }, description: "Array of message IDs to delete" },
            folderPath: { type: "string", description: "The folder URI containing the messages" },
          },
          required: ["messageIds", "folderPath"],
        },
      },
      {
        name: "updateMessage",
        title: "Update Message",
        description: "Update a message's read/flagged state and optionally move it to another folder or to Trash.",
        inputSchema: {
          type: "object",
          properties: {
            messageId: { type: "string", description: "The message ID (from searchMessages results)" },
            folderPath: { type: "string", description: "The folder URI path (from searchMessages results)" },
            read: { type: "boolean", description: "Set to true/false to mark read/unread (optional)" },
            flagged: { type: "boolean", description: "Set to true/false to flag/unflag (optional)" },
            moveTo: { type: "string", description: "Destination folder URI (optional). Cannot be used with trash." },
            trash: { type: "boolean", description: "Set to true to move message to Trash (optional). Cannot be used with moveTo." },
          },
          required: ["messageId", "folderPath"],
        },
      },
      {
        name: "createFolder",
        title: "Create Folder",
        description: "Create a new mail subfolder under an existing folder",
        inputSchema: {
          type: "object",
          properties: {
            parentFolderPath: { type: "string", description: "URI of the parent folder (from listFolders)" },
            name: { type: "string", description: "Name for the new subfolder" },
          },
          required: ["parentFolderPath", "name"],
        },
      },
    ];

    return {
      mcpServer: {
        start: async function() {
          // Guard against double-start on extension reload (port conflict)
          if (globalThis.__tbMcpStartPromise) {
            return await globalThis.__tbMcpStartPromise;
          }
          const startPromise = (async () => {
          try {
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
            try {
              const calModule = ChromeUtils.importESModule(
                "resource:///modules/calendar/calUtils.sys.mjs"
              );
              cal = calModule.cal;
              const { CalEvent: CE } = ChromeUtils.importESModule(
                "resource:///modules/CalEvent.sys.mjs"
              );
              CalEvent = CE;
            } catch {
              // Calendar not available
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
             * Thunderbird's httpd.sys.mjs writes response strings as raw bytes.
             * Pre-encode non-ASCII as UTF-8 byte chars and strip invalid controls.
             */
            function sanitizeForJson(text) {
              if (!text) return text;

              let sanitized = "";

              for (let i = 0; i < text.length; i++) {
                const code = text.charCodeAt(i);

                if (
                  (code >= 0x00 && code <= 0x08) ||
                  code === 0x0b ||
                  code === 0x0c ||
                  (code >= 0x0e && code <= 0x1f) ||
                  code === 0x7f
                ) {
                  continue;
                }

                if (code <= 0x7f) {
                  sanitized += text[i];
                  continue;
                }

                const codePoint = text.codePointAt(i);
                if (codePoint > 0xffff) {
                  sanitized += String.fromCharCode(
                    0xf0 | (codePoint >> 18),
                    0x80 | ((codePoint >> 12) & 0x3f),
                    0x80 | ((codePoint >> 6) & 0x3f),
                    0x80 | (codePoint & 0x3f)
                  );
                  i++;
                  continue;
                }

                if (codePoint <= 0x7ff) {
                  sanitized += String.fromCharCode(
                    0xc0 | (codePoint >> 6),
                    0x80 | (codePoint & 0x3f)
                  );
                  continue;
                }

                sanitized += String.fromCharCode(
                  0xe0 | (codePoint >> 12),
                  0x80 | ((codePoint >> 6) & 0x3f),
                  0x80 | (codePoint & 0x3f)
                );
              }

              return sanitized;
            }

            /**
             * Lists all email accounts and their identities.
             */
            function listAccounts() {
              const accounts = [];
              for (const account of MailServices.accounts.accounts) {
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
             * Lists all folders (optionally limited to a single account).
             * Depth is 0 for root children, increasing for subfolders.
             */
            function listFolders(accountId, folderPath) {
              const results = [];

              function walkFolder(folder, accountKey, depth) {
                try {
                  const prettyName = folder.prettyName;
                  results.push({
                    name: sanitizeForJson(prettyName) || folder.name || "(unnamed)",
                    path: folder.URI,
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
                const folder = MailServices.folderLookup.getFolderForURL(folderPath);
                if (!folder) {
                  return { error: `Folder not found: ${folderPath}` };
                }
                const accountKey = folder.server
                  ? (MailServices.accounts.findAccountForServer(folder.server)?.key || "unknown")
                  : "unknown";
                walkFolder(folder, accountKey, 0);
                return results;
              }

              if (accountId) {
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

              for (const account of MailServices.accounts.accounts) {
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
             * Finds an identity by email address or identity ID.
             * Returns null if not found.
             */
            function findIdentity(emailOrId) {
              if (!emailOrId) return null;
              const lowerInput = emailOrId.toLowerCase();
              for (const account of MailServices.accounts.accounts) {
                for (const identity of account.identities) {
                  if (identity.key === emailOrId || (identity.email || "").toLowerCase() === lowerInput) {
                    return identity;
                  }
                }
              }
              return null;
            }

            /**
             * Adds file attachments to compose fields.
             * Returns { added: number, failed: string[] } for failure reporting.
             */
            function addAttachments(composeFields, attachments) {
              const result = { added: 0, failed: [] };
              if (!attachments || !Array.isArray(attachments)) return result;
              for (const filePath of attachments) {
                try {
                  const file = Cc["@mozilla.org/file/local;1"].createInstance(Ci.nsIFile);
                  file.initWithPath(filePath);
                  if (file.exists()) {
                    const attachment = Cc["@mozilla.org/messengercompose/attachment;1"]
                      .createInstance(Ci.nsIMsgAttachment);
                    attachment.url = Services.io.newFileURI(file).spec;
                    attachment.name = file.leafName;
                    composeFields.addAttachment(attachment);
                    result.added++;
                  } else {
                    result.failed.push(filePath);
                  }
                } catch {
                  result.failed.push(filePath);
                }
              }
              return result;
            }

            function escapeHtml(s) {
              return s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
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
             * Returns warning string if `from` was specified but not found.
             */
	            function setComposeIdentity(msgComposeParams, from, fallbackServer) {
	              const identity = findIdentity(from);
	              if (identity) {
	                msgComposeParams.identity = identity;
	                return "";
	              }
              // Fallback to default identity for the account
              if (fallbackServer) {
                const account = MailServices.accounts.findAccountForServer(fallbackServer);
                if (account) msgComposeParams.identity = account.defaultIdentity;
              } else {
                const defaultAccount = MailServices.accounts.defaultAccount;
                if (defaultAccount) msgComposeParams.identity = defaultAccount.defaultIdentity;
              }
	              return from ? `unknown identity: ${from}, using default` : "";
	            }

	            /**
	             * Opens a folder and its message database.
	             * Best-effort refresh for IMAP folders (db may be stale).
	             * Returns { folder, db } or { error }.
	             */
	            function openFolder(folderPath) {
	              try {
	                const folder = MailServices.folderLookup.getFolderForURL(folderPath);
	                if (!folder) {
	                  return { error: `Folder not found: ${folderPath}` };
	                }

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

	            function searchMessages(query, folderPath, startDate, endDate, maxResults, sortOrder) {
	              const results = [];
	              const lowerQuery = (query || "").toLowerCase();
	              const hasQuery = !!lowerQuery;
	              const parsedStartDate = startDate ? new Date(startDate).getTime() : NaN;
              const parsedEndDate = endDate ? new Date(endDate).getTime() : NaN;
              const startDateTs = Number.isFinite(parsedStartDate) ? parsedStartDate * 1000 : null;
              // Add 24h only for date-only strings (no time component) to include the full day
              const endDateOffset = endDate && !endDate.includes("T") ? 86400000 : 0;
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

                    // IMPORTANT: Use mime2Decoded* properties for searching.
                    // Raw headers contain MIME encoding like "=?UTF-8?Q?...?="
                    // which won't match plain text searches.
                    const subject = (msgHdr.mime2DecodedSubject || msgHdr.subject || "").toLowerCase();
                    const author = (msgHdr.mime2DecodedAuthor || msgHdr.author || "").toLowerCase();
                    const recipients = (msgHdr.mime2DecodedRecipients || msgHdr.recipients || "").toLowerCase();
                    const ccList = (msgHdr.ccList || "").toLowerCase();
                    const msgDateTs = msgHdr.date || 0;

                    if (startDateTs !== null && msgDateTs < startDateTs) continue;
                    if (endDateTs !== null && msgDateTs > endDateTs) continue;

                    if (!hasQuery ||
                        subject.includes(lowerQuery) ||
                        author.includes(lowerQuery) ||
                        recipients.includes(lowerQuery) ||
                        ccList.includes(lowerQuery)) {
                      results.push({
                        id: msgHdr.messageId,
                        subject: sanitizeForJson(msgHdr.mime2DecodedSubject || msgHdr.subject),
                        author: sanitizeForJson(msgHdr.mime2DecodedAuthor || msgHdr.author),
                        recipients: sanitizeForJson(msgHdr.mime2DecodedRecipients || msgHdr.recipients),
                        ccList: sanitizeForJson(msgHdr.ccList),
                        date: msgHdr.date ? new Date(msgHdr.date / 1000).toISOString() : null,
                        folder: sanitizeForJson(folder.prettyName),
                        folderPath: folder.URI,
                        read: msgHdr.isRead,
                        flagged: msgHdr.isFlagged,
                        _dateTs: msgDateTs
                      });
                    }
                  }
                } catch {
                  // Skip inaccessible folders
                }

                if (folder.hasSubFolders) {
                  for (const subfolder of folder.subFolders) {
                    if (results.length >= SEARCH_COLLECTION_CAP) break;
                    searchFolder(subfolder);
                  }
                }
              }

              if (folderPath) {
                const folder = MailServices.folderLookup.getFolderForURL(folderPath);
                if (!folder) {
                  return { error: `Folder not found: ${folderPath}` };
                }
                searchFolder(folder);
              } else {
                for (const account of MailServices.accounts.accounts) {
                  if (results.length >= SEARCH_COLLECTION_CAP) break;
                  searchFolder(account.incomingServer.rootFolder);
                }
              }

              results.sort((a, b) => normalizedSortOrder === "asc" ? a._dateTs - b._dateTs : b._dateTs - a._dateTs);

              return results.slice(0, effectiveLimit).map(result => {
                delete result._dateTs;
                return result;
              });
            }

            function searchContacts(query) {
              const results = [];
              const lowerQuery = query.toLowerCase();

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
                      addressBook: book.dirName
                    });
                  }

                  if (results.length >= DEFAULT_MAX_RESULTS) break;
                }
                if (results.length >= DEFAULT_MAX_RESULTS) break;
              }

              return results;
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
                  readOnly: c.readOnly
                }));
              } catch (e) {
                return { error: e.toString() };
              }
            }

            function createEvent(title, startDate, endDate, location, description, calendarId, allDay) {
              if (!cal || !CalEvent) {
                return { error: "Calendar module not available" };
              }
              try {
                const win = Services.wm.getMostRecentWindow("mail:3pane");
                if (!win) {
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

	            function getMessage(messageId, folderPath, saveAttachments) {
	              return new Promise((resolve) => {
	                try {
	                  const found = findMessage(messageId, folderPath);
	                  if (found.error) {
	                    resolve({ error: found.error });
	                    return;
	                  }
	                  const { msgHdr } = found;

	                  const { MsgHdrToMimeMessage } = ChromeUtils.importESModule(
	                    "resource:///modules/gloda/MimeMessage.sys.mjs"
	                  );

                  MsgHdrToMimeMessage(msgHdr, null, (aMsgHdr, aMimeMsg) => {
                    if (!aMimeMsg) {
                      resolve({ error: "Could not parse message" });
                      return;
                    }

                    let body = "";
                    let bodyIsHtml = false;
                    try {
                      // sanitizeForJson removes control chars that break JSON
                      body = sanitizeForJson(aMimeMsg.coerceBodyToPlaintext());
                    } catch {
                      body = "";
                    }

                    // If plain text extraction failed, try to get HTML body from MIME parts
                    if (!body) {
                      try {
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
                            nbsp: " ",
                            amp: "&",
                            lt: "<",
                            gt: ">",
                            quot: "\"",
                            apos: "'",
                            "#39": "'",
                          };
                          text = text.replace(/&(#x?[0-9a-fA-F]+|[a-zA-Z]+);/gi, (match, entity) => {
                            if (entity.startsWith("#x") || entity.startsWith("#X")) {
                              const cp = parseInt(entity.slice(2), 16);
                              return cp ? String.fromCodePoint(cp) : match;
                            }
                            if (entity.startsWith("#")) {
                              const cp = parseInt(entity.slice(1), 10);
                              return cp ? String.fromCodePoint(cp) : match;
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

                        function findBody(part) {
                          const contentType = ((part.contentType || "").split(";")[0] || "").trim().toLowerCase();
                          if (contentType === "text/plain" && part.body) {
                            return { text: part.body, isHtml: false };
                          }
                          if (contentType === "text/html" && part.body) {
                            return { text: part.body, isHtml: true };
                          }
                          if (part.parts) {
                            let htmlFallback = null;
                            for (const sub of part.parts) {
                              const result = findBody(sub);
                              if (result && !result.isHtml) return result;
                              if (result && result.isHtml && !htmlFallback) htmlFallback = result;
                            }
                            if (htmlFallback) return htmlFallback;
                          }
                          return null;
                        }
                        const found = findBody(aMimeMsg);
                        if (found) {
                          let extracted = found.text;
                          if (found.isHtml) {
                            extracted = stripHtml(extracted);
                            bodyIsHtml = false;
                          } else {
                            bodyIsHtml = false;
                          }
                          body = sanitizeForJson(extracted);
                        } else {
                          body = "(Could not extract body text)";
                        }
                      } catch {
                        body = "(Could not extract body text)";
                      }
                    }

                    // Always collect attachment metadata
                    const attachments = [];
                    const attachmentSources = [];
                    if (aMimeMsg && aMimeMsg.allUserAttachments) {
                      for (const att of aMimeMsg.allUserAttachments) {
                        const info = {
                          name: sanitizeForJson(att?.name || ""),
                          contentType: sanitizeForJson(att?.contentType || ""),
                          size: typeof att?.size === "number" ? att.size : null
                        };
                        attachments.push(info);
                        attachmentSources.push({
                          info,
                          url: att?.url || "",
                          size: typeof att?.size === "number" ? att.size : null
                        });
                      }
                    }

                    const baseResponse = {
                      id: msgHdr.messageId,
                      subject: sanitizeForJson(msgHdr.mime2DecodedSubject || msgHdr.subject),
                      author: sanitizeForJson(msgHdr.mime2DecodedAuthor || msgHdr.author),
                      recipients: sanitizeForJson(msgHdr.mime2DecodedRecipients || msgHdr.recipients),
                      ccList: sanitizeForJson(msgHdr.ccList),
                      date: msgHdr.date ? new Date(msgHdr.date / 1000).toISOString() : null,
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
                      const root = Cc["@mozilla.org/file/local;1"].createInstance(Ci.nsIFile);
                      root.initWithPath("/tmp/thunderbird-mcp");
                      try {
                        root.create(Ci.nsIFile.DIRECTORY_TYPE, 0o755);
                      } catch (e) {
                        if (!root.exists() || !root.isDirectory()) throw e;
                        // already exists, fine
                      }
                      const dir = root.clone();
                      dir.append(sanitizedId);
                      try {
                        dir.create(Ci.nsIFile.DIRECTORY_TYPE, 0o755);
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
                            file.createUnique(Ci.nsIFile.NORMAL_FILE_TYPE, 0o644);
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
             * Opens a compose window with pre-filled fields.
             *
             * HTML body handling quirks:
             * 1. Strip newlines from HTML - Thunderbird adds <br> for each \n
             * 2. Encode non-ASCII as HTML entities - compose window has charset issues
             *    with emojis/unicode even with <meta charset="UTF-8">
             */
            function composeMail(to, subject, body, cc, bcc, isHtml, from, attachments) {
              try {
                const msgComposeService = Cc["@mozilla.org/messengercompose;1"]
                  .getService(Ci.nsIMsgComposeService);

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

                // Add file attachments
                const attResult = addAttachments(composeFields, attachments);

                msgComposeParams.type = Ci.nsIMsgCompType.New;
                msgComposeParams.format = Ci.nsIMsgCompFormat.HTML;
                msgComposeParams.composeFields = composeFields;

                const identityWarning = setComposeIdentity(msgComposeParams, from, null);

                msgComposeService.OpenComposeWindowWithParams(null, msgComposeParams);

                let msg = "Compose window opened";
                if (identityWarning) msg += ` (${identityWarning})`;
                if (attResult.failed.length > 0) {
                  msg += ` (failed to attach: ${attResult.failed.join(", ")})`;
                }
                return { success: true, message: msg };
              } catch (e) {
                return { error: e.toString() };
              }
            }

            /**
             * Opens a reply compose window for a message with quoted original.
             *
             * Uses nsIMsgCompType.New to preserve our body content, then manually
             * builds the quoted original message text. Threading is maintained
             * via the References and In-Reply-To headers.
             */
	            function replyToMessage(messageId, folderPath, body, replyAll, isHtml, to, cc, bcc, from, attachments) {
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
                      let originalBody = "";
                      if (aMimeMsg) {
                        try {
                          originalBody = aMimeMsg.coerceBodyToPlaintext() || "";
                        } catch {
                          originalBody = "";
                        }
                      }

                      const msgComposeService = Cc["@mozilla.org/messengercompose;1"]
                        .getService(Ci.nsIMsgComposeService);

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
                        // Get own email from the account identity for accurate self-filtering
                        const ownAccount = MailServices.accounts.findAccountForServer(folder.server);
                        const ownEmail = (ownAccount?.defaultIdentity?.email || "").toLowerCase();
                        const allRecipients = [
                          ...splitAddresses(msgHdr.recipients),
                          ...splitAddresses(msgHdr.ccList)
                        ]
                          .map(r => r.trim())
                          .filter(r => r && (!ownEmail || extractEmail(r) !== ownEmail));
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

                      const origSubject = msgHdr.subject || "";
                      composeFields.subject = origSubject.startsWith("Re:") ? origSubject : `Re: ${origSubject}`;

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

                      // Add file attachments
                      const attResult = addAttachments(composeFields, attachments);

                      msgComposeParams.type = Ci.nsIMsgCompType.New;
                      msgComposeParams.format = Ci.nsIMsgCompFormat.HTML;
                      msgComposeParams.composeFields = composeFields;

                      const identityWarning = setComposeIdentity(msgComposeParams, from, folder.server);

                      msgComposeService.OpenComposeWindowWithParams(null, msgComposeParams);

                      let msg = "Reply window opened";
                      if (identityWarning) msg += ` (${identityWarning})`;
                      if (attResult.failed.length > 0) {
                        msg += ` (failed to attach: ${attResult.failed.join(", ")})`;
                      }
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
             * Opens a forward compose window with attachments preserved.
             * Uses New type with manual forward quote to preserve both intro body and forwarded content.
             */
	            function forwardMessage(messageId, folderPath, to, body, isHtml, cc, bcc, from, attachments) {
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
                      const msgComposeService = Cc["@mozilla.org/messengercompose;1"]
                        .getService(Ci.nsIMsgComposeService);

                      const msgComposeParams = Cc["@mozilla.org/messengercompose/composeparams;1"]
                        .createInstance(Ci.nsIMsgComposeParams);

                      const composeFields = Cc["@mozilla.org/messengercompose/composefields;1"]
                        .createInstance(Ci.nsIMsgCompFields);

                      composeFields.to = to;
                      composeFields.cc = cc || "";
                      composeFields.bcc = bcc || "";

                      const origSubject = msgHdr.subject || "";
                      composeFields.subject = origSubject.startsWith("Fwd:") ? origSubject : `Fwd: ${origSubject}`;

                      // Get original body
                      let originalBody = "";
                      if (aMimeMsg) {
                        try {
                          originalBody = aMimeMsg.coerceBodyToPlaintext() || "";
                        } catch {
                          originalBody = "";
                        }
                      }

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

                      // Copy attachments from original message
                      let origAttCount = 0;
                      if (aMimeMsg && aMimeMsg.allUserAttachments) {
                        for (const att of aMimeMsg.allUserAttachments) {
                          try {
                            const attachment = Cc["@mozilla.org/messengercompose/attachment;1"]
                              .createInstance(Ci.nsIMsgAttachment);
                            attachment.url = att.url;
                            attachment.name = att.name;
                            attachment.contentType = att.contentType;
                            composeFields.addAttachment(attachment);
                            origAttCount++;
                          } catch {
                            // Skip unreadable original attachments
                          }
                        }
                      }

                      // Add user-specified file attachments
                      const attResult = addAttachments(composeFields, attachments);

                      // Use New type - we build forward quote manually
                      msgComposeParams.type = Ci.nsIMsgCompType.New;
                      msgComposeParams.format = Ci.nsIMsgCompFormat.HTML;
                      msgComposeParams.composeFields = composeFields;

                      const identityWarning = setComposeIdentity(msgComposeParams, from, folder.server);

                      msgComposeService.OpenComposeWindowWithParams(null, msgComposeParams);

                      let msg = `Forward window opened with ${origAttCount + attResult.added} attachment(s)`;
                      if (identityWarning) msg += ` (${identityWarning})`;
                      if (attResult.failed.length > 0) {
                        msg += ` (failed to attach: ${attResult.failed.join(", ")})`;
                      }
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

            function getRecentMessages(folderPath, daysBack, maxResults, unreadOnly) {
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

                    results.push({
                      id: msgHdr.messageId,
                      subject: sanitizeForJson(msgHdr.mime2DecodedSubject || msgHdr.subject),
                      author: sanitizeForJson(msgHdr.mime2DecodedAuthor || msgHdr.author),
                      recipients: sanitizeForJson(msgHdr.mime2DecodedRecipients || msgHdr.recipients),
                      date: msgHdr.date ? new Date(msgHdr.date / 1000).toISOString() : null,
                      folder: sanitizeForJson(folder.prettyName),
                      folderPath: folder.URI,
                      read: msgHdr.isRead,
                      flagged: msgHdr.isFlagged,
                      _dateTs: msgDateTs
                    });
                  }
                } catch {
                  // Skip inaccessible folders
                }

                if (folder.hasSubFolders) {
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
                // All folders across all accounts
                for (const account of MailServices.accounts.accounts) {
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

              return results.slice(0, effectiveLimit).map(r => {
                delete r._dateTs;
                return r;
              });
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
                  // Find trash folder
                  const TRASH_FLAG = 0x00000100;
                  try {
                    const account = MailServices.accounts.findAccountForServer(folder.server);
                    const root = account?.incomingServer?.rootFolder;
                    if (root) {
                      const stack = [root];
                      while (stack.length > 0 && !trashFolder) {
                        const current = stack.pop();
                        try {
                          if (current && typeof current.getFlag === "function" && current.getFlag(TRASH_FLAG)) {
                            trashFolder = current;
                          }
                        } catch {}
                        try {
                          if (current && current.hasSubFolders) {
                            for (const sf of current.subFolders) stack.push(sf);
                          }
                        } catch {}
                      }
                    }
                  } catch {}

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

            function updateMessage(messageId, folderPath, read, flagged, moveTo, trash) {
              try {
                if (typeof messageId !== "string" || !messageId) {
                  return { error: "messageId must be a non-empty string" };
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

                if (moveTo && trash === true) {
                  return { error: "Cannot specify both moveTo and trash" };
                }

                const found = findMessage(messageId, folderPath);
                if (found.error) return { error: found.error };

                const { msgHdr, folder } = found;
                const actions = [];

                if (read !== undefined) {
                  msgHdr.markRead(read);
                  actions.push({ type: "read", value: read });
                }

                if (flagged !== undefined) {
                  msgHdr.markFlagged(flagged);
                  actions.push({ type: "flagged", value: flagged });
                }

                let targetFolder = null;

                if (trash === true) {
                  const TRASH_FLAG = 0x00000100;
                  let account = null;
                  try {
                    account = MailServices.accounts.findAccountForServer(folder.server);
                  } catch {
                    account = null;
                  }
                  if (!account) {
                    return { error: "Could not determine account for message folder" };
                  }
                  const root = account.incomingServer?.rootFolder;
                  if (!root) {
                    return { error: "Account root folder not found" };
                  }

                  const stack = [root];
                  while (stack.length > 0 && !targetFolder) {
                    const current = stack.pop();
                    try {
                      if (current && typeof current.getFlag === "function" && current.getFlag(TRASH_FLAG)) {
                        targetFolder = current;
                        break;
                      }
                    } catch {
                      // ignore flag read errors
                    }
                    try {
                      if (current && current.hasSubFolders) {
                        for (const subfolder of current.subFolders) {
                          stack.push(subfolder);
                        }
                      }
                    } catch {
                      // ignore traversal errors
                    }
                  }

                  if (!targetFolder) {
                    return { error: "Trash folder not found" };
                  }
                } else if (moveTo) {
                  targetFolder = MailServices.folderLookup.getFolderForURL(moveTo);
                  if (!targetFolder) {
                    return { error: `Folder not found: ${moveTo}` };
                  }
                }

                if (targetFolder) {
                  MailServices.copy.copyMessages(folder, [msgHdr], targetFolder, true, null, null, false);
                  actions.push({ type: "move", to: targetFolder.URI });
                }

                return { success: true, actions };
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

                const parent = MailServices.folderLookup.getFolderForURL(parentFolderPath);
                if (!parent) {
                  return { error: `Parent folder not found: ${parentFolderPath}` };
                }

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

            async function callTool(name, args) {
              switch (name) {
                case "listAccounts":
                  return listAccounts();
                case "listFolders":
                  return listFolders(args.accountId, args.folderPath);
                case "searchMessages":
                  return searchMessages(args.query || "", args.folderPath, args.startDate, args.endDate, args.maxResults, args.sortOrder);
                case "getMessage":
                  return await getMessage(args.messageId, args.folderPath, args.saveAttachments);
                case "searchContacts":
                  return searchContacts(args.query || "");
                case "listCalendars":
                  return listCalendars();
                case "createEvent":
                  return createEvent(args.title, args.startDate, args.endDate, args.location, args.description, args.calendarId, args.allDay);
                case "sendMail":
                  return composeMail(args.to, args.subject, args.body, args.cc, args.bcc, args.isHtml, args.from, args.attachments);
                case "replyToMessage":
                  return await replyToMessage(args.messageId, args.folderPath, args.body, args.replyAll, args.isHtml, args.to, args.cc, args.bcc, args.from, args.attachments);
                case "forwardMessage":
                  return await forwardMessage(args.messageId, args.folderPath, args.to, args.body, args.isHtml, args.cc, args.bcc, args.from, args.attachments);
                case "getRecentMessages":
                  return getRecentMessages(args.folderPath, args.daysBack, args.maxResults, args.unreadOnly);
                case "deleteMessages":
                  return deleteMessages(args.messageIds, args.folderPath);
                case "updateMessage":
                  return updateMessage(args.messageId, args.folderPath, args.read, args.flagged, args.moveTo, args.trash);
                case "createFolder":
                  return createFolder(args.parentFolderPath, args.name);
                default:
                  throw new Error(`Unknown tool: ${name}`);
              }
            }

            const server = new HttpServer();

            server.registerPathHandler("/", (req, res) => {
              res.processAsync();

              if (req.method !== "POST") {
                res.setStatusLine("1.1", 405, "Method Not Allowed");
                res.write("POST only");
                res.finish();
                return;
              }

              let message;
              try {
                message = JSON.parse(readRequestBody(req));
              } catch {
                res.setStatusLine("1.1", 400, "Bad Request");
                res.write("Invalid JSON");
                res.finish();
                return;
              }

              const { id, method, params } = message;

              (async () => {
                try {
                  let result;
                  switch (method) {
                    case "tools/list":
                      result = { tools };
                      break;
                    case "tools/call":
                      if (!params?.name) {
                        throw new Error("Missing tool name");
                      }
                      result = {
                        content: [{
                          type: "text",
                          text: JSON.stringify(await callTool(params.name, params.arguments || {}), null, 2)
                        }]
                      };
                      break;
                    default:
                      res.setStatusLine("1.1", 404, "Not Found");
                      res.write(`Unknown method: ${method}`);
                      res.finish();
                      return;
                  }
                  res.setStatusLine("1.1", 200, "OK");
                  // charset=utf-8 is critical for proper emoji handling in responses
                  res.setHeader("Content-Type", "application/json; charset=utf-8", false);
                  res.write(JSON.stringify({ jsonrpc: "2.0", id, result }));
                } catch (e) {
                  res.setStatusLine("1.1", 200, "OK");
                  res.setHeader("Content-Type", "application/json; charset=utf-8", false);
                  res.write(JSON.stringify({
                    jsonrpc: "2.0",
                    id,
                    error: { code: -32000, message: e.toString() }
                  }));
                }
                res.finish();
              })();
            });

            server.start(MCP_PORT);
            console.log(`Thunderbird MCP server listening on port ${MCP_PORT}`);
            return { success: true, port: MCP_PORT };
          } catch (e) {
            console.error("Failed to start MCP server:", e);
            // Clear cached promise so a retry can attempt to bind again
            globalThis.__tbMcpStartPromise = null;
            return { success: false, error: e.toString() };
          }
          })();
          globalThis.__tbMcpStartPromise = startPromise;
          return await startPromise;
        }
      }
    };
  }

  onShutdown(isAppShutdown) {
    if (isAppShutdown) return;
    resProto.setSubstitution("thunderbird-mcp", null);
    Services.obs.notifyObservers(null, "startupcache-invalidate");
  }
};
