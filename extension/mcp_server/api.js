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
const MAX_SEARCH_RESULTS = 50;

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
        name: "searchMessages",
        title: "Search Mail",
        description: "Find messages using Thunderbird's search index",
        inputSchema: {
          type: "object",
          properties: {
            query: { type: "string", description: "Text to search for in messages (searches subject, body, author)" }
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
            folderPath: { type: "string", description: "The folder URI path (from searchMessages results)" }
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
            cc: { type: "string", description: "CC recipient (optional)" },
            isHtml: { type: "boolean", description: "Set to true if body contains HTML markup (default: false)" },
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
          },
          required: ["messageId", "folderPath", "body"],
        },
      },
      {
        name: "listAccounts",
        title: "List Email Accounts",
        description: "List all configured email accounts in Thunderbird",
        inputSchema: { type: "object", properties: {}, required: [] },
      },
      {
        name: "listFolders",
        title: "List Folders",
        description: "List all folders for a specific account or all accounts",
        inputSchema: {
          type: "object",
          properties: {
            accountKey: { type: "string", description: "Optional account key to list folders for (from listAccounts). If not provided, lists folders for all accounts" }
          },
          required: [],
        },
      },
      {
        name: "getRecentMessages",
        title: "Get Recent Messages",
        description: "Get recent messages from specific folders, sorted by date",
        inputSchema: {
          type: "object",
          properties: {
            folderPath: { type: "string", description: "Optional folder URI to get messages from. If not provided, searches Inbox folders" },
            limit: { type: "number", description: "Maximum number of messages to return (default: 20, max: 100)" },
            daysBack: { type: "number", description: "Number of days back to search (default: 30)" },
            unreadOnly: { type: "boolean", description: "Only return unread messages (default: false)" }
          },
          required: [],
        },
      },
    ];

    return {
      mcpServer: {
        start: async function() {
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
            try {
              const calModule = ChromeUtils.importESModule(
                "resource:///modules/calendar/calUtils.sys.mjs"
              );
              cal = calModule.cal;
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
             * Email bodies may contain control characters (BEL, etc.) that break
             * JSON.stringify. Remove them but preserve \n, \r, \t.
             */
            function sanitizeForJson(text) {
              if (!text) return text;
              return text.replace(/[\x00-\x08\x0b\x0c\x0e-\x1f\x7f]/g, '');
            }

            function searchMessages(query) {
              const results = [];
              const lowerQuery = query.toLowerCase();

              function searchFolder(folder) {
                if (results.length >= MAX_SEARCH_RESULTS) return;

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
                    if (results.length >= MAX_SEARCH_RESULTS) break;

                    // IMPORTANT: Use mime2Decoded* properties for searching.
                    // Raw headers contain MIME encoding like "=?UTF-8?Q?...?="
                    // which won't match plain text searches.
                    const subject = (msgHdr.mime2DecodedSubject || msgHdr.subject || "").toLowerCase();
                    const author = (msgHdr.mime2DecodedAuthor || msgHdr.author || "").toLowerCase();
                    const recipients = (msgHdr.mime2DecodedRecipients || msgHdr.recipients || "").toLowerCase();

                    if (subject.includes(lowerQuery) ||
                        author.includes(lowerQuery) ||
                        recipients.includes(lowerQuery)) {
                      results.push({
                        id: msgHdr.messageId,
                        subject: msgHdr.mime2DecodedSubject || msgHdr.subject,
                        author: msgHdr.mime2DecodedAuthor || msgHdr.author,
                        recipients: msgHdr.mime2DecodedRecipients || msgHdr.recipients,
                        date: msgHdr.date ? new Date(msgHdr.date / 1000).toISOString() : null,
                        folder: folder.prettyName,
                        folderPath: folder.URI,
                        read: msgHdr.isRead,
                        flagged: msgHdr.isFlagged
                      });
                    }
                  }
                } catch {
                  // Skip inaccessible folders
                }

                if (folder.hasSubFolders) {
                  for (const subfolder of folder.subFolders) {
                    if (results.length >= MAX_SEARCH_RESULTS) break;
                    searchFolder(subfolder);
                  }
                }
              }

              for (const account of MailServices.accounts.accounts) {
                if (results.length >= MAX_SEARCH_RESULTS) break;
                searchFolder(account.incomingServer.rootFolder);
              }

              return results;
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

                  if (results.length >= MAX_SEARCH_RESULTS) break;
                }
                if (results.length >= MAX_SEARCH_RESULTS) break;
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

            /**
             * List all configured email accounts in Thunderbird.
             * 
             * Returns account information including:
             * - key: Unique account identifier
             * - name: Account display name
             * - type: Account type (imap, pop3, rss, none)
             * - email: Primary email address
             * - username: Server username
             * - hostName: Server hostname
             * 
             * @returns {Array<Object>} Array of account objects
             */
            function listAccounts() {
              try {
                const accounts = [];
                for (const account of MailServices.accounts.accounts) {
                  const server = account.incomingServer;
                  accounts.push({
                    key: account.key,
                    name: account.name || server.prettyName,
                    type: server.type,
                    email: account.defaultIdentity?.email || "",
                    username: server.username,
                    hostName: server.hostName
                  });
                }
                return accounts;
              } catch (e) {
                return { error: e.toString() };
              }
            }

            /**
             * List all folders (mailboxes) for one or all accounts.
             * 
             * Recursively traverses folder hierarchy and returns detailed information:
             * - name: Folder display name
             * - path: Folder URI (used for other operations)
             * - type: inbox, sent, drafts, trash, templates, or folder
             * - accountKey: Parent account identifier
             * - accountName: Parent account name
             * - totalMessages: Total message count
             * - unreadMessages: Unread message count
             * 
             * Folder type detection uses Thunderbird's folder flags:
             * - 0x00001000: Inbox
             * - 0x00000200: Sent
             * - 0x00000400: Drafts  
             * - 0x00000100: Trash
             * - 0x00000800: Templates
             * 
             * @param {string} accountKey - Optional account key to filter by (from listAccounts)
             * @returns {Array<Object>} Array of folder objects
             */
            function listFolders(accountKey) {
              try {
                const folders = [];

                function addFolderAndSubfolders(folder, accountInfo) {
                  // Skip virtual folders (search folders, unified folders)
                  if (folder.flags & 0x1000) return; // Ci.nsMsgFolderFlags.Virtual

                  folders.push({
                    name: folder.prettyName,
                    path: folder.URI,
                    type: folder.flags & 0x00001000 ? "inbox" : 
                          folder.flags & 0x00000200 ? "sent" :
                          folder.flags & 0x00000400 ? "drafts" :
                          folder.flags & 0x00000100 ? "trash" :
                          folder.flags & 0x00000800 ? "templates" : "folder",
                    accountKey: accountInfo.key,
                    accountName: accountInfo.name,
                    totalMessages: folder.getTotalMessages(false),
                    unreadMessages: folder.getNumUnread(false)
                  });

                  // Recursively process subfolders
                  if (folder.hasSubFolders) {
                    for (const subfolder of folder.subFolders) {
                      addFolderAndSubfolders(subfolder, accountInfo);
                    }
                  }
                }

                if (accountKey) {
                  // List folders for specific account
                  const account = MailServices.accounts.getAccount(accountKey);
                  if (!account) {
                    return { error: `Account not found: ${accountKey}` };
                  }
                  const accountInfo = { key: account.key, name: account.name };
                  addFolderAndSubfolders(account.incomingServer.rootFolder, accountInfo);
                } else {
                  // List folders for all accounts
                  for (const account of MailServices.accounts.accounts) {
                    const accountInfo = { key: account.key, name: account.name };
                    addFolderAndSubfolders(account.incomingServer.rootFolder, accountInfo);
                  }
                }

                return folders;
              } catch (e) {
                return { error: e.toString() };
              }
            }

            /**
             * Get recent messages from folders, sorted by date.
             * 
             * This function provides fast, efficient access to recent messages with:
             * - Date-based filtering (default: last 30 days)
             * - Sorting by date (newest first)
             * - Optional unread-only filter
             * - Folder-specific or inbox-wide search
             * - Configurable result limit (max: 100)
             * 
             * When no folderPath is specified, searches all Inbox folders across accounts.
             * For IMAP folders, attempts to refresh the folder before reading.
             * 
             * Uses mime2Decoded* properties for proper character encoding.
             * 
             * @param {string} folderPath - Optional folder URI to search (from listFolders)
             * @param {number} limit - Maximum messages to return (default: 20, max: 100)
             * @param {number} daysBack - Days to look back (default: 30)
             * @param {boolean} unreadOnly - Only return unread messages (default: false)
             * @returns {Array<Object>} Array of message objects, sorted newest first
             */
            function getRecentMessages(folderPath, limit, daysBack, unreadOnly) {
              try {
                const maxLimit = Math.min(limit || 20, 100);
                const days = daysBack || 30;
                const cutoffDate = Date.now() - (days * 24 * 60 * 60 * 1000);
                const messages = [];

                function getMessagesFromFolder(folder) {
                  try {
                    // Attempt to refresh IMAP folders to get latest messages
                    // This is async and may not complete before we read
                    if (folder.server && folder.server.type === "imap") {
                      try {
                        folder.updateFolder(null);
                      } catch {}
                    }

                    const db = folder.msgDatabase;
                    if (!db) return;

                    for (const msgHdr of db.enumerateMessages()) {
                      // Thunderbird stores dates in microseconds
                      const msgDate = msgHdr.date ? msgHdr.date / 1000 : 0;
                      
                      // Skip messages older than cutoff date
                      if (msgDate < cutoffDate) continue;
                      
                      // Skip read messages if unreadOnly is true
                      if (unreadOnly && msgHdr.isRead) continue;

                      messages.push({
                        id: msgHdr.messageId,
                        subject: msgHdr.mime2DecodedSubject || msgHdr.subject,
                        author: msgHdr.mime2DecodedAuthor || msgHdr.author,
                        recipients: msgHdr.mime2DecodedRecipients || msgHdr.recipients,
                        date: new Date(msgDate).toISOString(),
                        folder: folder.prettyName,
                        folderPath: folder.URI,
                        read: msgHdr.isRead,
                        flagged: msgHdr.isFlagged,
                        dateTimestamp: msgDate  // Temporary field for sorting
                      });
                    }
                  } catch {}
                }

                if (folderPath) {
                  // Get messages from specific folder
                  const folder = MailServices.folderLookup.getFolderForURL(folderPath);
                  if (!folder) {
                    return { error: `Folder not found: ${folderPath}` };
                  }
                  getMessagesFromFolder(folder);
                } else {
                  // Get messages from all Inbox folders across all accounts
                  for (const account of MailServices.accounts.accounts) {
                    const rootFolder = account.incomingServer.rootFolder;
                    
                    // Recursively find inbox folders
                    function findInbox(folder) {
                      if (folder.flags & 0x00001000) { // Inbox flag
                        getMessagesFromFolder(folder);
                        return;
                      }
                      if (folder.hasSubFolders) {
                        for (const subfolder of folder.subFolders) {
                          findInbox(subfolder);
                        }
                      }
                    }
                    
                    findInbox(rootFolder);
                  }
                }

                // Sort by date descending (newest first)
                messages.sort((a, b) => b.dateTimestamp - a.dateTimestamp);
                const limitedMessages = messages.slice(0, maxLimit);
                
                // Remove temporary timestamp field
                limitedMessages.forEach(msg => delete msg.dateTimestamp);
                
                return limitedMessages;
              } catch (e) {
                return { error: e.toString() };
              }
            }

            function getMessage(messageId, folderPath) {
              return new Promise((resolve) => {
                try {
                  const folder = MailServices.folderLookup.getFolderForURL(folderPath);
                  if (!folder) {
                    resolve({ error: `Folder not found: ${folderPath}` });
                    return;
                  }

                  const db = folder.msgDatabase;
                  if (!db) {
                    resolve({ error: "Could not access folder database" });
                    return;
                  }

                  let msgHdr = null;
                  for (const hdr of db.enumerateMessages()) {
                    if (hdr.messageId === messageId) {
                      msgHdr = hdr;
                      break;
                    }
                  }

                  if (!msgHdr) {
                    resolve({ error: `Message not found: ${messageId}` });
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

                    let body = "";
                    try {
                      // sanitizeForJson removes control chars that break JSON
                      body = sanitizeForJson(aMimeMsg.coerceBodyToPlaintext());
                    } catch {
                      body = "(Could not extract body text)";
                    }

                    resolve({
                      id: msgHdr.messageId,
                      subject: msgHdr.subject,
                      author: msgHdr.author,
                      recipients: msgHdr.recipients,
                      ccList: msgHdr.ccList,
                      date: msgHdr.date ? new Date(msgHdr.date / 1000).toISOString() : null,
                      body
                    });
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
            function composeMail(to, subject, body, cc, isHtml) {
              try {
                const msgComposeService = Cc["@mozilla.org/messengercompose;1"]
                  .getService(Ci.nsIMsgComposeService);

                const msgComposeParams = Cc["@mozilla.org/messengercompose/composeparams;1"]
                  .createInstance(Ci.nsIMsgComposeParams);

                const composeFields = Cc["@mozilla.org/messengercompose/composefields;1"]
                  .createInstance(Ci.nsIMsgCompFields);

                composeFields.to = to || "";
                composeFields.cc = cc || "";
                composeFields.subject = subject || "";

                if (isHtml) {
                  let bodyText = (body || "").replace(/\n/g, '');
                  // Convert non-ASCII to HTML entities (handles emojis > U+FFFF)
                  bodyText = [...bodyText].map(c => c.codePointAt(0) > 127 ? `&#${c.codePointAt(0)};` : c).join('');
                  composeFields.body = bodyText.includes('<html')
                    ? bodyText
                    : `<html><head><meta charset="UTF-8"></head><body>${bodyText}</body></html>`;
                } else {
                  const htmlBody = (body || "")
                    .replace(/&/g, '&amp;')
                    .replace(/</g, '&lt;')
                    .replace(/>/g, '&gt;')
                    .replace(/\n/g, '<br>');
                  composeFields.body = `<html><body>${htmlBody}</body></html>`;
                }

                msgComposeParams.type = Ci.nsIMsgCompType.New;
                msgComposeParams.format = Ci.nsIMsgCompFormat.HTML;
                msgComposeParams.composeFields = composeFields;

                const defaultAccount = MailServices.accounts.defaultAccount;
                if (defaultAccount) {
                  msgComposeParams.identity = defaultAccount.defaultIdentity;
                }

                msgComposeService.OpenComposeWindowWithParams(null, msgComposeParams);

                return { success: true, message: "Compose window opened" };
              } catch (e) {
                return { error: e.toString() };
              }
            }

            /**
             * Opens a reply compose window for a message.
             *
             * IMPORTANT: Uses nsIMsgCompType.New instead of Reply/ReplyAll.
             * Using Reply type causes Thunderbird to overwrite our body with
             * the quoted original message. We manually set To, Subject, and
             * References headers for proper threading.
             *
             * Limitation: Does not include quoted original message text.
             */
            function replyToMessage(messageId, folderPath, body, replyAll, isHtml) {
              try {
                const folder = MailServices.folderLookup.getFolderForURL(folderPath);
                if (!folder) {
                  return { error: `Folder not found: ${folderPath}` };
                }

                const db = folder.msgDatabase;
                if (!db) {
                  return { error: "Could not access folder database" };
                }

                let msgHdr = null;
                for (const hdr of db.enumerateMessages()) {
                  if (hdr.messageId === messageId) {
                    msgHdr = hdr;
                    break;
                  }
                }

                if (!msgHdr) {
                  return { error: `Message not found: ${messageId}` };
                }

                const msgComposeService = Cc["@mozilla.org/messengercompose;1"]
                  .getService(Ci.nsIMsgComposeService);

                const msgComposeParams = Cc["@mozilla.org/messengercompose/composeparams;1"]
                  .createInstance(Ci.nsIMsgComposeParams);

                const composeFields = Cc["@mozilla.org/messengercompose/composefields;1"]
                  .createInstance(Ci.nsIMsgCompFields);

                if (replyAll) {
                  composeFields.to = msgHdr.author;
                  const otherRecipients = (msgHdr.recipients || "").split(",")
                    .map(r => r.trim())
                    .filter(r => r && !r.includes(folder.server.username));
                  if (otherRecipients.length > 0) {
                    composeFields.cc = otherRecipients.join(", ");
                  }
                } else {
                  composeFields.to = msgHdr.author;
                }

                const origSubject = msgHdr.subject || "";
                composeFields.subject = origSubject.startsWith("Re:") ? origSubject : `Re: ${origSubject}`;

                // References header enables proper email threading
                composeFields.references = `<${messageId}>`;

                // Same HTML/charset handling as composeMail
                if (isHtml) {
                  let bodyText = (body || "").replace(/\n/g, '');
                  bodyText = [...bodyText].map(c => c.codePointAt(0) > 127 ? `&#${c.codePointAt(0)};` : c).join('');
                  composeFields.body = `<html><head><meta charset="UTF-8"></head><body>${bodyText}</body></html>`;
                } else {
                  const htmlBody = (body || "")
                    .replace(/&/g, '&amp;')
                    .replace(/</g, '&lt;')
                    .replace(/>/g, '&gt;')
                    .replace(/\n/g, '<br>');
                  composeFields.body = `<html><body>${htmlBody}</body></html>`;
                }

                // New type preserves our body; Reply type overwrites it
                msgComposeParams.type = Ci.nsIMsgCompType.New;
                msgComposeParams.format = Ci.nsIMsgCompFormat.HTML;
                msgComposeParams.composeFields = composeFields;

                const account = MailServices.accounts.findAccountForServer(folder.server);
                if (account) {
                  msgComposeParams.identity = account.defaultIdentity;
                }

                msgComposeService.OpenComposeWindowWithParams(null, msgComposeParams);

                return { success: true, message: "Reply window opened" };
              } catch (e) {
                return { error: e.toString() };
              }
            }

            async function callTool(name, args) {
              switch (name) {
                case "searchMessages":
                  return searchMessages(args.query || "");
                case "getMessage":
                  return await getMessage(args.messageId, args.folderPath);
                case "searchContacts":
                  return searchContacts(args.query || "");
                case "listCalendars":
                  return listCalendars();
                case "listAccounts":
                  return listAccounts();
                case "listFolders":
                  return listFolders(args.accountKey);
                case "getRecentMessages":
                  return getRecentMessages(args.folderPath, args.limit, args.daysBack, args.unreadOnly);
                case "sendMail":
                  return composeMail(args.to, args.subject, args.body, args.cc, args.isHtml);
                case "replyToMessage":
                  return replyToMessage(args.messageId, args.folderPath, args.body, args.replyAll, args.isHtml);
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
            return { success: false, error: e.toString() };
          }
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
