/**
 * Mail domain module for the Thunderbird MCP server.
 *
 * Implements the message and folder tools: searchMessages, getMessage,
 * getMessageHeaders, getRecentMessages, batchGetMessageHeaders, searchByThread,
 * searchAttachments, getSenderHistory, displayMessage, updateMessage,
 * deleteMessages, exportMailbox, refreshFolder, and the folder-CRUD tools
 * createFolder / renameFolder / deleteFolder / moveFolder / emptyTrash /
 * emptyJunk -- plus their internal helpers (exportsDir, buildExportFilename,
 * glodaBodySearch, msgHdrToHeaderObject, findTrashFolder, findSpecialFolder,
 * deleteAllMessagesRecursive, htmlToMarkdown, extractFormattedBody) and the
 * export-related constants.
 *
 * The infra tools listAccounts / listFolders / getAccountAccess stay in api.js
 * because they are connection/access-control surface, not message operations.
 *
 * Loaded by api.js via Services.scriptloader.loadSubScript with a
 * { module: { exports: {} } } scope, exactly like security_helpers.js and the
 * other domain modules. register(ctx) destructures the shared dependencies it
 * needs from ctx -- the XPCOM bindings (Cc/Ci/Services/MailServices/NetUtil/
 * GlodaMsgSearcher), the search/result limit constants, the access-control and
 * audit infra helpers, the shared body-extraction helpers (stripHtml,
 * extractBodyContent, extractPlainTextBody) and folder/message lookup helpers
 * (openFolder, findMessage) that COMPOSE also uses, and the untrusted-content
 * sanitizers from security_helpers.js -- then defines the functions verbatim
 * and assigns the tool functions back onto ctx for the dispatch switch.
 *
 * NOTE: function bodies are copied byte-for-byte from the monolithic api.js
 * (original indentation preserved) so runtime behavior is identical.
 */
"use strict";

module.exports = function register(ctx) {
  const {
    Cc,
    Ci,
    Services,
    MailServices,
    NetUtil,
    GlodaMsgSearcher,
    MAX_SEARCH_RESULTS_CAP,
    DEFAULT_MAX_RESULTS,
    AUDIT_LOG_SUBDIR,
    paginate,
    appendComposeAudit,
    getUserTags,
    isAccountAllowed,
    isFolderAccessible,
    isMailboxExportBlocked,
    getAccessibleFolder,
    getAccessibleAccounts,
    openFolder,
    findMessage,
    stripHtml,
    extractBodyContent,
    extractPlainTextBody,
    isSafeImageSrc,
    isSafeMarkdownHref,
    isSystemPrincipalFetchAllowed,
    escapeMarkdownLinkText,
    renderMarkdownLink,
    wrapUntrustedBody,
    wrapUntrustedPreview,
  } = ctx;

  // ── Export-related constants (used only by the mailbox-export and search
  // tools, so they live with their consumers rather than in api.js). ──
  const EXPORTS_SUBDIR = "exports";
  const EXPORT_DEFAULT_MAX = 1000;
  const EXPORT_HARD_CAP = 50000;
  const EXPORT_DEFAULT_SCAN_CAP = 50000;
  const SEARCH_COLLECTION_CAP = 10000;

            function exportsDir() {
              const profDir = Services.dirsvc.get("ProfD", Ci.nsIFile);
              const d = profDir.clone();
              d.append(AUDIT_LOG_SUBDIR);
              d.append(EXPORTS_SUBDIR);
              return d;
            }

            /**
             * Build a safe destination filename for an export: ISO timestamp
             * with `:` replaced (Windows-hostile) and a sanitized folder
             * component derived from the folder URI's tail.
             */
            function buildExportFilename(folder) {
              const ts = new Date().toISOString().replace(/[:]/g, "-");
              const tail = folder.URI ? folder.URI.split("/").pop() : "folder";
              const safeTail = String(tail || "folder").replace(/[^a-zA-Z0-9._-]/g, "_").slice(0, 80);
              return `${ts}-${safeTail || "folder"}.jsonl`;
            }

            /**
             * Async streaming export. Resolves once every message has been
             * serialized and flushed to disk.
             */
            function exportMailbox(folderPath, maxMessages, includeBody, includeAttachmentMeta, scanCap, sortOrder) {
              return new Promise((resolve) => {
                try {
                  if (isMailboxExportBlocked()) {
                    resolve({ error: "User preference blocks mailbox export via MCP. Enable 'Allow mailbox export' in the extension options page if you trust this MCP client to bulk-export your mail." });
                    return;
                  }
                  if (!folderPath) { resolve({ error: "folderPath is required" }); return; }
                  const opened = openFolder(folderPath);
                  if (opened.error) { resolve({ error: opened.error }); return; }
                  const { folder, db } = opened;

                  const cap = Number.isFinite(maxMessages) && maxMessages > 0
                    ? Math.min(Math.floor(maxMessages), EXPORT_HARD_CAP)
                    : EXPORT_DEFAULT_MAX;
                  const scanLimit = Number.isFinite(scanCap) && scanCap > 0
                    ? Math.min(Math.floor(scanCap), EXPORT_HARD_CAP)
                    : EXPORT_DEFAULT_SCAN_CAP;
                  const order = sortOrder === "asc" ? "asc" : "desc";

                  // Walk + collect candidate headers with their timestamps.
                  // We sort BEFORE applying `cap` so "newest N" semantics are
                  // honored across the whole folder, not just the first N
                  // enumerated. enumerateMessages doesn't guarantee an order.
                  const candidates = [];
                  let totalScanned = 0;
                  for (const hdr of db.enumerateMessages()) {
                    totalScanned++;
                    if (totalScanned > scanLimit) break;
                    candidates.push({ hdr, ts: hdr.date ? hdr.date / 1000 : 0 });
                  }
                  candidates.sort((a, b) => order === "asc" ? a.ts - b.ts : b.ts - a.ts);
                  const slice = candidates.slice(0, cap);

                  // Open the output file.
                  const dir = exportsDir();
                  if (!dir.exists()) {
                    dir.create(Ci.nsIFile.DIRECTORY_TYPE, 0o700);
                  }
                  const outFile = dir.clone();
                  outFile.append(buildExportFilename(folder));
                  const ostream = Cc["@mozilla.org/network/file-output-stream;1"]
                    .createInstance(Ci.nsIFileOutputStream);
                  // 0x02 = O_WRONLY, 0x08 = O_CREAT, 0x20 = O_TRUNC. We always
                  // start fresh -- the filename has an ISO timestamp so
                  // collisions are essentially impossible.
                  ostream.init(outFile, 0x02 | 0x08 | 0x20, 0o600, 0);
                  const converter = Cc["@mozilla.org/intl/converter-output-stream;1"]
                    .createInstance(Ci.nsIConverterOutputStream);
                  converter.init(ostream, "UTF-8");

                  let bytesWritten = 0;
                  function writeLine(obj) {
                    const line = JSON.stringify(obj) + "\n";
                    converter.writeString(line);
                    bytesWritten += line.length;
                  }

                  // Synchronous header-only path is straightforward.
                  if (!includeBody) {
                    for (const { hdr } of slice) {
                      const obj = msgHdrToHeaderObject(hdr);
                      obj.folderPath = folder.URI;
                      writeLine(obj);
                    }
                    converter.close();
                    appendComposeAudit({
                      tool: "exportMailbox",
                      folderPath: folder.URI,
                      includeBody: false,
                      exported: slice.length,
                      scanned: totalScanned,
                      filePath: outFile.path,
                      bytesWritten,
                    });
                    resolve({
                      filePath: outFile.path,
                      folderPath: folder.URI,
                      exported: slice.length,
                      scanned: totalScanned,
                      truncated: totalScanned > scanLimit || candidates.length > cap,
                      bytesWritten,
                      includeBody: false,
                    });
                    return;
                  }

                  // With-body path: MIME-parse one at a time, write line, then
                  // schedule the next. Sequential to keep memory flat -- a
                  // parallel fan-out would buffer N MimeMessage trees.
                  const { MsgHdrToMimeMessage } = ChromeUtils.importESModule(
                    "resource:///modules/gloda/MimeMessage.sys.mjs"
                  );

                  let index = 0;
                  function next() {
                    if (index >= slice.length) {
                      converter.close();
                      appendComposeAudit({
                        tool: "exportMailbox",
                        folderPath: folder.URI,
                        includeBody: true,
                        exported: index,
                        scanned: totalScanned,
                        filePath: outFile.path,
                        bytesWritten,
                      });
                      resolve({
                        filePath: outFile.path,
                        folderPath: folder.URI,
                        exported: index,
                        scanned: totalScanned,
                        truncated: totalScanned > scanLimit || candidates.length > cap,
                        bytesWritten,
                        includeBody: true,
                      });
                      return;
                    }
                    const { hdr } = slice[index++];
                    const obj = msgHdrToHeaderObject(hdr);
                    obj.folderPath = folder.URI;
                    MsgHdrToMimeMessage(hdr, null, (aMsgHdr, aMimeMsg) => {
                      try {
                        if (aMimeMsg) {
                          obj.body = extractPlainTextBody(aMimeMsg);
                          if (includeAttachmentMeta && aMimeMsg.allUserAttachments) {
                            obj.attachments = aMimeMsg.allUserAttachments.map(a => ({
                              name: a && a.name ? String(a.name) : "",
                              contentType: a && a.contentType ? String(a.contentType) : "",
                              size: typeof a?.size === "number" ? a.size : null,
                            }));
                          }
                        }
                      } catch (e) {
                        obj.bodyError = String(e);
                      }
                      try { writeLine(obj); }
                      catch (e) {
                        converter.close();
                        resolve({ error: `Write failed at message ${index}: ${e}`, filePath: outFile.path, exported: index - 1, bytesWritten });
                        return;
                      }
                      next();
                    });
                  }
                  next();
                } catch (e) {
                  resolve({ error: e.toString() });
                }
              });
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
                          if (preview) result.preview = wrapUntrustedPreview(preview);
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
	              // Gloda full-body search path (async).
	              //
	              // SECURITY: GlodaMsgSearcher passes the query through to a
	              // tokenizer that supports boolean operators (AND / OR / NOT
	              // / *). We restrict the API surface to plain keyword search
	              // so an MCP caller can't (a) probe the index via wildcards
	              // they don't have rights to, (b) write surprisingly broad
	              // queries that scan the whole mailbox. Reject any token-
	              // level Gloda operator with a clear error rather than
	              // silently stripping -- the caller deserves to know their
	              // query was mangled.
	              if (searchBody) {
	                if (!GlodaMsgSearcher) return { error: "Gloda full-text index is not available" };
	                if (!query) return { error: "searchBody requires a non-empty query" };
	                const dangerous = /(?:^|\s)(?:AND|OR|NOT)(?:\s|$)|[*"()]/;
	                if (dangerous.test(query)) {
	                  return {
	                    error: "searchBody query must be plain keywords. Gloda operators (AND, OR, NOT, *, \", parentheses) are rejected by this MCP API.",
	                  };
	                }
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
                    if (preview) result.preview = wrapUntrustedPreview(preview);
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

	            /**
	             * Cheap headers-only variant of getMessage. Returns the same
	             * structural fields searchMessages returns plus references and
	             * in-reply-to, without paying the MIME parse / body decode cost.
	             * Useful when the LLM is triaging a list and only needs to know
	             * "who, when, what subject" before deciding whether to fetch full.
	             */
	            function getMessageHeaders(messageId, folderPath) {
	              const found = findMessage(messageId, folderPath);
	              if (found.error) return { error: found.error };
	              return msgHdrToHeaderObject(found.msgHdr);
	            }

	            // Shared shape extraction so getMessageHeaders and
	            // batchGetMessageHeaders return identical fields.
	            function msgHdrToHeaderObject(msgHdr) {
	              return {
	                id: msgHdr.messageId,
	                subject: msgHdr.mime2DecodedSubject || msgHdr.subject || "",
	                author: msgHdr.mime2DecodedAuthor || msgHdr.author || "",
	                recipients: msgHdr.mime2DecodedRecipients || msgHdr.recipients || "",
	                ccList: msgHdr.ccList || "",
	                date: msgHdr.date ? new Date(msgHdr.date / 1000).toISOString() : null,
	                tags: getUserTags(msgHdr),
	                isRead: msgHdr.isRead,
	                isFlagged: msgHdr.isFlagged,
	                threadId: msgHdr.threadId ? String(msgHdr.threadId) : null,
	                references: msgHdr.getStringProperty("references") || "",
	                inReplyTo: msgHdr.getStringProperty("in-reply-to") || "",
	                size: typeof msgHdr.messageSize === "number" ? msgHdr.messageSize : null,
	              };
	            }

	            /**
	             * Recent messages from a given sender, across the inbox of
	             * every accessible account. Uses substring match against the
	             * author field (which contains "Name <email>" -- so email
	             * matches naturally). Newest-first.
	             */
	            function getSenderHistory(email, maxResults, scanCap, sinceDays) {
	              if (typeof email !== "string" || !email.trim()) {
	                return { error: "email must be a non-empty string" };
	              }
	              const wantLower = email.toLowerCase();
	              const cap = Number.isFinite(maxResults) && maxResults > 0
	                ? Math.min(Math.floor(maxResults), 200)
	                : 50;
	              const scanLimit = Number.isFinite(scanCap) && scanCap > 0
	                ? Math.min(Math.floor(scanCap), 50000)
	                : 5000;
	              const sinceMicros = Number.isFinite(sinceDays) && sinceDays > 0
	                ? (Date.now() - sinceDays * 86400 * 1000) * 1000
	                : null;

	              const FLAG_INBOX = 0x1000;
	              const results = [];
	              let totalScanned = 0;
	              for (const account of getAccessibleAccounts()) {
	                if (results.length >= cap) break;
	                let inbox = null;
	                try {
	                  const root = account.incomingServer && account.incomingServer.rootFolder;
	                  if (!root) continue;
	                  if (typeof root.getFolderWithFlags === "function") {
	                    try { inbox = root.getFolderWithFlags(FLAG_INBOX); } catch { /* ignore */ }
	                  }
	                  if (!inbox) inbox = root;
	                  const db = inbox.msgDatabase;
	                  if (!db) continue;
	                  let scanned = 0;
	                  for (const hdr of db.enumerateMessages()) {
	                    scanned++;
	                    totalScanned++;
	                    if (scanned > scanLimit) break;
	                    if (sinceMicros !== null && hdr.date && hdr.date < sinceMicros) continue;
	                    const author = (hdr.mime2DecodedAuthor || hdr.author || "").toLowerCase();
	                    if (!author.includes(wantLower)) continue;
	                    const obj = msgHdrToHeaderObject(hdr);
	                    obj.folderPath = inbox.URI;
	                    obj._ts = hdr.date ? hdr.date / 1000 : 0;
	                    results.push(obj);
	                    if (results.length >= cap * 2) break; // generous pre-sort buffer
	                  }
	                } catch { /* skip inaccessible account */ }
	              }

	              results.sort((a, b) => (b._ts || 0) - (a._ts || 0));
	              const trimmed = results.slice(0, cap).map(({ _ts, ...rest }) => rest);
	              return {
	                sender: email,
	                messages: trimmed,
	                totalScanned,
	                truncated: results.length > cap,
	              };
	            }

	            /**
	             * Find messages with attachments matching a filename substring
	             * and/or MIME-type prefix. Works by pre-filtering to messages
	             * that have the Attachment flag set, then walking
	             * allUserAttachments via MsgHdrToMimeMessage. Async because
	             * MIME parse is async.
	             *
	             * Scoped to a specific folder, or fans out across the inbox of
	             * every accessible account when folderPath is omitted.
	             */
	            function searchAttachments(nameContains, contentType, folderPath, maxResults, scanCap) {
	              return new Promise((resolve) => {
	                try {
	                  if (!nameContains && !contentType) {
	                    resolve({ error: "Provide at least one of nameContains or contentType" });
	                    return;
	                  }
	                  const wantName = nameContains ? String(nameContains).toLowerCase() : null;
	                  const wantType = contentType ? String(contentType).toLowerCase() : null;
	                  const cap = Number.isFinite(maxResults) && maxResults > 0
	                    ? Math.min(Math.floor(maxResults), 200)
	                    : 50;
	                  const scanLimit = Number.isFinite(scanCap) && scanCap > 0
	                    ? Math.min(Math.floor(scanCap), 50000)
	                    : 5000;

	                  // Build the list of folders to scan.
	                  const folders = [];
	                  if (folderPath) {
	                    const opened = openFolder(folderPath);
	                    if (opened.error) { resolve({ error: opened.error }); return; }
	                    folders.push(opened.folder);
	                  } else {
	                    for (const account of getAccessibleAccounts()) {
	                      try {
	                        const root = account.incomingServer && account.incomingServer.rootFolder;
	                        if (!root) continue;
	                        // Pick the inbox if present, else the root.
	                        let inbox = null;
	                        if (typeof root.getFolderWithFlags === "function") {
	                          try {
	                            const FLAG_INBOX = 0x1000; // nsMsgFolderFlags.Inbox
	                            inbox = root.getFolderWithFlags(FLAG_INBOX);
	                          } catch { /* ignore */ }
	                        }
	                        folders.push(inbox || root);
	                      } catch { /* skip inaccessible accounts */ }
	                    }
	                  }

	                  // Collect candidate headers (those with the Attachment flag).
	                  const FLAG_ATTACH = 0x10000000; // nsMsgMessageFlags.Attachment
	                  const candidates = [];
	                  let totalScanned = 0;
	                  for (const folder of folders) {
	                    if (candidates.length >= cap * 4) break; // generous pre-MIME buffer
	                    let db;
	                    try { db = folder.msgDatabase; } catch { continue; }
	                    if (!db) continue;
	                    let scannedHere = 0;
	                    for (const hdr of db.enumerateMessages()) {
	                      scannedHere++;
	                      totalScanned++;
	                      if (scannedHere > scanLimit) break;
	                      if ((hdr.flags & FLAG_ATTACH) !== 0) {
	                        candidates.push({ hdr, folder });
	                        if (candidates.length >= cap * 4) break;
	                      }
	                    }
	                  }

	                  if (candidates.length === 0) {
	                    resolve({ matches: [], totalScanned, candidates: 0 });
	                    return;
	                  }

	                  // MIME-parse each candidate; collect matches.
	                  const { MsgHdrToMimeMessage } = ChromeUtils.importESModule(
	                    "resource:///modules/gloda/MimeMessage.sys.mjs"
	                  );

	                  const matches = [];
	                  let remaining = candidates.length;
	                  const finish = () => {
	                    matches.sort((a, b) => (b._ts || 0) - (a._ts || 0));
	                    const trimmed = matches.slice(0, cap).map(({ _ts, ...rest }) => rest);
	                    resolve({
	                      matches: trimmed,
	                      totalScanned,
	                      candidates: candidates.length,
	                      truncated: matches.length > cap,
	                    });
	                  };

	                  for (const { hdr, folder } of candidates) {
	                    MsgHdrToMimeMessage(hdr, null, (aMsgHdr, aMimeMsg) => {
	                      try {
	                        const atts = aMimeMsg && aMimeMsg.allUserAttachments ? aMimeMsg.allUserAttachments : [];
	                        const matched = [];
	                        for (const att of atts) {
	                          const name = (att && att.name) ? String(att.name).toLowerCase() : "";
	                          const ct = (att && att.contentType) ? String(att.contentType).toLowerCase() : "";
	                          const nameOk = !wantName || name.includes(wantName);
	                          const typeOk = !wantType || ct.startsWith(wantType);
	                          if (nameOk && typeOk) {
	                            matched.push({
	                              name: att.name || "",
	                              contentType: att.contentType || "",
	                              size: typeof att.size === "number" ? att.size : null,
	                            });
	                          }
	                        }
	                        if (matched.length > 0) {
	                          const headerObj = msgHdrToHeaderObject(hdr);
	                          headerObj.matchingAttachments = matched;
	                          headerObj._ts = hdr.date ? hdr.date / 1000 : 0;
	                          headerObj.folderPath = folder.URI;
	                          matches.push(headerObj);
	                        }
	                      } catch { /* skip parse errors */ }
	                      remaining--;
	                      if (remaining === 0) finish();
	                    });
	                  }
	                } catch (e) {
	                  resolve({ error: e.toString() });
	                }
	              });
	            }

	            /**
	             * Walk a folder once and return all headers that share the
	             * threadId of the anchor message. The anchor itself is included.
	             * Newest-first.
	             */
	            function searchByThread(messageId, folderPath, maxResults) {
	              const found = findMessage(messageId, folderPath);
	              if (found.error) return { error: found.error };
	              const { msgHdr: anchor, folder, db } = found;
	              const threadId = anchor.threadId;
	              if (!threadId) {
	                return { thread: [msgHdrToHeaderObject(anchor)], threadId: null, anchored: true };
	              }
	              const cap = Number.isFinite(maxResults) && maxResults > 0
	                ? Math.min(Math.floor(maxResults), 500)
	                : 100;
	              const results = [];
	              const SCAN_CAP = 50000;
	              let scanned = 0;
	              for (const hdr of db.enumerateMessages()) {
	                scanned++;
	                if (scanned > SCAN_CAP) break;
	                if (hdr.threadId === threadId) {
	                  results.push({ _hdr: hdr, _ts: hdr.date ? hdr.date / 1000 : 0 });
	                  if (results.length >= cap) break;
	                }
	              }
	              results.sort((a, b) => b._ts - a._ts);
	              return {
	                threadId: String(threadId),
	                anchorId: anchor.messageId,
	                thread: results.map(r => msgHdrToHeaderObject(r._hdr)),
	                truncated: results.length >= cap,
	                scanned,
	              };
	            }

	            /**
	             * Resolve N message IDs against a single folder in one pass.
	             * Returns { headers: { [id]: headerObj | {error} }, total, failed }.
	             * Hard cap of 200 -- search-result pages are typically smaller
	             * and this keeps the round-trip JSON bounded.
	             */
	            function batchGetMessageHeaders(messageIds, folderPath) {
	              if (!Array.isArray(messageIds)) return { error: "messageIds must be an array" };
	              if (messageIds.length === 0) return { headers: {}, total: 0, failed: 0 };
	              if (messageIds.length > 200) {
	                return { error: `Too many ids: ${messageIds.length}. Hard cap is 200; call multiple times if needed.` };
	              }
	              const opened = openFolder(folderPath);
	              if (opened.error) return { error: opened.error };
	              const { folder, db } = opened;

	              // Build a single id -> hdr map by enumerating ONCE rather than
	              // calling findMessage N times. For an inbox with M messages
	              // this is O(M) instead of O(M*N).
	              const wanted = new Set(messageIds);
	              const found = new Map();
	              const hasDirect = typeof db.getMsgHdrForMessageID === "function";
	              if (hasDirect) {
	                for (const id of messageIds) {
	                  try {
	                    const h = db.getMsgHdrForMessageID(id);
	                    if (h) found.set(id, h);
	                  } catch { /* ignore */ }
	                }
	              }
	              // Anything still unresolved gets one linear enumeration pass.
	              const missing = messageIds.filter(id => !found.has(id));
	              if (missing.length > 0) {
	                const missSet = new Set(missing);
	                let scanned = 0;
	                const SCAN_CAP = 50000;
	                for (const hdr of db.enumerateMessages()) {
	                  scanned++;
	                  if (scanned > SCAN_CAP) break;
	                  if (missSet.has(hdr.messageId)) {
	                    found.set(hdr.messageId, hdr);
	                    missSet.delete(hdr.messageId);
	                    if (missSet.size === 0) break;
	                  }
	                }
	              }

	              const headers = Object.create(null);
	              let failed = 0;
	              for (const id of messageIds) {
	                const h = found.get(id);
	                if (!h) {
	                  headers[id] = { error: "not found in this folder" };
	                  failed++;
	                } else {
	                  headers[id] = msgHdrToHeaderObject(h);
	                }
	              }
	              return { headers, total: messageIds.length, failed };
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
                          // Find header/body boundary. Prefer CRLFCRLF (RFC 5322),
                          // then LFLF (LF-normalized mbox), then CRCR (legacy
                          // classic Mac exports). Pick the earliest match so a
                          // stray LFLF inside CRLF-separated headers doesn't win.
                          const boundaryMatch = rawContent.match(/\r\n\r\n|\n\n|\r\r/);
                          const headerEnd = boundaryMatch ? boundaryMatch.index : -1;
                          const bodyStart = boundaryMatch
                            ? boundaryMatch.index + boundaryMatch[0].length
                            : -1;
                          if (bodyStart < 0) {
                            console.error(`${fallbackContext}: could not find header/body boundary`);
                          } else {
                            const headerBlock = rawContent.slice(0, headerEnd);
                            const rawBody = rawContent.slice(bodyStart);
                            // Unfold continuation lines for all three line-ending flavors.
                            const unfoldedHeaders = headerBlock
                              .replace(/(?:\r\n|\r|\n)[ \t]+/g, " ");
                            let contentTypeHeader = "";
                            let transferEncodingHeader = "";
                            for (const line of unfoldedHeaders.split(/\r\n|\r|\n/)) {
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
                                // Remove quoted-printable soft breaks: =CRLF, =LF, =CR.
                                const qpBody = rawBody.replace(/=(?:\r\n|\r|\n)/g, "");
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
                            // Append part parameter to the resolved fetchable URL.
                            // URL-encode partName: it is structural (e.g. "1.2.1")
                            // in current Thunderbird but encoding it costs nothing
                            // and closes any future regression where a non-digit
                            // character could land inside a URL query value.
                            const sep = baseUri.spec.includes("?") ? "&" : "?";
                            partUrl = `${baseUri.spec}${sep}part=${encodeURIComponent(part.partName)}`;
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

                          // SECURITY: the channel uses the system principal so
                          // mail-store protocols (mailbox:, imap-message:, ...)
                          // can be fetched without sandboxing. That privilege
                          // would let a file:// or chrome:// URL slipping into
                          // `url` exfiltrate arbitrary local files or read
                          // internal browser resources, so we hard-require an
                          // allow-listed mail-store scheme here. This is
                          // defense in depth: Thunderbird itself supplies the
                          // URLs we feed into this code path, but if a future
                          // regression or extension interaction ever returned
                          // a non-mail scheme, this check refuses to fetch it.
                          if (!isSystemPrincipalFetchAllowed(url)) {
                            info.error = `Refusing to fetch attachment from non-mail-store URL: ${String(url).slice(0, 80)}`;
                            try { file.remove(false); } catch {}
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

            /**
             * Force the IMAP server-side state for `folderPath` to sync into
             * Thunderbird's local cache. Non-IMAP folders short-circuit with
             * success since there's nothing to fetch. Returns when the
             * IMAP URL listener fires onStopRunningUrl (success or error),
             * or when the timeout elapses.
             *
             * Without this, recently-arrived mail in IMAP folders is invisible
             * to searchMessages / getRecentMessages until the user clicks the
             * folder in Thunderbird's UI. README documents this as a known
             * issue; refreshFolder is the programmatic fix.
             */
            function refreshFolder(folderPath, timeoutMs) {
              return new Promise((resolve) => {
                try {
                  const result = getAccessibleFolder(folderPath);
                  if (result.error) { resolve({ error: result.error }); return; }
                  const folder = result.folder;
                  // Non-IMAP folders (Local Folders, news, etc.) don't need
                  // server-side refresh -- just return success.
                  if (!folder.server || folder.server.type !== "imap") {
                    resolve({ success: true, folderPath: folder.URI, skipped: "not an IMAP folder" });
                    return;
                  }

                  // Stay under the bridge's HTTP REQUEST_TIMEOUT (30000ms). A
                  // tool that took the full 30s would race the bridge into a
                  // hard-fail "Request to Thunderbird timed out", losing the
                  // structured timeout result we'd otherwise return. 25s gives
                  // a 5s safety margin for the round-trip stdio I/O.
                  const REFRESH_TIMEOUT_CAP_MS = 25000;
                  const limit = Number.isFinite(timeoutMs) && timeoutMs > 0
                    ? Math.min(Math.floor(timeoutMs), REFRESH_TIMEOUT_CAP_MS)
                    : 15000;

                  let settled = false;
                  const settle = (value) => {
                    if (settled) return;
                    settled = true;
                    try { timer.cancel(); } catch {}
                    resolve(value);
                  };

                  const timer = Cc["@mozilla.org/timer;1"].createInstance(Ci.nsITimer);
                  timer.initWithCallback(
                    { notify() { settle({ error: `refreshFolder timed out after ${limit}ms`, folderPath: folder.URI }); } },
                    limit,
                    Ci.nsITimer.TYPE_ONE_SHOT
                  );

                  const beforeCount = folder.getTotalMessages(false);
                  const urlListener = {
                    QueryInterface: ChromeUtils.generateQI(["nsIUrlListener"]),
                    OnStartRunningUrl() {},
                    OnStopRunningUrl(url, exitCode) {
                      const ok = Components.isSuccessCode(exitCode);
                      const afterCount = (() => {
                        try { return folder.getTotalMessages(false); } catch { return null; }
                      })();
                      const newMessages = (afterCount !== null && typeof beforeCount === "number")
                        ? Math.max(0, afterCount - beforeCount)
                        : null;
                      if (ok) {
                        settle({
                          success: true,
                          folderPath: folder.URI,
                          totalBefore: beforeCount,
                          totalAfter: afterCount,
                          newMessages,
                        });
                      } else {
                        settle({
                          error: `IMAP fetch failed (status 0x${exitCode.toString(16)})`,
                          folderPath: folder.URI,
                        });
                      }
                    },
                  };

                  try {
                    folder.getNewMessages(null, urlListener);
                  } catch (e) {
                    settle({ error: e.toString(), folderPath: folder.URI });
                  }
                } catch (e) {
                  resolve({ error: e.toString() });
                }
              });
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
                    if (preview) result.preview = wrapUntrustedPreview(preview);
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
              } catch (e) {
                // Continue traversal; log per-folder failures so partial empties are visible.
                console.error("thunderbird-mcp: deleteMessages failed for folder", folder?.URI || folder?.name, ":", e);
              }
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

            /**
             * Converts HTML to markdown using DOMParser for structure-preserving
             * body extraction. Handles headings, links, bold/italic, lists,
             * blockquotes, code blocks, images, and horizontal rules. Email
             * tables (usually layout, not data) are flattened to text.
             * Falls back to stripHtml if DOMParser is unavailable.
             *
             * SECURITY: `<a href>` and `<img src>` values come from sender-
             * controlled HTML, so we strip dangerous schemes (javascript:,
             * data: non-image, etc.) and defang markdown-syntax injection in
             * the link text before emitting the final markdown.
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
                      const rawHref = node.getAttribute("href") || "";
                      const text = inner().trim();
                      // Skip empty/anchor-only links
                      if (!text && !rawHref) return "";
                      // Drop unsafe href schemes (javascript:, data: non-image,
                      // vbscript:, file:, chrome:, jar:, blob:, ...). Fall back
                      // to the visible text so the message stays readable.
                      if (rawHref && !isSafeMarkdownHref(rawHref)) {
                        return escapeMarkdownLinkText(text || rawHref);
                      }
                      if (rawHref && text && text !== rawHref) {
                        return renderMarkdownLink(text, rawHref);
                      }
                      return escapeMarkdownLinkText(text || rawHref);
                    }
                    case "img": {
                      const alt = node.getAttribute("alt") || "";
                      const rawSrc = node.getAttribute("src") || "";
                      // Skip tracking pixels (1x1, tiny, or data: without alt)
                      const w = parseInt(node.getAttribute("width")) || 0;
                      const h = parseInt(node.getAttribute("height")) || 0;
                      if ((w > 0 && w <= 3) || (h > 0 && h <= 3)) return "";
                      if (rawSrc.startsWith("data:") && !alt) return "";
                      // Allow http(s):, cid:, mailto: (rare), and data:image/*.
                      // Anything else (javascript:, data: non-image, file:, ...)
                      // is dropped to alt text.
                      if (rawSrc && !isSafeImageSrc(rawSrc)) {
                        return escapeMarkdownLinkText(alt);
                      }
                      if (!rawSrc) return escapeMarkdownLinkText(alt);
                      const safeAlt = escapeMarkdownLinkText(alt);
                      if (rawSrc.includes(">")) return safeAlt;
                      const srcForMd = /[()\s]/.test(rawSrc) ? `<${rawSrc}>` : rawSrc;
                      return `![${safeAlt}](${srcForMd})`;
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
             * Extracts body from a MIME message in the requested format.
             * For "text": uses coerceBodyToPlaintext fast path (original behavior).
             * For "markdown"/"html": walks MIME tree to find raw HTML content.
             */
            function extractFormattedBody(aMimeMsg, bodyFormat) {
              if (bodyFormat === "text") {
                return { body: wrapUntrustedBody(extractPlainTextBody(aMimeMsg)), bodyIsHtml: false };
              }
              // For markdown/html: need raw MIME content, not coerced text
              const { text, isHtml } = extractBodyContent(aMimeMsg);
              if (!text) {
                // MIME tree empty -- try coerce as last resort
                const fallback = extractPlainTextBody(aMimeMsg);
                return { body: wrapUntrustedBody(fallback), bodyIsHtml: false };
              }
              if (!isHtml) return { body: wrapUntrustedBody(text), bodyIsHtml: false };
              if (bodyFormat === "html") return { body: text, bodyIsHtml: true };
              // Default: markdown
              return { body: wrapUntrustedBody(htmlToMarkdown(text)), bodyIsHtml: false };
            }

  Object.assign(ctx, {
    exportsDir,
    buildExportFilename,
    exportMailbox,
    glodaBodySearch,
    searchMessages,
    getMessageHeaders,
    msgHdrToHeaderObject,
    getSenderHistory,
    searchAttachments,
    searchByThread,
    batchGetMessageHeaders,
    getMessage,
    findTrashFolder,
    displayMessage,
    refreshFolder,
    getRecentMessages,
    deleteMessages,
    updateMessage,
    createFolder,
    renameFolder,
    deleteFolder,
    findSpecialFolder,
    deleteAllMessagesRecursive,
    emptyTrash,
    emptyJunk,
    moveFolder,
    htmlToMarkdown,
    extractFormattedBody,
  });
};
