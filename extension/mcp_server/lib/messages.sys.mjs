/**
 * Message search, retrieval, and mutation for the Thunderbird MCP server.
 *
 * Largest module in the modular split. Bundles:
 *   - searchMessages / searchFolder / glodaBodySearch — folder + Gloda search
 *   - getMessage / findMessage — full retrieval with MIME parsing,
 *     attachment off-line cache fallback, inline-image discovery
 *   - getRecentMessages — date-window listing with subfolder recursion
 *   - displayMessage — open the message in 3pane / tab / window
 *   - deleteMessages — IMAP-aware deletion with Drafts->Trash redirection
 *   - updateMessage — read/flagged/tag/move/trash mutation
 *   - getUserTags / markMessageDispositionState — small header helpers
 *     also reused by the compose flow (so they stay exported here, not
 *     module-private; a future compose extraction should import them
 *     from this module instead of redefining).
 *
 * Factory pattern: api.js calls makeMessages({ ... }) once per start()
 * invocation. XPCOM globals (MailServices / NetUtil / Services / Cc /
 * Ci) are sandbox-only in api.js and MUST be passed in. Access-control
 * hooks (getAccessibleFolder / isFolderAccessible / getAccessibleAccounts)
 * are supplied by the access-control module created earlier in start().
 * openFolder is supplied by the folders module — it is the IMAP-aware
 * superset of getAccessibleFolder.
 *
 * GlodaMsgSearcher may be null on Thunderbird builds without the Gloda
 * full-text index. searchMessages with searchBody=true returns a polite
 * error in that case; the regular header-search path keeps working.
 *
 * MIME / encoding helpers (extractBodyContent, extractPlainTextBody,
 * extractFormattedBody, htmlToMarkdown, stripHtml) are intentionally
 * module-private here. api.js still owns its own copies because the
 * compose / reply / forward flow also uses them; a future commit should
 * split those into a shared encoding.sys.mjs.
 *
 * Internal `_dateTs` markers are stripped by paginate before results are
 * returned to the caller. The function is copied private here rather
 * than injected because no module outside messages currently uses it.
 *
 * findTrashFolder is also duplicated from folders.sys.mjs (where it is
 * module-private) — the folders module does not export it. Keep both
 * copies in sync until/unless someone exports it explicitly.
 */

export function makeMessages({
  MailServices,
  NetUtil,
  Services,
  Cc,
  Ci,
  GlodaMsgSearcher,                              // may be null on TB without Gloda
  isFolderAccessible,
  getAccessibleFolder,
  getAccessibleAccounts,
  openFolder,                                    // IMAP-aware folder+db opener from folders.sys.mjs
  SEARCH_COLLECTION_CAP,
  MAX_SEARCH_RESULTS_CAP,
  DEFAULT_MAX_RESULTS,
  INTERNAL_KEYWORDS,
  _tempAttachFiles,                              // shared mutable state from state.sys.mjs (currently unused here; reserved for future inline-attachment temp tracking)
  tracer,                                        // structured tracer (lib/trace.sys.mjs); optional -- code paths that use it must guard against null
}) {
  // ── Module-private helpers ──

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
   * Find the Trash folder for the account that owns the given folder.
   * Walks the account tree looking for a folder with the TRASH flag
   * (0x00000100) set; falls back to a folder named "trash" or
   * "deleted items" if no flagged folder is found.
   *
   * Duplicated from folders.sys.mjs (which keeps it module-private).
   * Keep both copies in sync.
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

  // ── MIME / encoding helpers (module-private; future encoding.sys.mjs candidates) ──

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

  // ── Public API ──

  /** Returns user-visible tag keywords from a message header, filtering out internal IMAP flags. */
  function getUserTags(msgHdr) {
    return (msgHdr.getStringProperty("keywords") || "").split(/\s+/).filter(k => k && !INTERNAL_KEYWORDS.has(k.toLowerCase()));
  }

  function markMessageDispositionState(msgHdr, dispositionState) {
    try {
      const folder = msgHdr?.folder;
      if (!folder || dispositionState == null) return false;
      if (typeof folder.addMessageDispositionState === "function") {
        folder.addMessageDispositionState(msgHdr, dispositionState);
        return true;
      }
      if (typeof folder.AddMessageDispositionState === "function") {
        folder.AddMessageDispositionState(msgHdr, dispositionState);
        return true;
      }
    } catch {}
    return false;
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

  // ── Aggregation-first tools ─────────────────────────────────────────────
  // These tools run the same folder-walk as searchMessages but never return
  // raw per-message headers. Designed to keep MCP client context small when
  // the answer is "count and bucket" rather than "show me each header".

  // Compiles query / date / unread / tag filters into a callback that returns
  // true for messages the caller wants. Shared by inbox_inventory and
  // bulk_move_by_query so they stay byte-faithful to searchMessages semantics.
  function _compileMessageFilter({ query, startDate, endDate, unreadOnly, flaggedOnly, tag }) {
    const lowerQuery = (query || "").toLowerCase();
    const hasQuery = !!lowerQuery;
    const OPERATOR_RE = /^(from|subject|to|cc):\s*/;
    let fieldTarget = null;
    let queryTokens = [];
    if (hasQuery) {
      const opMatch = lowerQuery.match(OPERATOR_RE);
      if (opMatch) {
        const opMap = { from: "author", subject: "subject", to: "recipients", cc: "ccList" };
        fieldTarget = opMap[opMatch[1]];
        queryTokens = lowerQuery.slice(opMatch[0].length).trim().split(/\s+/).filter(Boolean);
      } else {
        queryTokens = lowerQuery.split(/\s+/).filter(Boolean);
      }
    }
    const failedQuery = hasQuery && queryTokens.length === 0;
    const parsedStart = startDate ? new Date(startDate).getTime() : NaN;
    const parsedEnd = endDate ? new Date(endDate).getTime() : NaN;
    const startTs = Number.isFinite(parsedStart) ? parsedStart * 1000 : null;
    const isDateOnly = endDate && /^\d{4}-\d{2}-\d{2}$/.test(endDate.trim());
    const endTs = Number.isFinite(parsedEnd) ? (parsedEnd + (isDateOnly ? 86400000 : 0)) * 1000 : null;

    return function matches(msgHdr) {
      if (failedQuery) return false;
      const ts = msgHdr.date || 0;
      if (startTs !== null && ts < startTs) return false;
      if (endTs !== null && ts > endTs) return false;
      if (unreadOnly && msgHdr.isRead) return false;
      if (flaggedOnly && !msgHdr.isFlagged) return false;
      if (tag) {
        const keywords = (msgHdr.getStringProperty("keywords") || "").split(/\s+/);
        if (!keywords.includes(tag)) return false;
      }
      if (hasQuery) {
        const subject = (msgHdr.mime2DecodedSubject || msgHdr.subject || "").toLowerCase();
        const author = (msgHdr.mime2DecodedAuthor || msgHdr.author || "").toLowerCase();
        const recipients = (msgHdr.mime2DecodedRecipients || msgHdr.recipients || "").toLowerCase();
        const ccList = (msgHdr.ccList || "").toLowerCase();
        const preview = (msgHdr.getStringProperty("preview") || "").toLowerCase();
        const fields = { subject, author, recipients, ccList };
        const ok = fieldTarget
          ? queryTokens.every(t => (fields[fieldTarget] || "").includes(t))
          : queryTokens.every(t =>
              subject.includes(t) || author.includes(t) || recipients.includes(t) || ccList.includes(t) || preview.includes(t)
            );
        if (!ok) return false;
      }
      return true;
    };
  }

  function _walkFolderMessages(folder, callback, { includeSubfolders = true, cap = SEARCH_COLLECTION_CAP } = {}) {
    let processed = 0;
    let aborted = false;
    function walk(f) {
      if (aborted) return;
      try {
        if (f.server && f.server.type === "imap") {
          try { f.updateFolder(null); } catch { /* best-effort refresh */ }
        }
        const db = f.msgDatabase;
        if (!db) return;
        for (const msgHdr of db.enumerateMessages()) {
          if (processed >= cap) { aborted = true; return; }
          const cont = callback(msgHdr, f);
          processed++;
          if (cont === false) { aborted = true; return; }
        }
      } catch { /* skip inaccessible folders */ }
      if (!aborted && includeSubfolders && f.hasSubFolders) {
        for (const sub of f.subFolders) {
          if (aborted) break;
          walk(sub);
        }
      }
    }
    walk(folder);
    return { processed, capped: aborted && processed >= cap };
  }

  function _groupKeyFor(msgHdr, groupBy) {
    const author = (msgHdr.mime2DecodedAuthor || msgHdr.author || "").trim();
    if (groupBy === "from_address") {
      const m = author.match(/<([^>]+)>/);
      return (m ? m[1] : author).toLowerCase() || "(unknown)";
    }
    if (groupBy === "subject_prefix") {
      const subj = (msgHdr.mime2DecodedSubject || msgHdr.subject || "").trim();
      return subj.replace(/^(re|fwd|fw|aw|wg):\s*/i, "").slice(0, 30).toLowerCase() || "(no subject)";
    }
    if (groupBy === "tag") {
      const tags = getUserTags(msgHdr);
      return tags.length ? tags.slice().sort().join(",") : "(untagged)";
    }
    if (groupBy === "day") {
      if (!msgHdr.date) return "(no date)";
      return new Date(msgHdr.date / 1000).toISOString().slice(0, 10);
    }
    // default + explicit "from_domain"
    const addr = (author.match(/<([^>]+)>/) || [null, author])[1] || author;
    const m = addr.match(/@([A-Za-z0-9._-]+)/);
    return (m ? m[1] : "(unknown)").toLowerCase();
  }

  /**
   * Aggregate-first query. Returns grouped counts + sample IDs/subjects
   * instead of per-message headers. See the tool description in
   * tools-schema.sys.mjs for the full contract.
   */
  function inboxInventory(query, folderPath, groupBy, startDate, endDate, unreadOnly, flaggedOnly, tag, includeSubfolders, maxGroups, samplesPerGroup, collectAllIds) {
    const effectiveGroupBy = ["from_domain", "from_address", "subject_prefix", "tag", "day"].includes(groupBy) ? groupBy : "from_domain";
    const maxG = Math.max(1, Math.min(Number.isInteger(maxGroups) && maxGroups > 0 ? maxGroups : 50, 500));
    const samplesN = Math.max(0, Math.min(Number.isInteger(samplesPerGroup) && samplesPerGroup >= 0 ? samplesPerGroup : 3, 20));
    const ID_CAP_PER_GROUP = 5000;

    const matches = _compileMessageFilter({ query, startDate, endDate, unreadOnly, flaggedOnly, tag });
    const groups = new Map();

    function record(msgHdr) {
      if (!matches(msgHdr)) return;
      const key = _groupKeyFor(msgHdr, effectiveGroupBy);
      let g = groups.get(key);
      if (!g) {
        g = { count: 0, sampleIds: [], sampleSubjects: [], firstDate: null, lastDate: null };
        if (collectAllIds) g.messageIds = [];
        groups.set(key, g);
      }
      g.count++;
      if (msgHdr.date) {
        const iso = new Date(msgHdr.date / 1000).toISOString();
        if (!g.firstDate || iso < g.firstDate) g.firstDate = iso;
        if (!g.lastDate || iso > g.lastDate) g.lastDate = iso;
      }
      if (g.sampleIds.length < samplesN) {
        g.sampleIds.push(msgHdr.messageId);
        g.sampleSubjects.push((msgHdr.mime2DecodedSubject || msgHdr.subject || "(no subject)").slice(0, 120));
      }
      if (collectAllIds && g.messageIds.length < ID_CAP_PER_GROUP) {
        g.messageIds.push(msgHdr.messageId);
      }
    }

    let scanStats;
    if (folderPath) {
      const accessible = getAccessibleFolder(folderPath);
      if (accessible.error) return accessible;
      scanStats = _walkFolderMessages(accessible.folder, record, { includeSubfolders: includeSubfolders !== false });
    } else {
      let processed = 0;
      let capped = false;
      for (const account of getAccessibleAccounts()) {
        if (capped) break;
        const sub = _walkFolderMessages(account.incomingServer.rootFolder, record, { includeSubfolders: includeSubfolders !== false, cap: SEARCH_COLLECTION_CAP - processed });
        processed += sub.processed;
        capped = capped || sub.capped;
      }
      scanStats = { processed, capped };
    }

    const sorted = [...groups.entries()]
      .map(([key, g]) => ({ key, ...g }))
      .sort((a, b) => b.count - a.count);
    const truncatedGroups = sorted.length > maxG;
    const totalMatched = sorted.reduce((n, g) => n + g.count, 0);

    return {
      groupedBy: effectiveGroupBy,
      totalMatched,
      groupCount: sorted.length,
      truncatedGroups,
      truncatedScan: !!scanStats.capped,
      messagesScanned: scanStats.processed,
      groups: sorted.slice(0, maxG),
    };
  }

  /**
   * Search + move in one atomic call with dry-run safety. See tools-schema
   * description for contract. Refuses to execute (dry_run=false) without a
   * destination folder.
   */
  function bulkMoveByQuery(query, folderPath, to, dryRun, markRead, flagged, addTags, removeTags, limit, startDate, endDate, unreadOnly, flaggedOnly, includeSubfolders) {
    if (!folderPath) return { error: "folderPath is required" };
    const effectiveDry = (dryRun === undefined || dryRun === null) ? true : !!dryRun;
    const cap = Math.max(1, Math.min(Number.isInteger(limit) && limit > 0 ? limit : 1000, 10000));

    const source = getAccessibleFolder(folderPath);
    if (source.error) return source;

    let dest = null;
    if (!effectiveDry) {
      if (!to) return { error: "to (destination folder URI) is required when dry_run is false" };
      const destResult = getAccessibleFolder(to);
      if (destResult.error) return destResult;
      dest = destResult.folder;
    }

    const matches = _compileMessageFilter({ query, startDate, endDate, unreadOnly, flaggedOnly });
    const matchedHdrs = [];
    const sample = [];

    _walkFolderMessages(source.folder, (msgHdr, folder) => {
      if (matchedHdrs.length >= cap) return false; // stop walking
      if (!matches(msgHdr)) return true;
      matchedHdrs.push({ msgHdr, folder });
      if (sample.length < 5) {
        sample.push({
          id: msgHdr.messageId,
          subject: (msgHdr.mime2DecodedSubject || msgHdr.subject || "").slice(0, 120),
          author: msgHdr.mime2DecodedAuthor || msgHdr.author,
          date: msgHdr.date ? new Date(msgHdr.date / 1000).toISOString() : null,
        });
      }
      return true;
    }, { includeSubfolders: !!includeSubfolders, cap });

    if (effectiveDry) {
      return {
        dryRun: true,
        matched: matchedHdrs.length,
        reachedLimit: matchedHdrs.length >= cap,
        sample,
        wouldMoveTo: to || null,
      };
    }

    // Actually execute. Group by source folder (all the same folder here,
    // since we scoped to folderPath, but _walkFolderMessages may include
    // subfolders, in which case the move-per-source-folder grouping matters).
    const bySource = new Map();
    for (const { msgHdr, folder } of matchedHdrs) {
      if (!bySource.has(folder)) bySource.set(folder, []);
      bySource.get(folder).push(msgHdr);
    }

    let moved = 0;
    const errors = [];
    // tracer comes from the deps bag (.sys.mjs modules don't share
    // globalThis with api.js's experiment-API sandbox).
    const moveTraceId = tracer ? tracer.newTraceId() : null;
    if (tracer) {
      tracer.info("bulk_move", "live.start", {
        trace_id: moveTraceId,
        matched: matchedHdrs.length,
        sources: bySource.size,
        dest_uri: dest && dest.URI,
        addTags, removeTags, markRead, flagged,
      });
    }
    for (const [srcFolder, hdrs] of bySource) {
      try {
        // Apply read/flag before move (IMAP may not preserve on moved copy).
        for (const h of hdrs) {
          try {
            if (markRead === true && !h.isRead) h.markRead(true);
            if (markRead === false && h.isRead) h.markRead(false);
            if (flagged === true && !h.isFlagged) h.markFlagged(true);
            if (flagged === false && h.isFlagged) h.markFlagged(false);
          } catch (e) { errors.push({ id: h.messageId, error: `mark: ${e.toString()}` }); }
          // (Removed an inert ternary loop here that had no side effect:
          //  `for (const t of addTags) h.addProperty ? null : null;` -- the
          //  real tag application is the keyword call below.)
        }
        // Tag add/remove via folder.addKeywordsToMessages / removeKeywordsFromMessages
        try {
          if (Array.isArray(addTags) && addTags.length) {
            srcFolder.addKeywordsToMessages(hdrs, addTags.join(" "));
          }
          if (Array.isArray(removeTags) && removeTags.length) {
            srcFolder.removeKeywordsFromMessages(hdrs, removeTags.join(" "));
          }
        } catch (e) { errors.push({ error: `tag op: ${e.toString()}` }); }

        // CRITICAL: nsIMsgFolder.copyMessages is called on the DESTINATION
        // folder with the SOURCE folder as the first argument. The earlier
        // `srcFolder.copyMessages(dest, hdrs, ...)` had it inverted, which
        // (depending on TB version) either silently no-ops or throws -- in
        // CI run 24916918837 the live path returned moved=0 because the
        // call threw and was caught by the outer try below.
        //
        // Verified against searchfox comm-esr128 nsIMsgFolder.idl. Exact
        // 7-arg signature for TB 128:
        //   void copyMessages(in nsIMsgFolder srcFolder,
        //                     in Array<nsIMsgDBHdr> messages,
        //                     in boolean isMove,
        //                     in nsIMsgWindow msgWindow,
        //                     in nsIMsgCopyServiceListener listener,
        //                     in boolean isFolder,
        //                     in boolean allowUndo);
        // (Note: msgWindow comes BEFORE listener -- different from the
        // pre-128 order that some online docs still show. CI run
        // 24917242345 caught the previous 6-arg attempt with
        // 'Not enough arguments' from XPConnect.)
        // We pass null for both listener and msgWindow (fire-and-forget;
        // completion surfaces via the destination's folder listeners).
        try {
          dest.copyMessages(srcFolder, hdrs, /* isMove */ true,
                            /* msgWindow */ null, /* listener */ null,
                            /* isFolder */ false, /* allowUndo */ false);
          moved += hdrs.length;
          if (tracer) {
            tracer.info("bulk_move", "copy_dispatched", {
              trace_id: moveTraceId,
              src_uri: srcFolder.URI, dst_uri: dest.URI, count: hdrs.length,
            });
          }
        } catch (copyErr) {
          errors.push({ folder: srcFolder.URI, error: `copyMessages: ${copyErr.toString()}` });
          if (tracer) {
            tracer.error("bulk_move", "copy_threw", {
              trace_id: moveTraceId,
              src_uri: srcFolder.URI, dst_uri: dest.URI,
              count: hdrs.length, err: copyErr && copyErr.message,
            });
          }
        }
      } catch (e) {
        errors.push({ folder: srcFolder.URI, error: e.toString() });
        if (tracer) {
          tracer.error("bulk_move", "loop_threw", {
            trace_id: moveTraceId, src_uri: srcFolder.URI, err: e && e.message,
          });
        }
      }
    }
    if (tracer) {
      tracer.info("bulk_move", "live.end", {
        trace_id: moveTraceId, moved, errors: errors.length,
      });
    }

    return {
      dryRun: false,
      matched: matchedHdrs.length,
      moved,
      movedTo: to,
      reachedLimit: matchedHdrs.length >= cap,
      sample,
      errors: errors.length ? errors : undefined,
      note: "Move is async on IMAP. Verify with searchMessages/listFolders after a short delay.",
    };
  }

  return {
    searchMessages,
    getMessage,
    getRecentMessages,
    displayMessage,
    deleteMessages,
    updateMessage,
    findMessage,
    getUserTags,
    markMessageDispositionState,
    inboxInventory,
    bulkMoveByQuery,
    // Shared primitives exposed so registry-driven tool files under
    // lib/tools/ can do their own folder walks without duplicating the
    // searchMessages filter-compilation logic.
    _compileMessageFilter,
    _walkFolderMessages,
    _groupKeyFor,
  };
}
