/**
 * inbox_inventory -- aggregate-first query over mail folders.
 *
 * Groups matching messages by from_domain / from_address / subject_prefix
 * / tag / day and returns counts + sample IDs + sample subjects instead
 * of raw per-message headers. Drops a 25k-message inbox from ~80k tokens
 * (searchMessages + N getMessage calls) to <2k tokens (this in one call).
 *
 * First tool migrated to the auto-registry pattern in
 * lib/tool-registry.sys.mjs -- see that module's header for the full
 * contract. Adding a new tool now = drop one file into lib/tools/,
 * rebuild, done.
 */

export const tool = {
  name: "inbox_inventory",
  group: "messages",
  crud: "read",
  title: "Inventory Messages",
  description:
    "Aggregate messages by a key (from_domain, from_address, subject_prefix, tag, day) " +
    "and return grouped counts + sample IDs + sample subjects instead of raw headers. " +
    "Same filter syntax as searchMessages (query, folderPath, date range, unreadOnly, " +
    "flaggedOnly, tag, includeSubfolders). Use this for \"what's in the inbox?\" " +
    "questions on large folders -- a 25k-message inbox becomes ~50 groups and fits in " +
    "<2k tokens instead of 80k+. Set collectAllIds=true to get the full ID lists for " +
    "use with bulk_move_by_query.",
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
  },

  /**
   * Factory receives the shared deps bag and returns a handler that takes
   * the parsed args object. Pure extraction from the legacy implementation
   * in messages.sys.mjs -- same behavior, just reachable without touching
   * tools-schema.sys.mjs / dispatch switch / toolHandlers bag.
   */
  factory({
    MailServices,
    getAccessibleFolder,
    getAccessibleAccounts,
    getUserTags,
    SEARCH_COLLECTION_CAP,
    _compileMessageFilter,
    _walkFolderMessages,
    _groupKeyFor,
  }) {
    return function handler(args) {
      const {
        query = "",
        folderPath, groupBy,
        startDate, endDate, unreadOnly, flaggedOnly, tag,
        includeSubfolders, maxGroups, samplesPerGroup, collectAllIds,
      } = args || {};

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
    };
  },
};
