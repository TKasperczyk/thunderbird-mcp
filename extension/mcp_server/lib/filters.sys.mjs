/**
 * Mail filter (rule) management for the Thunderbird MCP server.
 *
 * Wraps Thunderbird's nsIMsgFilterList / nsIMsgFilter / nsIMsgRuleAction
 * APIs and exposes a small CRUD surface to MCP clients:
 * list / create / update / delete / reorder / apply.
 *
 * The attribute, operator, and action types are XPCOM bitmask/enum
 * constants -- they are kept verbatim from the original api.js extraction
 * (no refactor) so that bit-level compatibility with Thunderbird's
 * filter format is preserved.
 *
 * Factory: api.js calls makeFilters({ MailServices, Cc, Ci,
 * isAccountAllowed, getAccessibleFolder, getAccessibleAccounts }).
 *
 * getAccessibleAccounts comes from the access-control module and is
 * needed by listFilters when the caller doesn't pin a specific account.
 *
 * Module-private helpers (not exported):
 *   - getFilterListForAccount  : access-checked filter-list lookup
 *   - serializeFilter          : nsIMsgFilter -> plain JSON shape
 *   - buildTerms / buildActions: reverse direction, MCP shape -> XPCOM
 */

export function makeFilters({
  MailServices,
  Cc,
  Ci,
  isAccountAllowed,
  getAccessibleFolder,
  getAccessibleAccounts,
}) {
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

  /**
   * Run filters (or simulate) on a folder and return per-filter hit counts.
   *
   * Always runs a dry-evaluation pass first using nsIMsgFilter.matchHdr
   * to compute how many messages each enabled filter would match. If
   * dry_run is true, stops there. Otherwise calls the real filterService
   * to actually execute the actions and returns the predicted counts
   * (which track the real run assuming no mid-flight mutations).
   */
  function applyFilters(accountId, folderPath, dryRun) {
    try {
      const fl = getFilterListForAccount(accountId);
      if (fl.error) return fl;
      const { filterList } = fl;

      const afResult = getAccessibleFolder(folderPath);
      if (afResult.error) return afResult;
      const folder = afResult.folder;

      // Collect enabled filters (we only evaluate these; disabled filters
      // are skipped to match what filterService.applyFiltersToFolders does).
      const enabledFilters = [];
      for (let i = 0; i < filterList.filterCount; i++) {
        const f = filterList.getFilterAt(i);
        if (f.enabled) enabledFilters.push({ index: i, filter: f, hit: 0 });
      }

      // Per-filter dry evaluation using nsIMsgFilter.matchHdr (TB's own
      // evaluator). Single pass over folder messages for all filters.
      let messagesProcessed = 0;
      let evaluationError = null;
      try {
        const db = folder.msgDatabase;
        if (db) {
          for (const msgHdr of db.enumerateMessages()) {
            messagesProcessed++;
            for (const entry of enabledFilters) {
              try {
                if (entry.filter.matchHdr(msgHdr, folder, db, null, 0)) {
                  entry.hit++;
                }
              } catch { /* filter.matchHdr may throw on exotic conditions; skip */ }
            }
          }
        }
      } catch (e) {
        evaluationError = e.toString();
      }

      const byFilter = enabledFilters.map(e => ({
        name: e.filter.filterName,
        index: e.index,
        hit: e.hit,
        actions: (() => {
          const acts = [];
          const n = e.filter.actionCount || 0;
          for (let i = 0; i < n; i++) {
            try {
              const a = e.filter.getActionAt(i);
              const actionType = (ACTION_NAMES[a.type] || `type${a.type}`);
              const entry = { type: actionType };
              try { if (a.strValue) entry.value = a.strValue; } catch { /* ignore */ }
              try { if (a.targetFolderUri) entry.targetFolderUri = a.targetFolderUri; } catch { /* ignore */ }
              try { if (typeof a.priority === "number") entry.priority = a.priority; } catch { /* ignore */ }
              acts.push(entry);
            } catch { /* skip unreadable actions */ }
          }
          return acts;
        })(),
      }));

      const baseReport = {
        folder: folderPath,
        filtersRun: enabledFilters.length,
        messagesProcessed,
        byFilter,
        totalHits: byFilter.reduce((n, e) => n + e.hit, 0),
      };
      if (evaluationError) baseReport.evaluationWarning = evaluationError;

      if (dryRun === true) {
        return { ...baseReport, dryRun: true };
      }

      // Real execution -- hand off to Thunderbird's filter service after
      // computing the dry report so the caller still gets per-filter counts.
      let filterService;
      try { filterService = MailServices.filters; } catch { /* fall through */ }
      if (!filterService) {
        try {
          filterService = Cc["@mozilla.org/messenger/filter-service;1"]
            .getService(Ci.nsIMsgFilterService);
        } catch { /* ignore */ }
      }
      if (!filterService) {
        return { error: "Filter service not available in this Thunderbird version" };
      }
      filterService.applyFiltersToFolders(filterList, [folder], null);
      return {
        ...baseReport,
        dryRun: false,
        note: "applyFiltersToFolders is async; counts are predicted hit rates, verify state with searchMessages after a short delay",
      };
    } catch (e) {
      return { error: e.toString() };
    }
  }

  return {
    listFilters,
    createFilter,
    updateFilter,
    deleteFilter,
    reorderFilters,
    applyFilters,
  };
}
