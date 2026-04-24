/**
 * JSON-RPC dispatch + tool-call validation for the Thunderbird MCP server.
 *
 * Owns the validateToolArgs / coerceToolArgs / callTool trio that backs
 * the MCP tools/call method. The HTTP request handler closure stays
 * inline in api.js because it captures res / server which are HTTP-server
 * specific and not worth threading through a factory.
 *
 * callTool dispatches by tool name to a flat handlers bag the caller
 * assembles from every domain factory (accounts + contacts + folders +
 * calendar + filters + messages + compose + access-control).
 */

export function makeDispatch({
  NetUtil,
  tools,
  toolHandlers,
  permissions,    // optional fine-grained permission engine; if absent only the legacy
                  // isToolEnabled check from access-control gates calls
  isToolEnabled,  // optional Level-1 enable/disable check (still used as the
                  // pre-permissions sandbox floor in callers that haven't yet
                  // adopted the engine; kept for backward-compat)
}) {
  // Pre-build a name -> inputSchema map for fast lookup at validate time.
  const toolSchemas = Object.create(null);
  for (const t of tools) {
    toolSchemas[t.name] = t.inputSchema;
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
        } catch (e) {
          // Leave value as-is so validator surfaces a typed error to the client.
          console.warn(`thunderbird-mcp: coerceToolArgs JSON.parse failed for key=${key}:`, e.message);
        }
      }
    }
    return args;
  }

  async function callTool(name, args) {
    // Permission gate: evaluate before doing any work. May mutate args
    // (e.g. force compose tools off skipReview when policy says alwaysReview).
    if (permissions && typeof permissions.evaluate === "function") {
      // Sandbox floor: pass baseAllowed=true if isToolEnabled accepts the
      // tool, false otherwise. The engine then applies explicit policy
      // overrides on top.
      const baseAllowed = (typeof isToolEnabled === "function") ? !!isToolEnabled(name) : true;
      const decision = permissions.evaluate(name, args, baseAllowed);
      if (!decision.allow) {
        throw new Error(decision.reason || `Tool '${name}' denied by permission policy`);
      }
      args = decision.args;
    }
    switch (name) {
      case "listAccounts":
        return toolHandlers.listAccounts();
      case "listFolders":
        return toolHandlers.listFolders(args.accountId, args.folderPath);
      case "searchMessages":
        return await toolHandlers.searchMessages(args.query || "", args.folderPath, args.startDate, args.endDate, args.maxResults, args.offset, args.sortOrder, args.unreadOnly, args.flaggedOnly, args.tag, args.includeSubfolders, args.countOnly, args.searchBody);
      case "getMessage":
        return await toolHandlers.getMessage(args.messageId, args.folderPath, args.saveAttachments, args.bodyFormat, args.rawSource);
      case "searchContacts":
        return toolHandlers.searchContacts(args.query || "", args.maxResults);
      case "createContact":
        return toolHandlers.createContact(args.email, args.displayName, args.firstName, args.lastName, args.addressBookId);
      case "updateContact":
        return toolHandlers.updateContact(args.contactId, args.email, args.displayName, args.firstName, args.lastName);
      case "deleteContact":
        return toolHandlers.deleteContact(args.contactId);
      case "listCalendars":
        return toolHandlers.listCalendars();
      case "createEvent":
        return await toolHandlers.createEvent(args.title, args.startDate, args.endDate, args.location, args.description, args.calendarId, args.allDay, args.skipReview, args.status);
      case "listEvents":
        return await toolHandlers.listEvents(args.calendarId, args.startDate, args.endDate, args.maxResults);
      case "updateEvent":
        return await toolHandlers.updateEvent(args.eventId, args.calendarId, args.title, args.startDate, args.endDate, args.location, args.description, args.status);
      case "deleteEvent":
        return await toolHandlers.deleteEvent(args.eventId, args.calendarId);
      case "createTask":
        return toolHandlers.createTask(args.title, args.dueDate, args.calendarId);
      case "listTasks":
        return await toolHandlers.listTasks(args.calendarId, args.completed, args.dueBefore, args.maxResults);
      case "updateTask":
        return await toolHandlers.updateTask(args.taskId, args.calendarId, args.title, args.dueDate, args.description, args.completed, args.percentComplete, args.priority);
      case "sendMail":
        return await toolHandlers.composeMail(args.to, args.subject, args.body, args.cc, args.bcc, args.isHtml, args.from, args.attachments, args.skipReview);
      case "replyToMessage":
        return await toolHandlers.replyToMessage(args.messageId, args.folderPath, args.body, args.replyAll, args.isHtml, args.to, args.cc, args.bcc, args.from, args.attachments, args.skipReview);
      case "forwardMessage":
        return await toolHandlers.forwardMessage(args.messageId, args.folderPath, args.to, args.body, args.isHtml, args.cc, args.bcc, args.from, args.attachments, args.skipReview);
      case "getRecentMessages":
        return toolHandlers.getRecentMessages(args.folderPath, args.daysBack, args.maxResults, args.offset, args.unreadOnly, args.flaggedOnly, args.includeSubfolders);
      case "displayMessage":
        return toolHandlers.displayMessage(args.messageId, args.folderPath, args.displayMode);
      case "deleteMessages":
        return toolHandlers.deleteMessages(args.messageIds, args.folderPath);
      case "updateMessage":
        return toolHandlers.updateMessage(args.messageId, args.messageIds, args.folderPath, args.read, args.flagged, args.addTags, args.removeTags, args.moveTo, args.trash);
      case "createFolder":
        return toolHandlers.createFolder(args.parentFolderPath, args.name);
      case "renameFolder":
        return toolHandlers.renameFolder(args.folderPath, args.newName);
      case "deleteFolder":
        return toolHandlers.deleteFolder(args.folderPath);
      case "emptyTrash":
        return toolHandlers.emptyTrash(args.accountId);
      case "emptyJunk":
        return toolHandlers.emptyJunk(args.accountId);
      case "moveFolder":
        return toolHandlers.moveFolder(args.folderPath, args.newParentPath);
      case "listFilters":
        return toolHandlers.listFilters(args.accountId);
      case "createFilter":
        return toolHandlers.createFilter(args.accountId, args.name, args.enabled, args.type, args.conditions, args.actions, args.insertAtIndex);
      case "updateFilter":
        return toolHandlers.updateFilter(args.accountId, args.filterIndex, args.name, args.enabled, args.type, args.conditions, args.actions);
      case "deleteFilter":
        return toolHandlers.deleteFilter(args.accountId, args.filterIndex);
      case "reorderFilters":
        return toolHandlers.reorderFilters(args.accountId, args.fromIndex, args.toIndex);
      case "applyFilters":
        return toolHandlers.applyFilters(args.accountId, args.folderPath);
      case "getAccountAccess":
        return toolHandlers.getAccountAccess();
      default:
        throw new Error(`Unknown tool: ${name}`);
    }
  }

  return { readRequestBody, paginate, validateToolArgs, coerceToolArgs, callTool };
}
