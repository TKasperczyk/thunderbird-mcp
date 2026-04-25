/**
 * Calendar + task operations for the Thunderbird MCP server.
 *
 * Wraps Thunderbird's calendar XPCOM (cal manager, CalEvent, CalTodo) as a
 * small set of MCP-shaped tool handlers. All public functions return either
 * a result object or { error: string } -- they never throw across the
 * dispatch boundary.
 *
 * NULL-CALENDAR HANDLING (critical):
 * Some Thunderbird builds ship without the calendar component. api.js
 * imports calUtils / CalEvent / CalTodo lazily inside a try/catch and
 * passes whatever it got -- which may be all three nulls -- into this
 * factory. Every public function therefore guards on the specific
 * calendar pieces it needs and returns { error: "Calendar not available" }
 * (or "Calendar module not available" when the item-class is also
 * required) instead of dereferencing a null.
 *
 * Factory: api.js calls makeCalendar({ Services, Cc, Ci, cal, CalEvent,
 * CalTodo }) once per start(). Cc / Ci aren't currently dereferenced but
 * are accepted for parity with sibling factories and to leave room for
 * future XPCOM use without changing the call site.
 */

// VEVENT STATUS values per iCal RFC 5545 § 3.8.1.11.
const VEVENT_STATUS_MAP = {
  tentative: "TENTATIVE",
  confirmed: "CONFIRMED",
  cancelled: "CANCELLED",
  canceled: "CANCELLED",
};

export function makeCalendar({ Services, Cc, Ci, cal, CalEvent, CalTodo }) {
  // --- module-private helpers -------------------------------------------

  /**
   * Fetch events from a calendar over [rangeStart, rangeEnd], using the
   * modern getItemsAsArray API when available and falling back to the
   * stream-based getItems for older Thunderbird builds.
   *
   * Caller must have already confirmed `cal` is non-null.
   */
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

  function normalizeEventStatus(status) {
    if (status === undefined || status === null) return null;
    return VEVENT_STATUS_MAP[String(status).trim().toLowerCase()] || null;
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
      // VEVENT STATUS (tentative/confirmed/cancelled). Empty string
      // when the event has no explicit status (iCal spec treats this
      // as implicit -- Thunderbird renders it like confirmed).
      status: (item.getProperty("STATUS") || "").toLowerCase(),
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

  // --- public surface ---------------------------------------------------

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

  async function createEvent(title, startDate, endDate, location, description, calendarId, allDay, skipReview, status) {
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
      if (status !== undefined && status !== null && status !== "") {
        const normalized = normalizeEventStatus(status);
        if (!normalized) {
          return { error: `Invalid status: "${status}". Expected tentative, confirmed, or cancelled.` };
        }
        event.setProperty("STATUS", normalized);
      }

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

  async function updateEvent(eventId, calendarId, title, startDate, endDate, location, description, status) {
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
      if (status !== undefined) {
        if (status === null || status === "") {
          newItem.deleteProperty("STATUS");
        } else {
          const normalized = normalizeEventStatus(status);
          if (!normalized) {
            return { error: `Invalid status: "${status}". Expected tentative, confirmed, or cancelled.` };
          }
          newItem.setProperty("STATUS", normalized);
        }
        changes.push("status");
      }

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

  return {
    listCalendars,
    createEvent,
    listEvents,
    updateEvent,
    deleteEvent,
    createTask,
    listTasks,
    updateTask,
  };
}
