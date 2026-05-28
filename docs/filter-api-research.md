# Thunderbird Mail Filters — API Research & Implementation Guide

Research completed 2026-02-21. Everything needed to implement filter control via MCP tools.

## TL;DR

Thunderbird exposes full filter CRUD + execution via XPCOM interfaces. Our extension already uses an Experiment API with full XPCOM access — no new permissions needed. We can list, create, modify, delete, reorder, and manually apply filters.

---

## 1. Current Extension Architecture

```
MCP Client <--stdio--> mcp-bridge.cjs <--HTTP POST--> Thunderbird Extension (port 8765)
```

- **mcp-bridge.cjs**: Node.js process, translates stdio JSON-RPC ↔ HTTP
- **Extension**: Embedded HTTP server (`httpd.sys.mjs`) on `localhost:8765`
- **Protocol**: JSON-RPC 2.0 with methods `tools/list` and `tools/call`
- **API file**: `extension/mcp_server/api.js` (~1920 lines, monolithic — all tool defs + handlers)

### Experiment API Setup

The extension uses Thunderbird's Experiment API for full XPCOM access:

**manifest.json** (relevant section):
```json
{
  "experiment_apis": {
    "mcpServer": {
      "schema": "mcp_server/schema.json",
      "parent": {
        "scopes": ["addon_parent"],
        "paths": [["mcpServer"]],
        "script": "mcp_server/api.js"
      }
    }
  },
  "permissions": [
    "accountsRead", "addressBooks", "messagesRead",
    "messagesMove", "accountsFolders", "compose"
  ]
}
```

**No additional manifest permissions needed for filters** — filter access is via XPCOM, not WebExtension APIs.

### How XPCOM Is Accessed (existing patterns in api.js)

```js
// Already imported at line 259:
const { MailServices } = ChromeUtils.importESModule(
  "resource:///modules/MailServices.sys.mjs"
);

// Account iteration pattern (line 349):
for (const account of MailServices.accounts.accounts) {
  const server = account.incomingServer;
  // server.getFilterList(null) ← THIS is how we get filters
}

// XPCOM service instantiation pattern (line 17):
const resProto = Cc[
  "@mozilla.org/network/protocol;1?name=resource"
].getService(Ci.nsISubstitutingProtocolHandler);

// Globals available: ExtensionCommon, ChromeUtils, Services, Cc, Ci
```

### How Tools Are Registered

In `api.js`, there are three places to touch:

1. **Tool definition** — the `tools` array (starts at line 37), each entry has `name`, `title`, `description`, `inputSchema`
2. **Handler function** — standalone `function` or `async function` defined in the same scope
3. **Dispatch** — the `callTool()` switch statement (starts at line 1821)

Example (createFolder, the simplest existing tool):

```js
// 1. Tool definition (in the tools array):
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

// 2. Handler function:
function createFolder(parentFolderPath, name) {
  try {
    const parentFolder = MailServices.folderLookup.getFolderForURL(parentFolderPath);
    if (!parentFolder) return { error: "Parent folder not found" };
    parentFolder.createSubfolder(name, null);
    const newFolder = parentFolder.getChildNamed(name);
    const newPath = newFolder.URI;
    return { created: true, name, path: newPath };
  } catch (e) {
    const msg = e.toString();
    if (msg.includes("NS_MSG_FOLDER_EXISTS")) {
      return { error: `Folder "${name}" already exists under this parent` };
    }
    return { error: msg };
  }
}

// 3. Dispatch (in callTool switch):
case "createFolder":
  return createFolder(args.parentFolderPath, args.name);
```

---

## 2. Thunderbird Filter XPCOM Interfaces

### nsIMsgFilterService — Central Service

Access:
```js
const filterService = Cc["@mozilla.org/messenger/filter-service;1"]
  .getService(Ci.nsIMsgFilterService);
```

Or possibly via `MailServices.filters` if available in the ESModule import.

Key methods:
- `OpenFilterList(filterFile)` → `nsIMsgFilterList` — Open filter list from file
- `SaveFilterList(filterList)` — Save to file
- `getTempFilterList(folder)` → `nsIMsgFilterList` — Temporary filter list for testing
- `applyFiltersToFolders(filterList, folders[], msgWindow)` — **Run filters on folders**
- `applyFilters(filterType, msgHdrArray, folder, msgWindow, filterList)` — Run on specific messages
- `addCustomAction(action)` / `getCustomActions()` / `getCustomAction(id)` — Custom action registration
- `addCustomTerm(term)` / `getCustomTerms()` / `getCustomTerm(id)` — Custom search term registration
- `filterTypeName(filterType)` → readable name from type flags

### nsIMsgFilterList — Per-Server Filter Collection

Access via server:
```js
const server = account.incomingServer;
const filterList = server.getFilterList(null);  // null = no msgWindow
// or: server.getEditableFilterList(null);
```

Key properties:
- `filterCount` — Number of filters
- `folder` — Associated folder
- `loggingEnabled` — Whether filter logging is on
- `version` — Filter file format version
- `defaultFile` — Path to `msgFilterRules.dat`
- `listId` — Identifier string

Key methods:
- `getFilterAt(index)` → `nsIMsgFilter`
- `getFilterNamed(name)` → `nsIMsgFilter`
- `createFilter(name)` → `nsIMsgFilter` — Creates a new empty filter (NOT yet in the list)
- `insertFilterAt(index, filter)` — Insert into list at position
- `removeFilter(filter)` — Remove from list
- `removeFilterAt(index)` — Remove by index
- `moveFilterAt(sourceIndex, destIndex)` — Reorder
- `moveFilter(filter, motion)` — Move up/down (motion is a constant)
- `saveToFile(file)` / `saveToDefaultFile()` — **Persist changes**
- `applyFiltersToHdr(filterType, msgHdr, folder, msgDatabase, headers, listener, msgWindow)` — Apply to single message
- `parseCondition(filter, condition)` — Parse a condition string like `"AND (from,contains,foo)"`
- `clearLog()` / `flushLogIfNecessary()` — Log management
- `logURL` / `logStream` — Access log data

### nsIMsgFilter — Individual Filter

Key properties:
- `filterName` — Display name (string)
- `filterDesc` — Description (string)
- `enabled` — Boolean
- `temporary` — Boolean (temp filters don't persist)
- `filterType` — Bitmask (see Filter Types below)
- `filterList` — Owning `nsIMsgFilterList`
- `unparseable` — Boolean (broken filter)
- `searchTerms` — Array of `nsIMsgSearchTerm`
- `scope` — `nsIMsgSearchScopeTerm`
- `actionCount` — Number of actions
- `sortedActionList` — Actions array

Key methods:
- `createTerm()` → `nsIMsgSearchTerm` — Create new search term
- `appendTerm(term)` — Add search term to filter
- `createAction()` → `nsIMsgRuleAction` — Create new action
- `appendAction(action)` — Add action to filter
- `getActionAt(index)` → `nsIMsgRuleAction`
- `clearActionList()` — Remove all actions
- `MatchHdr(msgHdr, folder, db, headers, headerSize)` → boolean — Test if message matches
- `SaveToTextFile(stream)` — Serialize to file format

### nsIMsgSearchTerm — Filter Criteria

Key properties:
- `attrib` — Search attribute (see Search Attributes below)
- `op` — Search operator (see Search Operators below)
- `value` — `nsIMsgSearchValue` (has `.str` for strings, `.date` for dates, etc.)
- `booleanAnd` — `true` for AND, `false` for OR
- `arbitraryHeader` — Custom header name when `attrib` is arbitrary header
- `hdrProperty` — Header property name
- `customId` — ID of custom search term
- `beginsGrouping` / `endsGrouping` — Parenthetical grouping

Key methods:
- `matchRfc822String(str)`, `matchRfc2047String(str)` — Match against header strings
- `matchDate(date)`, `matchStatus(status)`, `matchPriority(priority)` — Typed matching
- `matchAge(date, now)`, `matchSize(size)` — Relative matching
- `matchBody(folderScopeTerm, offset, length, charset, msgHdr, db)` — Body search
- `matchArbitraryHeader(headers)` — Custom header matching
- `matchKeyword(keyword)` — Tag/keyword matching
- `termAsString` — Read-only serialization (e.g., `"(from,contains,newsletter@)"`)

### nsIMsgRuleAction — Filter Action

Key properties:
- `type` — Action type constant (see Filter Actions below)
- `priority` — Priority value when action is "set priority"
- `targetFolderUri` — Destination folder when action is move/copy
- `junkScore` — Junk score value
- `customId` — ID of custom action
- `customAction` — `nsIMsgFilterCustomAction` reference

---

## 3. Constants & Enumerations

### Filter Types (bitmask, combinable)

```
nsMsgFilterType.InboxRule          = 0x1     // Applied on new mail
nsMsgFilterType.InboxJavaScript    = 0x2     // JS filter on new mail
nsMsgFilterType.Inbox              = 0x3     // Combined inbox types
nsMsgFilterType.NewsRule           = 0x4     // News filter
nsMsgFilterType.NewsJavaScript     = 0x8     // JS news filter
nsMsgFilterType.News               = 0xC     // Combined news types
nsMsgFilterType.Incoming           = 0xF     // All incoming types
nsMsgFilterType.Manual             = 0x10    // Manually applied
nsMsgFilterType.PostPlugin         = 0x20    // After junk/bayesian
nsMsgFilterType.PostOutgoing       = 0x40    // After sending
nsMsgFilterType.Archive            = 0x80    // On archive action
nsMsgFilterType.Periodic           = 0x100   // Periodic execution
```

Common combination: type `17` = InboxRule (0x1) + Manual (0x10)

### Search Attributes (nsMsgSearchAttrib)

```
Subject        = 0     // Message subject
Sender         = 1     // From header
Body           = 2     // Message body
Date           = 3     // Date header
Priority       = 4     // Priority
MsgStatus      = 5     // Read/replied/forwarded status
To             = 6     // To header
CC             = 7     // CC header
ToOrCC         = 8     // To or CC
AllAddresses   = 9     // From, To, CC, BCC
AgeInDays      = 10    // Message age
Size           = 11    // Message size
Keywords       = 12    // Tags/keywords
HasAttachment  = 13    // Has attachments
JunkStatus     = 14    // Junk classification
JunkPercent    = 15    // Junk probability
OtherHeader    = 16    // Arbitrary header (uses arbitraryHeader property)
```

### Search Operators (nsMsgSearchOp)

```
Contains       = 0
DoesntContain  = 1
Is             = 2
Isnt           = 3
IsEmpty        = 4
IsBefore       = 5     // Date comparison
IsAfter        = 6     // Date comparison
IsHigherThan   = 7     // Priority comparison
IsLowerThan    = 8     // Priority comparison
BeginsWith     = 9
EndsWith       = 10
IsInAB         = 11    // Is in address book
IsntInAB       = 12    // Not in address book
IsGreaterThan  = 13    // Size comparison
IsLessThan     = 14    // Size comparison
Matches        = 15    // Regex match
DoesntMatch    = 16    // Regex no match
```

### Filter Actions (nsMsgFilterAction)

```
MoveToFolder       = 0x01
CopyToFolder       = 0x02
ChangePriority     = 0x03
Delete             = 0x04
MarkRead           = 0x05
KillThread         = 0x06
WatchThread        = 0x07
MarkFlagged        = 0x08
Label              = 0x09    // Deprecated, use AddTag
Reply              = 0x0A
Forward            = 0x0B
StopExecution      = 0x0C    // Stop processing more filters
DeleteFromServer   = 0x0D    // POP3: don't download
LeaveOnServer      = 0x0E    // POP3: leave on server
JunkScore          = 0x0F
FetchBodyFromServer = 0x10   // IMAP: fetch full body
AddTag             = 0x11    // Add a tag/keyword
DeleteBody         = 0x12
MarkUnread         = 0x14
Custom             = 0x15    // Custom action via nsIMsgFilterCustomAction
```

---

## 4. Filter File Format (msgFilterRules.dat)

Filters are persisted in plain text. Each server has its own file, typically at:
`<profile>/ImapMail/<server>/msgFilterRules.dat` or `<profile>/Mail/<server>/msgFilterRules.dat`

Format:
```
version="9"
logging="no"
name="Sort newsletters"
enabled="yes"
type="17"
action="Move to folder"
actionValue="imap://user@imap.example.com/Newsletters"
condition="AND (from,contains,newsletter@) OR (subject,contains,[newsletter])"
name="Flag important"
enabled="yes"
type="1"
action="Mark flagged"
condition="AND (from,is,boss@company.com)"
```

---

## 5. Proposed MCP Tools

### 5.1 listFilters

**Purpose**: List all filters for an account (or all accounts).

```js
{
  name: "listFilters",
  title: "List Filters",
  description: "List all mail filters/rules for an account with their conditions and actions",
  inputSchema: {
    type: "object",
    properties: {
      accountId: {
        type: "string",
        description: "Account ID from listAccounts (omit for all accounts)"
      }
    },
    required: []
  }
}
```

**Implementation sketch**:
```js
function listFilters(accountId) {
  const results = [];
  const accounts = accountId
    ? [MailServices.accounts.getAccount(accountId)]
    : Array.from(MailServices.accounts.accounts);

  for (const account of accounts) {
    const server = account.incomingServer;
    if (!server.canHaveFilters) continue;

    const filterList = server.getFilterList(null);
    const filters = [];

    for (let i = 0; i < filterList.filterCount; i++) {
      const filter = filterList.getFilterAt(i);
      const terms = [];
      const actions = [];

      // Extract search terms
      for (const term of filter.searchTerms) {
        terms.push({
          attrib: term.attrib,        // numeric — map to name for readability
          op: term.op,                // numeric — map to name
          value: term.value.str || term.value.date || String(term.value),
          booleanAnd: term.booleanAnd,
          arbitraryHeader: term.arbitraryHeader || undefined,
        });
      }

      // Extract actions
      for (let a = 0; a < filter.actionCount; a++) {
        const action = filter.getActionAt(a);
        actions.push({
          type: action.type,          // numeric — map to name
          targetFolderUri: action.targetFolderUri || undefined,
          priority: action.priority || undefined,
          customId: action.customId || undefined,
        });
      }

      filters.push({
        index: i,
        name: filter.filterName,
        enabled: filter.enabled,
        type: filter.filterType,
        temporary: filter.temporary,
        terms,
        actions,
      });
    }

    results.push({
      accountId: account.key,
      accountName: server.prettyName,
      filterCount: filterList.filterCount,
      loggingEnabled: filterList.loggingEnabled,
      filters,
    });
  }

  return results;
}
```

**Important**: Map numeric attrib/op/action constants to human-readable names in output. Create lookup objects like:
```js
const ATTRIB_NAMES = { 0: "subject", 1: "from", 2: "body", 3: "date", ... };
const OP_NAMES = { 0: "contains", 1: "doesntContain", 2: "is", 3: "isnt", ... };
const ACTION_NAMES = { 1: "moveToFolder", 2: "copyToFolder", 5: "markRead", 8: "markFlagged", 0x11: "addTag", ... };
```

### 5.2 createFilter

**Purpose**: Create a new filter with conditions and actions.

```js
{
  name: "createFilter",
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
            op: { type: "string", description: "Operator: contains, doesntContain, is, isnt, isEmpty, beginsWith, endsWith, isGreaterThan, isLessThan, isBefore, isAfter" },
            value: { type: "string", description: "Value to match against" },
            booleanAnd: { type: "boolean", description: "true=AND with previous, false=OR (default: true)" },
            header: { type: "string", description: "Custom header name (only when attrib is otherHeader)" },
          },
          required: ["attrib", "op", "value"]
        },
        description: "Array of filter conditions"
      },
      actions: {
        type: "array",
        items: {
          type: "object",
          properties: {
            type: { type: "string", description: "Action: moveToFolder, copyToFolder, markRead, markUnread, markFlagged, addTag, changePriority, delete, stopExecution, forward, reply" },
            value: { type: "string", description: "Action parameter (folder URI for move/copy, tag name for addTag, priority for changePriority, email for forward)" },
          },
          required: ["type"]
        },
        description: "Array of actions to perform"
      },
      insertAtIndex: { type: "number", description: "Position to insert (0 = top priority, default: end of list)" },
    },
    required: ["accountId", "name", "conditions", "actions"]
  }
}
```

**Implementation sketch**:
```js
function createFilter(accountId, name, enabled, type, conditions, actions, insertAtIndex) {
  const account = MailServices.accounts.getAccount(accountId);
  if (!account) return { error: "Account not found" };
  const server = account.incomingServer;
  if (!server.canHaveFilters) return { error: "Account does not support filters" };

  const filterList = server.getFilterList(null);
  const filter = filterList.createFilter(name);

  filter.enabled = enabled !== false;
  filter.filterType = type || 17; // inbox + manual

  // Map string attribute names → numeric constants
  const ATTRIB_MAP = {
    subject: 0, from: 1, body: 2, date: 3, priority: 4,
    status: 5, to: 6, cc: 7, toOrCc: 8, allAddresses: 9,
    ageInDays: 10, size: 11, tag: 12, hasAttachment: 13,
    junkStatus: 14, junkPercent: 15, otherHeader: 16,
  };
  const OP_MAP = {
    contains: 0, doesntContain: 1, is: 2, isnt: 3, isEmpty: 4,
    isBefore: 5, isAfter: 6, isHigherThan: 7, isLowerThan: 8,
    beginsWith: 9, endsWith: 10,
    soundsLike: 11, ldapDwim: 12,
    isGreaterThan: 13, isLessThan: 14,
    nameCompletion: 15, isInAB: 16, isntInAB: 17, isntEmpty: 18,
    matches: 19, doesntMatch: 20,
  };

  for (const cond of conditions) {
    const term = filter.createTerm();
    term.attrib = ATTRIB_MAP[cond.attrib] ?? parseInt(cond.attrib);
    term.op = OP_MAP[cond.op] ?? parseInt(cond.op);
    // Setting value depends on type — for string attributes:
    const value = term.value;
    value.attrib = term.attrib;
    value.str = cond.value;
    term.value = value;
    term.booleanAnd = cond.booleanAnd !== false;
    if (cond.header) term.arbitraryHeader = cond.header;
    filter.appendTerm(term);
  }

  const ACTION_MAP = {
    moveToFolder: 0x01, copyToFolder: 0x02, changePriority: 0x03,
    delete: 0x04, markRead: 0x05, killThread: 0x06,
    watchThread: 0x07, markFlagged: 0x08, reply: 0x0A,
    forward: 0x0B, stopExecution: 0x0C, deleteFromServer: 0x0D,
    leaveOnServer: 0x0E, junkScore: 0x0F, addTag: 0x11,
    markUnread: 0x14, custom: 0x15,
  };

  for (const act of actions) {
    const action = filter.createAction();
    action.type = ACTION_MAP[act.type] ?? parseInt(act.type);
    if (act.value) {
      if (action.type === 0x01 || action.type === 0x02) {
        action.targetFolderUri = act.value;
      } else if (action.type === 0x03) {
        action.priority = parseInt(act.value);
      } else {
        // For addTag, forward, etc. — check what property to set
        // addTag uses strValue, forward/reply use strValue as email address
        action.strValue = act.value;
      }
    }
    filter.appendAction(action);
  }

  // Insert at position or append
  const idx = (insertAtIndex != null && insertAtIndex >= 0)
    ? Math.min(insertAtIndex, filterList.filterCount)
    : filterList.filterCount;
  filterList.insertFilterAt(idx, filter);
  filterList.saveToDefaultFile();

  return {
    created: true,
    name: filter.filterName,
    index: idx,
    filterCount: filterList.filterCount,
  };
}
```

### 5.3 updateFilter

**Purpose**: Modify an existing filter (toggle enabled, rename, change conditions/actions).

```js
{
  name: "updateFilter",
  title: "Update Filter",
  description: "Modify an existing filter's properties, conditions, or actions",
  inputSchema: {
    type: "object",
    properties: {
      accountId: { type: "string", description: "Account ID" },
      filterIndex: { type: "number", description: "Filter index (from listFilters)" },
      name: { type: "string", description: "New filter name (optional)" },
      enabled: { type: "boolean", description: "Enable/disable (optional)" },
      conditions: { type: "array", description: "Replace conditions (optional, same format as createFilter)" },
      actions: { type: "array", description: "Replace actions (optional, same format as createFilter)" },
    },
    required: ["accountId", "filterIndex"]
  }
}
```

### 5.4 deleteFilter

**Purpose**: Remove a filter.

```js
{
  name: "deleteFilter",
  title: "Delete Filter",
  description: "Delete a mail filter by index",
  inputSchema: {
    type: "object",
    properties: {
      accountId: { type: "string", description: "Account ID" },
      filterIndex: { type: "number", description: "Filter index to delete (from listFilters)" },
    },
    required: ["accountId", "filterIndex"]
  }
}
```

**Implementation sketch**:
```js
function deleteFilter(accountId, filterIndex) {
  const account = MailServices.accounts.getAccount(accountId);
  const filterList = account.incomingServer.getFilterList(null);
  if (filterIndex < 0 || filterIndex >= filterList.filterCount) {
    return { error: `Invalid filter index ${filterIndex}` };
  }
  const filter = filterList.getFilterAt(filterIndex);
  const name = filter.filterName;
  filterList.removeFilterAt(filterIndex);
  filterList.saveToDefaultFile();
  return { deleted: true, name, remainingCount: filterList.filterCount };
}
```

### 5.5 reorderFilters

**Purpose**: Change filter execution priority.

```js
{
  name: "reorderFilters",
  title: "Reorder Filters",
  description: "Move a filter to a different position in the execution order",
  inputSchema: {
    type: "object",
    properties: {
      accountId: { type: "string", description: "Account ID" },
      fromIndex: { type: "number", description: "Current filter index" },
      toIndex: { type: "number", description: "Target index (0 = highest priority)" },
    },
    required: ["accountId", "fromIndex", "toIndex"]
  }
}
```

**Implementation**:
```js
function reorderFilters(accountId, fromIndex, toIndex) {
  const account = MailServices.accounts.getAccount(accountId);
  const filterList = account.incomingServer.getFilterList(null);
  filterList.moveFilterAt(fromIndex, toIndex);
  filterList.saveToDefaultFile();
  return { moved: true, fromIndex, toIndex };
}
```

### 5.6 applyFilters

**Purpose**: Manually run filters on a folder. This is the killer feature — an AI agent can organize mail on demand.

```js
{
  name: "applyFilters",
  title: "Apply Filters",
  description: "Manually run all enabled filters on a folder to organize existing messages",
  inputSchema: {
    type: "object",
    properties: {
      accountId: { type: "string", description: "Account ID (uses its filters)" },
      folderPath: { type: "string", description: "Folder URI to apply filters to (from listFolders)" },
    },
    required: ["accountId", "folderPath"]
  }
}
```

**Implementation sketch**:
```js
function applyFilters(accountId, folderPath) {
  const account = MailServices.accounts.getAccount(accountId);
  const server = account.incomingServer;
  const filterList = server.getFilterList(null);
  const folder = MailServices.folderLookup.getFolderForURL(folderPath);
  if (!folder) return { error: "Folder not found" };

  // Method signature: applyFiltersToFolders(filterList, folders, msgWindow)
  const filterService = Cc["@mozilla.org/messenger/filter-service;1"]
    .getService(Ci.nsIMsgFilterService);
  filterService.applyFiltersToFolders(filterList, [folder], null);

  return { applied: true, folder: folderPath, filterCount: filterList.filterCount };
}
```

**Note**: `applyFiltersToFolders` may be asynchronous in practice — the function returns immediately but processing continues. We may need to investigate whether there's a completion callback or listener to await.

---

## 6. Gotchas & Edge Cases

### Persistence
- **Always call `filterList.saveToDefaultFile()`** after mutations (create, update, delete, reorder). Without this, changes exist only in memory and are lost on restart.
- Mutate the filter list in-place. Don't try `server.setFilterList(newList)` — that doesn't persist properly.

### Filter List Access
- `server.getFilterList(null)` — pass `null` for msgWindow unless you need UI updates
- `server.getEditableFilterList(null)` — may differ from `getFilterList` in some contexts, but usually the same for local/IMAP accounts
- `server.canHaveFilters` — check this before attempting filter operations (news servers may not support filters)

### Search Term Value Setting
- The `value` property on `nsIMsgSearchTerm` is an `nsIMsgSearchValue` object
- You must set `value.attrib` to match the term's attrib before setting the value content
- For string attributes: `value.str = "something"`
- For date attributes: `value.date` (PRTime, microseconds since epoch)
- For priority: `value.priority` (numeric constant)
- For status: `value.status` (bitmask)
- For junk: `value.junkStatus` / `value.junkPercent`

### Action Value Setting
- `MoveToFolder` / `CopyToFolder`: set `action.targetFolderUri`
- `ChangePriority`: set `action.priority`
- `AddTag`: tag value goes in `action.strValue` (the keyword, e.g., `"$label1"` or custom tag keyword)
- `Forward` / `Reply`: email address in `action.strValue`
- `JunkScore`: set `action.junkScore`
- Actions like `MarkRead`, `MarkFlagged`, `StopExecution`, `Delete` have no value parameter

### Async Considerations
- `applyFiltersToFolders` returns immediately — the actual filtering happens asynchronously
- For move/copy actions, the messages may not be relocated instantly
- Consider returning a "filters applied, processing may take a moment" status rather than waiting

### nsIMsgSearchTerm Iteration
- `filter.searchTerms` should be iterable, but depending on Thunderbird version it may return an `nsIMutableArray` requiring `.enumerate()` or similar
- Test with: `for (const term of filter.searchTerms) { ... }` — if that doesn't work, try `filter.searchTerms.enumerate(Ci.nsIMsgSearchTerm)` or use indexed access

### Error Handling
- Wrap all XPCOM calls in try/catch — XPCOM throws NS_ERROR exceptions as JS errors
- Common: `NS_ERROR_UNEXPECTED` if filter list file is locked or corrupt
- Account lookup: `MailServices.accounts.getAccount(id)` may return null for invalid IDs

---

## 7. Testing Strategy

### Manual Verification
1. Create a filter via the MCP tool
2. Open Thunderbird → Account Settings → Message Filters → verify it appears
3. Send a test email matching the filter → verify it triggers
4. Modify the filter via MCP → verify changes in Thunderbird UI
5. Delete the filter → verify removal

### Edge Cases to Test
- Filter with multiple conditions (AND + OR)
- Filter with multiple actions
- Filter on custom/arbitrary headers
- Reordering filters
- Applying filters to IMAP folder (async concerns)
- Applying filters to local folder
- Account with `canHaveFilters = false`
- Filter with special characters in name/values (Unicode, quotes)

---

## 8. Alternative: Condition String Parsing

Instead of structured conditions, we could accept raw condition strings:
```
"AND (from,contains,newsletter@) OR (subject,contains,[news])"
```

The `filterList.parseCondition(filter, conditionString)` method can parse these directly. This would simplify the API for power users but make it harder for AI agents to construct programmatically. **Recommendation**: use the structured format as primary, but consider adding a `rawCondition` escape hatch.

---

## 9. Contract ID Discovery

If `MailServices.filters` isn't available in the ESModule import, the XPCOM contract ID is:
```
@mozilla.org/messenger/filter-service;1
```

Instantiate with:
```js
const filterService = Cc["@mozilla.org/messenger/filter-service;1"]
  .getService(Ci.nsIMsgFilterService);
```

Alternatively, check if `MailServices` exposes it:
```js
// In api.js, MailServices is already imported at line 259:
const { MailServices } = ChromeUtils.importESModule("resource:///modules/MailServices.sys.mjs");
// Check: MailServices.filters — may or may not exist depending on TB version
```

If `MailServices.filters` exists, prefer that over manual `Cc` lookup for consistency with existing code.

---

## 10. Implementation Checklist

- [ ] Add constant lookup maps (ATTRIB_NAMES, OP_NAMES, ACTION_NAMES and reverse maps)
- [ ] Implement `listFilters(accountId)` handler
- [ ] Implement `createFilter(...)` handler
- [ ] Implement `updateFilter(...)` handler
- [ ] Implement `deleteFilter(accountId, filterIndex)` handler
- [ ] Implement `reorderFilters(accountId, fromIndex, toIndex)` handler
- [ ] Implement `applyFilters(accountId, folderPath)` handler
- [ ] Add all 6 tool definitions to the `tools` array
- [ ] Add all 6 dispatch cases to `callTool()` switch
- [ ] Test each tool with real Thunderbird instance
- [ ] Handle edge cases (canHaveFilters, null accounts, empty filter lists)
- [ ] Ensure `saveToDefaultFile()` called after all mutations
- [ ] Verify `searchTerms` iteration works on target Thunderbird version
