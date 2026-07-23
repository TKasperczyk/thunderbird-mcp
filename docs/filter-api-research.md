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

Source: [`nsMsgFilterCore.idl` L13-L24](https://github.com/mozilla/releases-comm-central/blob/72b8ba0761b3881d926be53f77fbf75d5a9316d5/mailnews/search/public/nsMsgFilterCore.idl#L13-L24) (`nsMsgFilterType`) — verified against upstream 2026-07-23.

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

Source: [`nsMsgSearchCore.idl` L49-L104](https://github.com/mozilla/releases-comm-central/blob/72b8ba0761b3881d926be53f77fbf75d5a9316d5/mailnews/search/public/nsMsgSearchCore.idl#L49-L104) (`nsMsgSearchAttrib`) — verified against upstream 2026-07-23.

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
AgeInDays      = 12    // Message age
Size           = 14    // Message size
Keywords       = 16    // Tags/keywords
HasAttachment  = 44    // Has attachments (HasAttachmentStatus)
JunkStatus     = 45    // Junk classification
JunkPercent    = 46    // Junk probability
OtherHeader    = 52    // Arbitrary header (uses arbitraryHeader property)
```

### Search Operators (nsMsgSearchOp)

Source: [`nsMsgSearchCore.idl` L113-L139](https://github.com/mozilla/releases-comm-central/blob/72b8ba0761b3881d926be53f77fbf75d5a9316d5/mailnews/search/public/nsMsgSearchCore.idl#L113-L139) (`nsMsgSearchOp`) — verified against upstream 2026-07-23.

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
SoundsLike     = 11    // Phonetic match (LDAP)
LdapDwim       = 12    // LDAP do-what-I-mean
IsGreaterThan  = 13    // Size comparison
IsLessThan     = 14    // Size comparison
NameCompletion = 15    // LDAP name completion
IsInAB         = 16    // Is in address book
IsntInAB       = 17    // Not in address book
IsntEmpty      = 18
Matches        = 19    // Regex match
DoesntMatch    = 20    // Regex no match
```

### Filter Actions (nsMsgFilterAction)

Source: [`nsMsgFilterCore.idl` L38-L60](https://github.com/mozilla/releases-comm-central/blob/72b8ba0761b3881d926be53f77fbf75d5a9316d5/mailnews/search/public/nsMsgFilterCore.idl#L38-L60) (`nsMsgRuleActionType`) — verified against upstream 2026-07-23.

```
Custom             = -1      // Custom action via nsIMsgFilterCustomAction
MoveToFolder       = 1
ChangePriority     = 2
Delete             = 3
MarkRead           = 4
KillThread         = 5
WatchThread        = 6
MarkFlagged        = 7
// 8 (Label) removed -- use AddTag
Reply              = 9
Forward            = 10
StopExecution      = 11      // Stop processing more filters
DeleteFromServer   = 12      // POP3: don't download (DeleteFromPop3Server)
LeaveOnServer      = 13      // POP3: leave on server (LeaveOnPop3Server)
JunkScore          = 14
FetchBodyFromServer = 15     // IMAP/POP3: fetch full body (FetchBodyFromPop3Server)
CopyToFolder       = 16
AddTag             = 17      // Add a tag/keyword
KillSubthread      = 18
MarkUnread         = 19
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

## 5. MCP Tools

> **Status: implemented.** The tools below (`listFilters`, `createFilter`,
> `updateFilter`, `deleteFilter`, `reorderFilters`, `applyFilters`) live in
> `extension/mcp_server/api.js` — that file is canonical, and
> `extension/mcp_server/schema.json` is the authoritative tool schema. The
> `inputSchema` blocks here are design intent only. Enum values live in Section 3
> and union-member handling in Section 6; do not inline pseudocode in this file —
> point at the handler.

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

**Implemented in** `extension/mcp_server/api.js` — `listFilters()` / `serializeFilter()`. Numeric attrib/op/action constants are mapped to names via the `ATTRIB_NAMES` / `OP_NAMES` / `ACTION_NAMES` reverse maps; typed term values are read through `readSearchValue()` (correct union member per attribute — see Section 6).

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

**Implemented in** `extension/mcp_server/api.js` — `createFilter()`, which builds terms via `buildTerms()` and actions via `buildActions()`. Attribute/operator/action names are resolved through the strict allow-lists `ATTRIB_MAP` / `OP_MAP` / `ACTION_MAP` (the real `nsMsgSearchAttrib` / `nsMsgRuleActionType` values from Sections 3 — no `parseInt` fallback). Condition values are written by `setSearchValue()`, which dispatches on the attribute to the correct `nsIMsgSearchValue` union member (`.date`, `.age`, `.size`, `.priority`, `.status`, `.junkStatus`, `.junkPercent`, else `.str`) per Section 6.

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

**Implemented in** `extension/mcp_server/api.js` — `updateFilter()`. When only some fields change, it rebuilds the filter via remove+insert and **copies** the untouched terms/actions. The copy path goes through `setSearchValue()`/`readSearchValue()`, so typed condition values (date/age/size/etc.) are preserved across an update.

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

**Implemented in** `extension/mcp_server/api.js` — `deleteFilter()`.

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

**Implemented in** `extension/mcp_server/api.js` — `reorderFilters()`.

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

**Implemented in** `extension/mcp_server/api.js` — `applyFilters()`, via `nsIMsgFilterService.applyFiltersToFolders(filterList, [folder], null)`.

**Note**: `applyFiltersToFolders` is asynchronous — it returns immediately while processing continues; move/copy actions may not complete instantly.

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

Sources — verified against upstream @ `72b8ba0` 2026-07-23: member names from [`nsIMsgSearchValue.idl`](https://github.com/mozilla/releases-comm-central/blob/72b8ba0761b3881d926be53f77fbf75d5a9316d5/mailnews/search/public/nsIMsgSearchValue.idl); which member each attribute uses from the [`IS_STRING_ATTRIBUTE`](https://github.com/mozilla/releases-comm-central/blob/72b8ba0761b3881d926be53f77fbf75d5a9316d5/mailnews/search/public/nsMsgSearchCore.idl#L195-L202) macro and the match code in `nsMsgSearchValue.cpp` / `nsMsgSearchTerm.cpp`.

`nsIMsgSearchValue` is a **tagged union**: the value lives in a different member
depending on `attrib`. You must set `value.attrib` first, then write the member
that matches — writing the wrong one (e.g. `.str` on a Date attrib) makes XPCOM
throw `NS_ERROR_ILLEGAL_VALUE` and fails the whole operation. This is the spec
`setSearchValue()` / `readSearchValue()` implement; keep them in sync with this
table.

| Attribute (`nsMsgSearchAttrib`) | Member | Notes |
|---|---|---|
| Date (3) | `value.date` | PRTime, **microseconds** since epoch (`Date.parse(x) * 1000`) |
| Priority (4) | `value.priority` | numeric constant |
| MsgStatus (5) | `value.status` | bitmask |
| AgeInDays (12) | `value.age` | whole days (int32) |
| Size (14) | `value.size` | bytes (uint32) |
| HasAttachmentStatus (44) | `value.status` | fixed flag `nsMsgMessageFlags.Attachment` = `0x10000000`; has/hasn't chosen by operator, value ignored |
| JunkStatus (45) | `value.junkStatus` | |
| JunkPercent (46) | `value.junkPercent` | |
| Subject (0), From (1), Body (2), To (6), CC (7), ToOrCC (8), AllAddresses (9), Keywords/tag (16), OtherHeader (52) | `value.str` | the string attributes |

Verified live end-to-end (create + read-back round-trip, incl. `updateFilter`
copy path) 2026-07-23.

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

- [x] Add constant lookup maps (ATTRIB_NAMES, OP_NAMES, ACTION_NAMES and reverse maps)
- [x] Implement `listFilters(accountId)` handler
- [x] Implement `createFilter(...)` handler
- [x] Implement `updateFilter(...)` handler
- [x] Implement `deleteFilter(accountId, filterIndex)` handler
- [x] Implement `reorderFilters(accountId, fromIndex, toIndex)` handler
- [x] Implement `applyFilters(accountId, folderPath)` handler
- [x] Add all 7 tool definitions to the `tools` array
- [x] Add all 7 dispatch cases to `callTool()` switch
- [x] Test each tool with real Thunderbird instance
- [x] Handle edge cases (canHaveFilters, null accounts, empty filter lists)
- [x] Ensure `saveToDefaultFile()` called after all mutations
- [x] Verify `searchTerms` iteration works on target Thunderbird version
