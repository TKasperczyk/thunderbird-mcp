# Thunderbird Mail Filters — API Research & Implementation Guide

Initial research completed 2026-02-21. The current validation and execution rules below supplement the interface notes and historical sketches.

## TL;DR

Thunderbird exposes full filter CRUD + execution via XPCOM interfaces. Our extension already uses an Experiment API with full XPCOM access — no new permissions needed. We can list, create, modify, delete, reorder, and manually apply filters.

### Current tool behavior

- **Forward/Reply require user opt-in.** `extensions.thunderbird-mcp.allowFilterSendActions`
  defaults to `false` and fails closed if unreadable. The extension settings page exposes
  it as **Allow automatic Forward/Reply filter actions** through
  `getAllowFilterSendActions()` / `setAllowFilterSendActions(boolean)`, not MCP tools.
  It is independent of `blockSkipReview`: native filter sends have no compose review
  window. Turning it off does not change saved filters or Thunderbird's own automatic
  filter execution.
- **Check the resulting rule.** While the preference is off, create/update rejects a
  candidate containing Forward or Reply, even if disabled or if the edit merely retains
  an existing sending action. An update that only sets `enabled: false` can disable a
  sending rule. Deleting a rule remains available under the existing account/tool permissions.
- **Custom actions are unsupported.** They may be listed, but cannot be created, copied
  by an update (including a disable-only update), or submitted for execution. Enabling
  Forward/Reply does not enable Custom actions.
- **Validate before replacing.** All updates, including metadata-only changes, build and
  validate a complete candidate before replacing the existing rule with `setFilterAt`.
  The description and temporary flag are preserved. Rejected validation leaves the
  original rule unchanged. Unparseable rules cannot be updated; deletion remains available.
  Filter names and condition/action strings that will be persisted reject U+0000–U+001F,
  U+007F, and backslash; normal Unicode and quotes remain supported.
- **Apply a selected temporary list.** `applyFilters` submits enabled rules with the
  Manual flag, skipping disabled, non-manual, unparseable, and disallowed sending rules.
  An otherwise eligible rule with a Custom action rejects the entire call before
  submission. Results include `submittedFilters` (count), `submitted` (rule names), and
  `skipped` (`{ name, reason }` entries). No candidates means no native apply call.
  Submission starts asynchronous processing; it does not report completion.

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
- `setFilterAt(index, filter)` — Replace the filter at a position
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
- `label` — Legacy numeric label value for a Label action
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

`nsMsgFilterTypeType` is a signed 32-bit `long`. Explicit `type` values in
`createFilter` and `updateFilter` must be positive integers at most 2147483647,
using only the known flags resolved by name from the running Thunderbird's `Ci`.
The native range is checked before bitmask operations or native assignments, so
overflow cannot silently become a different type. The allowed mask includes
PostPlugin, PostOutgoing, Archive and Periodic; `nsMsgFilterType.All` omits them.

### Search Attributes (nsMsgSearchAttrib)

Resolve attributes by their native names, such as `Ci.nsMsgSearchAttrib.Sender`,
`Ci.nsMsgSearchAttrib.AgeInDays`, and `Ci.nsMsgSearchAttrib.Keywords`. The values are
not a contiguous sequence. `FILTER_ATTRIBUTE_DEFS` maps tool names to these native
names and their value codecs; `FILTER_ATTRIBUTES` contains the constants available in
the running Thunderbird. See the attribute and union-accessor reference in Section 6.

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
SoundsLike     = 11    // Phonetic match (LDAP)
LdapDwim       = 12    // LDAP do-what-I-mean
IsGreaterThan  = 13    // Size comparison
IsLessThan     = 14    // Size comparison
NameCompletion = 15    // LDAP name completion
IsInAB         = 16    // Is in address book
IsntInAB       = 17    // Not in address book
IsntEmpty      = 18
Matches        = 19    // Generic match for custom terms
DoesntMatch    = 20    // Generic non-match for custom terms
```

### Filter Actions (nsMsgFilterAction)

Resolve action IDs by native name, such as `Ci.nsMsgFilterAction.MoveToFolder`,
`Ci.nsMsgFilterAction.CopyToFolder`, `Ci.nsMsgFilterAction.Forward`, and
`Ci.nsMsgFilterAction.Reply`. Do not infer their values from list order or maintain a
separate numeric map. `FILTER_ACTION_DEFS` supplies tool/native names and typed value
members; `FILTER_ACTIONS` resolves the available constants. The Action Value Setting
reference in Section 6 describes the correct accessors, including legacy `action.label`.
Custom actions are unsupported, and Forward/Reply use the separate opt-in described above.

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

The early schema and listing sketches in this section illustrate the interface shapes.
Production uses the runtime constant tables and validation rules described above and in
Section 6; numeric fallback parsing and direct use of an unfiltered stored list are not
supported tool paths.

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

**Important**: Map numeric attribute/operator/action IDs to human-readable names using
the production `ATTRIB_NAMES`, `OP_NAMES`, and `ACTION_NAMES` tables, which are generated
from named native constants. Values must be read through their typed accessors; the
historical listing sketch above omits that dispatch. Do not recreate numeric maps.

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
      type: { type: "integer", description: "Filter type bitmask (default: 17 = inbox + manual). 1=inbox, 16=manual, 32=post-plugin, 64=post-outgoing" },
      conditions: {
        type: "array",
        items: {
          type: "object",
          properties: {
            attrib: { type: "string", description: "Attribute: subject, from, to, cc, toOrCc, allAddresses, body, date, priority, status, size, ageInDays, hasAttachment, junkStatus, junkPercent, tag, otherHeader" },
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
      insertAtIndex: { type: "integer", description: "Position to insert (0 = top priority, default: end of list)" },
    },
    required: ["accountId", "name", "conditions", "actions"]
  }
}
```

**Implementation behavior:** resolve the account through the access checks, validate the
name, and build the conditions and actions on a candidate filter with the runtime
allowlists and typed value codecs. Validate the complete candidate, including the
Forward/Reply permission and persisted text, before inserting it into the list and
saving. Custom actions and numeric action/attribute/operator fallbacks are unsupported.

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
      filterIndex: { type: "integer", description: "Filter index (from listFilters)" },
      name: { type: "string", description: "New filter name (optional)" },
      enabled: { type: "boolean", description: "Enable/disable (optional)" },
      conditions: { type: "array", description: "Replace conditions (optional, same format as createFilter)" },
      actions: { type: "array", description: "Replace actions (optional, same format as createFilter)" },
    },
    required: ["accountId", "filterIndex"]
  }
}
```

**Implementation note — rebuilding.** All updates, including edits to name, enabled
state, or type alone, create a complete candidate before changing the stored list.
They preserve `filterDesc` and `temporary`, copy any omitted conditions/actions, and
validate the result before replacing the original with `filterList.setFilterAt`.
If saving throws, the original filter is restored in the in-memory list. An unparseable
rule is rejected before candidate construction; it may still be deleted.
The copy (`copySearchTerms` /
`copyActions` in `api.js`) preserves the supported properties `nsMsgFilter` writes to
`msgFilterRules.dat` — `attrib`, `op`, `booleanAnd`, `beginsGrouping`,
`endsGrouping`, `matchAll`, `arbitraryHeader`, `hdrProperty`, `customId` and the value
through the union member its attribute owns (see Section 6); actions copy `type`, the typed
member of that type and supported string data. Custom actions are rejected, and persisted
text is validated even when it came from the old rule. Failures propagate and abort the
update without changing the original. Move/Copy destinations in the resulting rule,
including retained actions, are checked with `getAccessibleFolder` against current account
restrictions. An update that only sets `enabled: false` bypasses this destination check;
deletion remains available too. An earlier copy read `.str` for everything and swallowed the resulting
`NS_ERROR_ILLEGAL_VALUE`, which silently reset priority, status, age, size, junk status and
junk percent conditions to 0 on every `updateFilter` — changing only the action of an
"age in days > 30" rule left "age in days > 0". The typed copy follows #175 (@rdkr).

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
      filterIndex: { type: "integer", description: "Filter index to delete (from listFilters)" },
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
      fromIndex: { type: "integer", description: "Current filter index" },
      toIndex: { type: "integer", description: "Final target index (0 = highest priority)" },
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

**Purpose**: Start eligible enabled Manual filters on a folder.

```js
{
  name: "applyFilters",
  title: "Apply Filters",
  description: "Start eligible enabled Manual filters and report submitted and skipped rules",
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

**Implementation behavior:** check account/folder access, then select enabled, parseable
rules with the Manual flag. Check each rule's Move/Copy destinations with
`getAccessibleFolder` and skip the whole rule if any destination is missing, restricted,
or cannot be resolved. Skip Forward/Reply rules while the preference is off. An
otherwise eligible Custom action rejects the operation before any submission. Put the
eligible rules into `getTempFilterList(folder)` and call
`applyFiltersToFolders(tempList, [folder], null)` only when the list is nonempty. The
persisted list is not handed to the execution service or modified for selection.

The result reports `submittedFilters` (number), `submitted` (names), and `skipped`
(`{ name, reason }` entries, with reasons `disabled`, `non-manual`, `sending`,
`inaccessible-destination`, or `unparseable`). These describe selection and submission, not completed processing or a
count of messages changed. No eligible rules means no native apply call.

The sending-action preference governs rules MCP creates, modifies, or runs manually.
It does not govern Thunderbird's automatic execution of saved rules. Reordering or
deleting a rule containing `StopExecution` can change which later existing rules run,
including rules that Forward or Reply.

---

## 6. Gotchas & Edge Cases

### Persistence
- **Always call `filterList.saveToDefaultFile()`** after mutations (create, update, delete, reorder). Without this, changes exist only in memory and are lost on restart.
- Mutate the filter list in-place. Don't try `server.setFilterList(newList)` — that doesn't persist properly.

### Filter List Access
- `server.getFilterList(null)` — pass `null` for msgWindow unless you need UI updates
- `server.getEditableFilterList(null)` — may differ from `getFilterList` in some contexts, but usually the same for local/IMAP accounts
- `server.canHaveFilters` — check this before attempting filter operations (news servers may not support filters)

### Search Term Attribute Values

`nsMsgSearchAttrib` (`mailnews/search/public/nsMsgSearchCore.idl`) is **not contiguous**
past `AllAddresses = 9` — indices 10/11/13 are `Location`/`MessageKey`/`FolderInfo`, the
LDAP address-book attributes occupy 17–33, and the junk/attachment/header attributes sit
in the 44–52 range:

| Name | Value | | Name | Value |
|---|---|---|---|---|
| `Subject` | 0 | | `AgeInDays` | 12 |
| `Sender` | 1 | | `FolderInfo` | 13 |
| `Body` | 2 | | `Size` | 14 |
| `Date` | 3 | | `AnyText` | 15 |
| `Priority` | 4 | | `Keywords` (tags) | 16 |
| `MsgStatus` | 5 | | `HasAttachmentStatus` | 44 |
| `To` | 6 | | `JunkStatus` | 45 |
| `CC` | 7 | | `JunkPercent` | 46 |
| `ToOrCC` | 8 | | `JunkScoreOrigin` | 47 |
| `AllAddresses` | 9 | | `HdrProperty` | 49 |
| `Location` | 10 | | `FolderFlag` | 50 |
| `MessageKey` | 11 | | `Uint32HdrProperty` | 51 |
| | | | `OtherHeader` | 52 |

Guessing these numbers is how `ageInDays` ended up pointing at `Location` and `tag` at
`AgeInDays`. Read them from `Ci.nsMsgSearchAttrib.<Name>` at runtime instead.

Tag conditions have no attribute of their own — Thunderbird stores tags as keywords, so a
tag condition is `Keywords` with the tag key (e.g. `"$label1"`) in `value.str`.

### Search Term Value Setting
- The `value` property on `nsIMsgSearchTerm` is an `nsIMsgSearchValue` object
- You must set `value.attrib` to match the term's attrib before setting the value content
- `nsIMsgSearchValue` is a **tagged union**: using an accessor that doesn't match the
  attribute's type throws
  `Component returned failure code: 0x80070057 (NS_ERROR_ILLEGAL_VALUE) [nsIMsgSearchValue.str]`

The authoritative dispatch is Thunderbird's own, in
`chrome/messenger/content/messenger/searchWidgets.js` (`save()` / `updateDisplay()`):

| Attribute | Accessor | Notes |
|---|---|---|
| `Priority` | `value.priority` | numeric constant |
| `MsgStatus` | `value.status` | `nsMsgMessageFlags` bitmask |
| `Date` | `value.date` | PRTime — **microseconds** since epoch |
| `AgeInDays` | `value.age` | integer days |
| `Size` | `value.size` | integer KB |
| `JunkStatus` | `value.junkStatus` | `nsMsgJunkStatus`: 0 unclassified, 1 good, 2 junk |
| `JunkPercent` | `value.junkPercent` | 0–100 |
| `HasAttachmentStatus` | `value.status` | always `nsMsgMessageFlags.Attachment`; `is`/`isnt` carries has/hasn't. A caller-supplied value is ignored by Thunderbird (`is` + `"false"` persists as `is,true`), so the tools refuse one |
| `FolderFlag` | `value.status` | copied from existing terms only; not advertised for creation |
| `Uint32HdrProperty` | `value.status` | copied from existing terms only, with `term.hdrProperty`; not advertised for creation |
| legacy `Label` | `value.label` | copied from existing terms only on versions that expose it; not advertised for creation |
| `Custom` (-2) | `value.str` | plus `term.customId`, the add-on's term id, which is what the `.dat` file names the term by. Read and copied, never created here |
| everything else | `value.str` | including `Keywords`/tags, `JunkScoreOrigin` and `OtherHeader`. This is also the union member for every attribute not in `IS_STRING_ATTRIBUTE`'s exclusion list (`nsMsgSearchCore.idl`) |

- `OtherHeader` additionally needs the header name in `term.arbitraryHeader`, otherwise the
  term never matches. `HdrProperty`/`Uint32HdrProperty` terms name their property in
  `term.hdrProperty`.
- **Dates are local days.** `nsMsgSearchTerm` writes `Date` values with
  `PR_LocalTimeParameters` as `%d-%b-%Y` and reads them back as local midnight, so the day is
  all that survives a restart. A date-only tool value (`YYYY-MM-DD`) is therefore parsed as a
  *local* calendar day — `Date.parse` would take it as UTC midnight, which is the previous day
  anywhere west of UTC (`"2026-01-01"` in America/Toronto was saved as `31-Dec-2025`; fix from
  #175, @ncrosty58). Tool writes reject date-times because their time and timezone would not
  survive native filter persistence. Bare numbers are refused: `"2026"` used to
  be taken as epoch milliseconds and saved as `01-Jan-1970`. Read-back reports a local-midnight
  value as `YYYY-MM-DD`, anything else as an ISO-8601 instant.
- **ALL is a term, not a rule-wide override.** Current
  [`nsMsgLocalSearch.cpp`](https://searchfox.org/comm-central/source/mailnews/search/src/nsMsgLocalSearch.cpp)
  evaluates ALL as true within its Boolean expression. An empty term list matches nothing
  for filtering (`MatchTerms` returns `!Filtering`), unlike an empty search. Accordingly,
  `listFilters` collapses only a lone ALL term to rule-level `matchAll: true`; compound
  conditions keep their real terms and ALL operators, and empty rules are not match-all.
- **Integers are strict and bounded.** Values are matched with `/^-?\d+$/` (no `parseInt`,
  which took `"30abc"` as 30 and `"1.5"` as 1). The
  [search-value IDL](https://searchfox.org/comm-central/source/mailnews/search/public/nsIMsgSearchValue.idl)
  defines `size` and `status` as unsigned 32-bit values and `age` as signed 32-bit.
  The tools accept `size` from 0 to 4294967295 KB, `age` from 0 to 2147483647 days, and
  `status` from 1 to 4294967295, rejecting overflow before assigning native fields.
  Narrower semantic limits still apply: `junkPercent` is 0–100, `junkStatus` 0–2 (or
  `junk`/`good`/`unclassified`), `priority` is `nsMsgPriority.lowest`..`highest` (2–6) and
  `status` is a non-zero `nsMsgMessageFlags` bitmask. Priority is signed 32-bit;
  junk status/percent and the copied-only folder flags, uint32 header properties and
  legacy Label values are unsigned 32-bit. Dates use signed 64-bit PRTime; parsed
  JavaScript dates in microseconds fit that native range. The schema hints spell the values out
  (`4=normal`, `2=replied`, `(KB)`), with the numbers resolved from `Ci.nsMsgPriority` and
  `Ci.nsMsgMessageFlags` by name — the same way the attribute ids are.

The implementation reads the whole vocabulary from the running Thunderbird instead of
hardcoding it: attribute ids are resolved by name from `Ci.nsMsgSearchAttrib` (**not** by
enumerating it — `Object.keys` on an interface object returns nothing in the extension
experiment context, even though Gecko's `IID_NewEnumerate` implements it), operator ids
likewise from `Ci.nsMsgSearchOp` (the IDL constant names map to our operator names by
lowering the first letter — all 21 match), and the
`createFilter`/`updateFilter` schema text for attrib, op, value and header is generated
from what that enumeration yields on each `tools/list`. An attribute the running version
does not define is simply not offered.

The attribute metadata in `FILTER_ATTRIBUTE_DEFS` in `api.js` records each API
name, the IDL constant to resolve against, and the `nsIMsgSearchValue` member/codec — the
value typing above has no queryable API and exists only in C++ and in Thunderbird's own
hardcoded UI dispatch, so it cannot be derived at runtime. There are deliberately **no
fallback ids**: `Ci` is guaranteed (api.js dereferences it at module load and would not
load without it), so the only way name resolution fails is the search interface being
absent or renamed — the same situation in which `nsIMsgSearchTerm`, `nsIMsgSearchValue`
and the filter list are gone and no filter tool can work regardless. Thunderbird's own
filter UI takes the same position: `searchWidgets.js`, `searchTerm.js` and
`FilterEditor.js` dereference these constants 49 times between them without a single
guard. When the interface is missing, the generated descriptions say so and the tools
refuse with a message naming the cause.

### Version compatibility (verified TB 102 → 154-beta, Aug 2026)

Checked by diffing `mailnews/search/public/*.idl` across comm-esr102/115/128/140,
comm-beta and comm-central, plus the commit history of the C++ implementations:

- `nsMsgSearchOp`: byte-identical across the entire range. No compatibility concern.
- `nsMsgSearchAttrib`: exactly one change in four years — `Label = 48` was removed in
  TB 115 (Bug 1802815). The runtime enumeration handles this class of change by
  construction: a constant the running version lacks is simply not offered.
- `nsIMsgSearchValue`: the `label` member went with it; TB ≥ 141 added a readonly
  `utf8Str` getter (Bug 1971060). None of the members we write were ever touched. The
  write path guards against a member disappearing (as `label` did) with a clear error;
  the read path degrades to the string form.
- `nsIMsgFilter`/`nsIMsgFilterList`: only string-type refinements
  (`ACString` → `AUTF8String`, Bug 1999822, TB ~146) — invisible to JS callers.
- The union-enforcement code in `nsMsgSearchValue.cpp` is unchanged since 2022.

Worth watching: the Panorama database rework converts **virtual folder** search terms to
SQL (`LiveViewFilters`, Bug 1971060 ff.). Message filter lists (`nsIMsgFilterList`, what
these tools use) are so far untouched by it.

### Action Value Setting
- `MoveToFolder` / `CopyToFolder`: set `action.targetFolderUri`
- `ChangePriority`: set `action.priority` (`nsMsgPriority.lowest`..`highest`, 2–6; anything
  else is written as "Change priority" with no value)
- `AddTag`: tag value goes in `action.strValue` (the keyword, e.g., `"$label1"` or custom tag keyword)
- `Forward` / `Reply`: email address / template URI in `action.strValue`
- Legacy `Label`: use the typed `action.label` accessor, not `action.strValue`
- `JunkScore`: set `action.junkScore` — `nsMsgRuleAction::SetJunkScore` rejects anything
  outside 0..100
- `Custom` (`nsMsgFilterAction.Custom`): `action.customId` names the add-on's
  `nsIMsgFilterCustomAction`; its optional argument is in `action.strValue`. These actions
  can be listed, but creating, copying by update, and submitting them are unsupported
- Actions like `MarkRead`, `MarkFlagged`, `StopExecution`, `Delete` have no value parameter;
  the tools refuse a value on them
- The typed accessors are guarded (`nsMsgFilter.cpp`): `priority` throws
  `NS_ERROR_ILLEGAL_VALUE` unless `type` is `ChangePriority`, `targetFolderUri` unless
  Move/Copy, `junkScore` unless `JunkScore`, and `label` unless `Label`. `strValue` and
  `customId` are unguarded. This is
  why writing and copying are table-driven off `FILTER_ACTION_DEFS` in `api.js`, and why the
  tools require a value for every action that takes one: Thunderbird itself saves
  "Move to folder" with no folder, and the rule then does nothing when it runs

### Async Considerations
- `applyFiltersToFolders` returns immediately — the actual filtering happens asynchronously
- For move/copy actions, the messages may not be relocated instantly
- Report submitted rule counts/names and skipped rules; do not claim processing completed

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
- Rejected C0 controls, DEL, and backslashes in new and copied persisted text
- Default-off Forward/Reply creation, retained actions in updates, and disable-only updates
- Retained Move/Copy destinations respect current access restrictions; manual runs skip inaccessible rules and disable/delete recovery remains available
- Numeric search values accept their boundaries and reject native overflow before mutation
- Custom actions rejected on create/update/apply, with deletion still available
- Candidate validation failures leave the original filter and list unchanged
- Metadata-only updates use a complete typed copy and preserve description and temporary state
- Unparseable rules reject updates, remain deletable, and are skipped during manual execution
- Manual execution selects only eligible rules, reports skips, and does not call native apply for an empty selection
- Typed copies of FolderFlag, Uint32HdrProperty, legacy Label terms, and legacy Label actions

---

## 8. Alternative: Condition String Parsing

Instead of structured conditions, we could accept raw condition strings:
```
"AND (from,contains,newsletter@) OR (subject,contains,[news])"
```

The native `filterList.parseCondition(filter, conditionString)` method can parse these
directly. The MCP tools accept structured conditions only, so callers cannot bypass the
attribute/operator allowlists and persisted-text checks through a raw condition string.

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
