# Thunderbird MCP

[![CI](https://github.com/TKasperczyk/thunderbird-mcp/actions/workflows/ci.yml/badge.svg)](https://github.com/TKasperczyk/thunderbird-mcp/actions/workflows/ci.yml)
[![Tools](https://img.shields.io/badge/50_Tools-email%2C_compose%2C_filters%2C_calendar%2C_contacts-blue.svg)](#what-you-can-do)
[![Localhost Only](https://img.shields.io/badge/Privacy-localhost_only-green.svg)](#security)
[![Thunderbird](https://img.shields.io/badge/Thunderbird-128%2B-0a84ff.svg)](https://www.thunderbird.net/)
[![License: MIT](https://img.shields.io/badge/License-MIT-grey.svg)](LICENSE)

Give your AI assistant full access to Thunderbird -- search mail, compose messages, manage filters, and organize your inbox. All through the [Model Context Protocol](https://modelcontextprotocol.io/).

<p align="center">
  <img src="docs/demo.gif" alt="Thunderbird MCP Demo" width="600">
</p>

> Inspired by [bb1/thunderbird-mcp](https://github.com/bb1/thunderbird-mcp). Rewritten from scratch with a bundled HTTP server, proper MIME decoding, and UTF-8 handling throughout.

---

## Why?

Thunderbird has no official API for AI tools. Your AI assistant can't read your email, can't help you draft replies, can't organize your inbox. This extension fixes that -- it exposes 50 tools over MCP so any compatible AI (Claude, GPT, local models) can work with your mail the way you'd expect.

Compose sends and event/task creation require review by default because **Block `skipReview`** starts enabled. `skipReview: true` is honored only after you explicitly disable that safety setting. The same block refuses `saveDraft` with `sendAt` for scheduled delivery through Send Later; plain drafts remain available. Re-enabling the block does not cancel already scheduled drafts. While enabled, meetings with attendees other than the calendar user are read-only through MCP, and adding attendees is blocked: calendar servers can email invitations, updates, or cancellations without review. Automatic Forward/Reply filter actions have a separate opt-in, disabled by default.

---

## How it works

```
                    stdio              HTTP (localhost:8765-8774)
  MCP Client  <----------->  Bridge  <--------------------->  Thunderbird
  (Claude, etc.)           mcp-bridge.cjs                    Extension + HTTP Server
```

The Thunderbird extension embeds a local HTTP server with session-scoped auth tokens. The Node.js bridge translates between MCP's stdio protocol and HTTP, discovering the port and token automatically via a connection file. The bridge handles MCP lifecycle methods (initialize, ping) locally, so clients can connect even before Thunderbird is fully loaded.

The bridge starts when executed directly or loaded by a desktop client's Node bootstrap. Tests or library consumers that only need its exports must set `THUNDERBIRD_MCP_NO_AUTOSTART=1` before requiring `mcp-bridge.cjs`.

---

## What you can do

### Mail

| Tool | Description |
|------|-------------|
| `listAccounts` | List all email accounts and their identities |
| `listFolders` | Browse folder tree with message counts and `isFavorite` in object or table format. Filter by account, subtree, or `favoritesOnly: true`; favorites nested under ordinary folders are included. |
| `searchMessages` | Search by subject, sender, recipient, body preview, date range, or tags. Multi-word queries are AND-of-tokens (every word must appear somewhere). Prefix with `from:`, `subject:`, `to:`, or `cc:` to restrict to one field. Set `searchBody: true` for full-text body search via Thunderbird's Gloda index. Supports `includeSubfolders`, `countOnly`, and offset-based pagination. Results include `threadId` and `preview` snippet. By default, `dedupByMessageId` collapses the same RFC Message-ID found in multiple folders/labels into one row and reports the other folder paths in `dupLocations`; set `dedupByMessageId: false` to return every location. Rows show the first 3 recipients and first 3 Cc addresses, each followed by `(+N more)` when shortened; `fullRecipients: true` returns the full lists. RFC Message-ID values and `mid:` links are not matched here; use `findMessageById`. |
| `findMessageById` | Find a message by RFC Message-ID or a `mid:local@domain` link. Accepts `mid:`, angle brackets, and one level of percent-encoding. A stored id that already contains `@` is not percent-decoded. Searches local folder databases, trying Inbox, Sent, Drafts, and Archive before other folders, then Gmail All Mail, with Junk and Trash last. Returns `id` and `folderPath` for `getMessage`. Set `allLocations` to return every copy. Optional `folderPath` limits the search to one folder and its subfolders. |
| `getMessage` | Read full email content -- `bodyFormat`: `markdown` (default), `text`, or `html`. Set `rawSource: true` for the complete RFC 2822 source (all headers + MIME parts). Optional attachment saving. Set `includeInlineImages: true` to append supported inline CID images as MCP image blocks (PNG, JPEG, GIF, or WebP; max 1 MiB base64 per image and 4 MiB total). Skipped images are reported in attachment metadata. |
| `getMessages` | Read full email content for up to the configured batch limit in one call (default 10, max 20). Uses the same `bodyFormat`, `rawSource`, and attachment options as `getMessage`; each item supplies `messageId` and `folderPath`. |
| `getMessageHeaders` | Read one message's headers by `messageId` and `folderPath`, without parsing its body. Includes decoded Cc, raw `references`, parsed parent `inReplyTo`, size, numeric `threadId`, read/flagged state, and tags. |
| `batchGetMessageHeaders` | Read headers for up to 200 IDs in one folder. Returns an ID-keyed `headers` map with per-ID errors; incomplete lookups carry `notScanned` rather than establishing absence. |
| `searchAttachments` | Find messages by attachment filename substring and/or Content-Type prefix. Searches complete local messages up to 16 MiB of raw MIME within structure and work limits, never decrypts, and reports omissions. Returns `messages` with shared headers, `folderPath`, and matching attachment names, declared types, decoded sizes and dispositions. |
| `getRecentMessages` | Get recent messages with date, unread, and tag filtering. Supports pagination. Results include `threadId` and `preview`; recipient lists are capped like `searchMessages` rows unless `fullRecipients: true`. |
| `refreshFolder` | Refresh an individual IMAP, built-in Exchange (EWS/Graph), or Owl folder. Reports counts, completion status, and whether a successful server update was confirmed. Account roots are refused. |
| `displayMessage` | Open a message in Thunderbird's GUI -- `3pane` (default), `tab`, or `window` mode |
| `updateMessage` | Mark read/unread, flag/unflag, add/remove tags, move or copy between folders, or trash -- supports bulk via `messageIds`. `copyTo` preserves the source and can add a Gmail label. |
| `deleteMessages` | Delete messages -- drafts are safely moved to Trash |
| `createFolder` | Create new subfolders to organize your mail |
| `createVirtualFolder` | Create a saved search over accessible real folders without moving or copying mail. Existing folder names or storage aliases, Windows short-name patterns, special root names, hidden servers, and unsuitable parents are refused. |
| `listVirtualFolders` | List saved searches with account keys, scope, and conditions. Omit the entire view if any scope is inaccessible, hidden, wildcard, unresolved, or unreadable. |
| `deleteVirtualFolder` | Detach a saved-search view and clean inspected storage nonrecursively. Refuse real folders, subfolders, own headers, regular backing files, shared, linked, nonempty or unreadable storage. Partial results list retained paths and persistence status. |
| `renameFolder` | Rename an existing mail folder |
| `deleteFolder` | Delete a folder (moves to Trash, or permanently deletes if already in Trash) |
| `moveFolder` | Move a folder to a new parent within the same account |
| `emptyTrash` | Permanently delete all messages in Trash (including subfolders) |
| `emptyJunk` | Permanently delete all messages in Junk/Spam (including subfolders) |

`getMessage` and `getMessages` read messages up to 50 MiB, retrieving IMAP messages online when no complete offline copy is available. Other providers need readable local message content. Each request has a 120-second budget, shared by all items in a batch; synchronous Thunderbird operations can delay timeout handling. Raw-source batches have a 50 MiB aggregate source limit; structured batches have a 50 MiB output limit. Results explicitly identify messages that timed out, exceeded the remaining output budget, or were not attempted. Unreadable or oversized messages return an error rather than partial content.

`searchAttachments` accepts `nameContains` and `contentType`, both case-insensitive; provide at least one, and both must match when supplied. Omit `folderPath` to search accessible accounts' Inboxes. `maxResults` defaults to 50 and is capped at 200; `scanCap` defaults to 5,000 and is capped at 50,000 messages per folder. The attachment flag is a heuristic: messages without it are not inspected. Collection visits native message keys in reverse order, keeps at most `maxResults × 4` candidates per folder, then selects at most that many across folders by date. Returned matches are sorted by date within that inspected subset, which need not include the mailbox's newest messages.

Attachment searches use a shared 20-second budget and yield between candidates. Each candidate must have a complete local snapshot and pass strict raw MIME checks; attachment metadata is extracted from those same bytes without Gloda, a message URI or another stream. Snapshots over **16 MiB of raw MIME, including encoded attachments**, are skipped as `tooLarge`. A cheap structural preflight runs before classification; `tooComplex` counts candidates exceeding 256 MIME entities (including the root and embedded-message roots), nesting depth 10, 64 KiB of headers per entity and 256 KiB total, 300,000 logical line breaks, or 32 MiB of aggregate traversal/decoding work. Each entity charges its full raw length, so nested bodies count repeatedly; quoted-printable bodies charge three additional lengths to bound decoding cost. These limits can omit complex messages below 16 MiB. Extraction also yields while walking parts and counting encoded payload sizes.

Remote messages without Thunderbird's Offline flag are counted in `notOffline` and never downloaded. Enable offline use and sync the folder before retrying. `getMessage` can separately fetch an individual uncached IMAP message within its own limits; doing so does not make an attachment search complete. Native database, storage and synchronous classification work can delay the best-effort deadline.

Search never decrypts, even with encrypted-message access enabled. Encrypted, opaque-signed, ambiguous or otherwise unverifiable contents are withheld. Top-level signed mail is also withheld while the opt-in is off; with it on, strictly verified clear-signed mail may be searched. No later native MIME tree can replace the checked contents. Access, message identity and privacy settings are rechecked before publishing.

Matching attachments are MIME leaf parts with a filename or `attachment` disposition, plus whole attached messages (`message/rfc822`, `message/news`, or `message/global`). Named inline images and named text parts qualify, including singlepart messages. Attached-message contents are not flattened; detached signatures inside `multipart/signed` and unnamed inline parts are omitted. MIME defaults apply when a type is absent. Unlike Gloda, search uses declared/default MIME types and decoded header filenames only: it does not infer types or names, invent names for unnamed attachments, or discover uuencoded/BinHex/yEnc attachments in body text. `matchingAttachments` entries have `name` (possibly empty), `contentType`, `size` (decoded bytes), and `disposition` (empty when absent). External-attachment stubs have `size: null`; their targets are never fetched.

The response always includes `messages`, `totalScanned`, `candidates` and `truncated`. Any observed omission sets `truncated: true` with an explanation. Nonzero counts distinguish `protectedWithheld`, `tooLarge`, `tooComplex`, `notParsed` (unavailable, failed, timed-out or unattempted candidates), `notOffline`, and `scanFailures` (failed or inaccessible folder scans). `scanCapHit` and `candidateLimitHit` identify collection limits; `overflow` counts inspected matches omitted by `maxResults`. A timeout is explained in the message. `truncated: false` means the flagged candidates in the scanned local snapshot were covered, not that every attachment in the mailbox was found. Attachment search was contributed by [@JordanRO2](https://github.com/JordanRO2) in [#149](https://github.com/TKasperczyk/thunderbird-mcp/pull/149).

Starting in 0.10.0, `searchMessages` and `getRecentMessages` shorten each recipient list by default to help choose which messages to read. Pass `fullRecipients: true` when enumerating participants or otherwise using complete address lists; omitted, false, or null values select summaries. `getMessage`, `getMessages`, `getMessageHeaders`, and `batchGetMessageHeaders` continue to return complete lists. Matching still uses the full headers, so a `to:` or `cc:` query can find an address beyond the three shown. A summary containing `(+N more)` is display text, not a complete address list.

Lists with at most three mailboxes and full-mode output keep their previous text, including any encoded Cc names. Longer lists are parsed from the original headers and rendered with decoded names; quoting and group labels may change. If parsing or rendering fails, that field keeps its full text and the message remains in the results. The cap therefore limits recipient count in successfully formatted summaries, not bytes in every response.

`refreshFolder` returns `status: "loaded"` when Thunderbird reports loading; `updateConfirmed: true` additionally means a folder-update URL reported success. An unconfirmed first load, including a native busy update or synchronous "No update ran" notification, gets one trailing update within the caller's remaining timeout. Concurrent callers share that update without extending their individual deadlines. The final outcome can still be unconfirmed, fail or time out; generic providers that expose only loaded events cannot confirm server success. `netMessageDelta` is the net count change during this call, not a count of arrivals. Account roots, saved searches and no-select folders are refused; individual local, POP, RSS and news folders are skipped.

`getMessageHeaders` and `batchGetMessageHeaders` accept the `id` returned by `getMessage` or `searchMessages`, with optional surrounding angle brackets. Returned `id` values preserve Thunderbird's stored spelling and case, normally without brackets; batch map keys preserve the caller's input. `inReplyTo` is the last parsed reference without brackets, or an empty string when there is no parent. Thunderbird stores In-Reply-To in its references field when a References header is absent. To pass a non-empty parent ID to `saveDraft`, wrap it in angle brackets; References must be space-separated bracketed IDs within `saveDraft`'s limits. To draft a reply to the fetched message itself, use `<id>` as `inReplyTo` and append `<id>` to its returned references chain. The raw references string is preserved, so malformed or oversized chains may need correction before `saveDraft` accepts them.

Header lookups first use the folder's Message-ID index. Misses share one cooperative fallback scan with a best-effort 20-second budget, including when an index entry is missing for a surviving duplicate. A completed scan can establish "not found in this folder"; an incomplete scan reports `notScanned`, and batch results also include `truncated: true` and a message. Verify `folderPath` before retrying: an ID may belong to another folder, and requesting fewer IDs does not generally reduce the work for an absent ID. Batch `total` counts requested IDs including duplicates, while `failed` counts distinct error keys. Account access is rechecked after yields; disabling encrypted-message access during the scan also withholds previously collected protected subjects. These tools do not parse message bodies, but may read up to 16 KiB of stored outer headers per protected subject. Native database/index operations and key snapshots are synchronous, so the budget cannot interrupt them.

A saved search is a live view in Thunderbird's folder pane. Its message counts populate after it is first opened. MCP message tools read folder storage, not saved-search results; use `searchMessages` on the source folders for retrieval. Saved searches cannot be move/copy destinations or parents for `createFolder`/`moveFolder`.

`deleteVirtualFolder` checks both the database and raw storage. An mbox backing path must be absent or Thunderbird's empty directory placeholder with no entries at all; a maildir may contain only empty `cur`/`tmp` directories; the folder's `.sbd` directory must be absent or empty. Regular backing files, even zero-byte files, are refused because nonrecursive file removal could still discard mail appended after inspection. Storage shared with another registered folder is refused, including hidden and restricted accounts. Ownership checks also refuse unreadable storage and links in registered storage paths or their ancestors. Unexpected contents, unsupported storage types, and the experimental Panorama backend are refused.

Deletion detaches the view once without deleting storage, saves the updated saved-search list, and removes only inspected items with nonrecursive removal. A directory that gains contents is retained. A complete result has `viewRemoved: true`, `persisted: true`, `cleanupComplete: true`, and `partial: false`. A failure after detaching returns `partial: true`, `viewRemoved: true`, persistence status, and `leftoverStorage` paths; an unreadable final state lists paths that could not be verified as absent. Do not retry deletion after the view has been removed. Review retained storage manually: it may be rediscovered as a folder after restart, and a persistence failure may restore the saved search itself.

`createVirtualFolder` uses the same condition vocabulary as `createFilter`. `booleanAnd: true` means AND with the previous term; `false` means OR (default true). Use the same value for all conditions when editing the view through Thunderbird's single all/any choice. Supply individual real folders, not account roots or other saved searches; duplicate folders are searched once. Names containing `~` followed by a digit are refused on every platform because they can address existing folders through Windows short-name aliases. Storage identities are also checked before configuring the view. Address-book conditions (`isInAB`/`isntInAB`) require **Allow all address books** while account restrictions are active. A failed native creation can leave an empty folder; the tool never deletes a folder to undo a failed creation.

`searchMessages` scans cooperatively with a best-effort 20-second budget. Header searches collect and sort matches before pagination, including those beyond 10,000; only the returned page loads full result fields. `countOnly` returns a best-effort count of currently resolvable, non-expunged collected candidates, respecting `dedupByMessageId` and including candidates beyond the requested page. It rechecks candidates cooperatively without loading result content. If the deadline cuts that recheck short, only validated candidates count and the response reports `truncated: true`; the count can be zero if collection already used the budget. Counts need not equal a page's length or `totalMatches`, which describes collected candidates before page hydration. Folders can change during a search, even if no truncation is reported. Opening a native database and taking its key snapshot are synchronous and cannot be interrupted by this budget.

Scoped header searches may wait up to 5 seconds for the named folder's sync, within the same 20-second search budget. This wait applies only to the explicitly named, selectable IMAP folder; account roots, saved-search (virtual) folders, no-select folders and non-IMAP folders do not wait. Each page or count request has one bounded wait, which can include one shared trailing update after an unconfirmed load, as with `refreshFolder`. A native update already in progress may have an older server snapshot; searches await the trailing update before scanning, while their budget allows. Unscoped searches and descendant scans keep their background IMAP refresh and use the locally available cache. `searchBody` uses Gloda without requesting or awaiting a folder refresh; a folder sync does not establish Gloda index freshness.

Count and paginated responses can include `folderRefresh`, containing the final refresh outcome: `status` (`loaded`, `timeout` or `error`), `joinedInFlight`, and optional `updateConfirmed`, `warning` and `error` fields. Only `updateConfirmed: true` confirms a successful server update. Offline calls still return promptly with cached results. A trailing update that cannot finish in time or again reports "No update ran" does not claim freshness. Unconfirmed refreshes retain the usual background refresh fallback. This freshness outcome is separate from scan completeness: a complete scan of the cache need not be truncated because its refresh was unconfirmed. Plain-array responses remain unchanged and carry no refresh metadata.

Message results remain a plain array when `offset` is omitted or null, including incomplete searches and `searchBody: true`. To receive completeness information, provide `offset: 0` (or another offset) for a paginated object. Only object responses can include `truncated: true` and a `message` explaining why; `countOnly` retains its existing object response and can also include these fields. Partial counts and `totalMatches` cover only observed matches, and date ordering covers only that subset. `hasMore` describes additional pages within the collected results and becomes false at their end, even if `truncated` is true; an empty returned page also sets it false. To improve coverage, narrow the folder, date range, or query (or set `includeSubfolders: false`).

For `searchBody: true`, object responses are conservatively marked `truncated: true`: Gloda returns candidates ranked by relevance and exposes no reliable completeness indicator. It applies its limit before removing stale index rows, so even a short result can omit other matches. Plain-array responses carry no completeness information. Narrowing the text query improves coverage; use header search when header/preview matching is sufficient.

`updateMessage` takes tag **keys**, not display labels: for example, `addTags: ["my=20project"]` uses a key generated for a label with a space. Keys must be non-empty printable ASCII without spaces, parentheses, brackets, braces, `%`, `*`, double quotes, backslashes, `<`, `>`, or `;`. This follows [Thunderbird's keyword handling](https://searchfox.org/comm-central/source/mailnews/imap/public/nsIImapService.idl), including its extra restrictions beyond [RFC 3501 atoms](https://www.rfc-editor.org/rfc/rfc3501#section-9). `=` and legacy modified UTF-7 keys containing `&` are accepted. Invalid entries in either `addTags` or `removeTags` fail the entire call with an error naming those entries before any messages are changed.

Use `copyTo: "<destination folder URI>"` to copy without removing the original. On Gmail, copying from Sent Mail to a project folder adds that label while preserving Sent Mail and existing labels. `copyTo`, `moveTo`, and `trash: true` are mutually exclusive. Both source and destination must be accessible. Copies and moves within or across accounts use the same Thunderbird native copy service call. On IMAP, completion is asynchronous and later failures are not reported back; success means only that the operation was submitted, so verify the destination. Tags applied together with a copy or move may not transfer on IMAP; verify the destination tags as well.

Message body formats (`getMessage` and `getMessages`):

- `markdown` (default) converts HTML to visible structure, allowing only `http:`, `https:`, and `mailto:` link destinations. Other HTML links become their visible text. HTML images become alt text (or nothing); their source URLs are omitted. Text nodes and alt text escape backslashes, backticks, brackets, angle brackets, and `!` before `[` to prevent literal links, images, or raw HTML. A literal `!` immediately before a generated link is also escaped, including across element boundaries. Other punctuation (including underscores, asterisks, hashes, pipes, and tildes) is preserved; HTML code spans and blocks use code formatting.
- For plain-text MIME bodies, `markdown` escapes Markdown image openers (`![`) while preserving the rest of the text, apart from the existing invisible-character removal. This also applies to coerced text and raw-MIME recovery. Other Markdown, links, and literal HTML in plain-text bodies are **not sanitized**; clients must disable raw HTML and remote-content loading when rendering untrusted message text.
- `text` extracts visible text. Both text formats remove hidden HTML content and invisible control characters.
- For `text` and `markdown`, HTML input over 2 MiB (UTF-8) is truncated at a Unicode character boundary before being parsed with the same HTML privacy rules. The output ends with `[Message body truncated at 2 MiB]`. Content after the cut is omitted; retained content keeps the normal formatting and link rules. Oversized compose fragments use the same capped parsing path. Encryption checks inspect the full source: oversized HTML containing `-----BEGIN PGP MESSAGE-----` anywhere is withheld unless encrypted-message access is enabled, even when the marker is only quoted in prose.
- `html` returns the original HTML unchanged. It is **untrusted** and may contain hidden content, scripts, unsafe links, and remote images; clients must sanitize it before rendering. `rawSource: true` also remains unchanged.

Structured MIME extraction keeps the first body's representation outside multipart/alternative and joins later fragments of that same type in message order, including after inline attachments. A footer of another type cannot replace the main body. Alternatives select the requested representation; attached text files are excluded using Thunderbird's attachment metadata. Encryption checks also classify the joined content before exposing it.

`includeInlineImages: true` is a separate opt-in for supported inline CID images as MCP image blocks; it does not add image URLs to the Markdown body.

### Compose

| Tool | Description |
|------|-------------|
| `sendMail` | Compose a new email -- opens a review window; direct sending requires explicitly disabling the `skipReview` safety block |
| `dryRunCompose` | Read-only `sendMail` preflight with the same arguments and first refusal: resolves identity, counts effective recipients, inspects attachments, and assesses encoded size. Requires `sendMail` enabled |
| `saveDraft` | Save a new or replacement draft without opening a window; supports threading headers, optional delayed delivery through Send Later (`sendAt`), and reports an accessible Drafts folder |
| `replyToMessage` | Reply with quoted original and proper threading -- `skipReview` is subject to the same safety block; `saveAsDraft` saves the threaded reply to Drafts without sending |
| `forwardMessage` | Forward with all original attachments preserved -- `skipReview` is subject to the same safety block |

`sendMail`, `replyToMessage`, and `forwardMessage` open a window for you to review and edit before sending by default. The **Block `skipReview`** preference is on by default, so `skipReview: true` is rejected until you explicitly disable the preference. The block also rejects `saveDraft` with `sendAt`. Direct sends and scheduled drafts are subject to the identity encryption checks under [Send Safety](#access-control). Attachments can be file paths or inline base64 objects.

Direct sends go through Thunderbird's compose pipeline without opening a window, so provider senders such as Owl apply. The selected identity's Reply-To, automatic Cc/Bcc, return receipt, delivery status notification (DSN), and vCard settings apply too. Automatic Cc/Bcc addresses are not echoed; `dryRunCompose` includes them in its recipient counts.

Use `dryRunCompose` before `sendMail` to check the add-on's preparation without sending, saving, opening a window, creating native attachments, or writing/removing files. It takes the same arguments (including identity-based `isHtml` defaults) and reports the same first preflight refusal. This is not an SMTP or provider delivery check and does not predict other Thunderbird prompts or later changes to settings/files.

The report includes `resolvedIdentity: { id, email, name }`, effective parsed-mailbox counts in `recipients` (including enabled automatic Cc/Bcc; empty groups do not count), and metadata for **caller attachments only** in `attachments`. `inputBodyLength` counts UTF-16 code units in the supplied body. `estimatedEncodedBytes`, when available, is a conservative byte estimate including prepared formatting/signature, headers, identity-added content such as a vCard, and encoding overhead; it is not the input length or an exact MIME size. An unbounded estimate is `null`. Size concerns are blockers for a dry direct send and `warnings` for a dry review window. No subject or signature-setting echo is returned. `sendAt` remains exclusive to `saveDraft`.

Allowed unreviewed sends, including Send Later deliveries, are never encrypted or cryptographically signed, including for identities set to sign by default. The text signature controlled by `includeSignature` is separate from cryptographic signing. Use `skipReview: false` for the review window, where Thunderbird applies its own encryption and signing decisions.

Direct sends report success only after delivery succeeds, and Thunderbird marks originals replied or forwarded. If Thunderbird's 120-second send timeout expires, the outcome is unknown: check Sent and the Outbox before retrying. The bridge waits 150 seconds for direct sends, draft saves, `getMessage`, `getMessages`, and `displayMessage`, and 30 seconds for other calls. A timeout or lost connection after submitting a mail operation also reports an unknown outcome; check Sent/Outbox for sends or Drafts for saves before retrying.

An explicit `isHtml: true` or `false` selects HTML or plain-text composition. Omitted or `null` uses the identity's compose-format preference, treating the supplied body as plain text. Concurrent review calls keep their own bodies, recipients, and attachments. Native forward reviews open in sequence until each window is identified. If an opening fails or times out before identification, the queue recovers automatically when that window initializes or closes, or after a 15-second recovery period. Late windows are left untouched and cannot receive another call's content.

`replyToMessage` accepts `saveAsDraft: true` to build a native reply with quoted text, the identity's signature and threading headers, save it, and close the compose window without sending. This requires the `saveDraft` tool to be enabled and cannot be combined with `skipReview`. Protected or unverifiable originals require a review window, even with encrypted-message access enabled. Before saving, the current compose identity's configured destination must be accessible under account restrictions and carry the Drafts flag. A timeout or revoked access during saving returns `saveOutcome: "uncertain"`: the outstanding save may still complete, so check Drafts before retrying. Failures and timeouts restore the window's prior close and save-dialog behavior.

`saveDraft` supports these optional parameters:

- `sendAt`: schedule delivery through Send Later using an ISO 8601 date-time. Requires `sendMail` enabled and **Block `skipReview`** off, and obeys the identity encryption checks. Omit it or pass `null` for a plain draft. See [Delayed delivery](#delayed-delivery-with-send-later) for prerequisites and cancellation.
- `inReplyTo`: one bracketed Message-ID such as `<original@example.com>`. `references`: up to 100 such IDs, oldest first, separated by single ASCII spaces. Each ID is limited to 998 characters and the full References value to 16,384. Missing brackets, extra tokens, whitespace inside IDs, and control characters (including CR/LF) are rejected, never repaired. References defaults to `inReplyTo` when omitted, and may also be supplied independently. The caller supplies the subject and quoted text.
- `replaceMessageId` and `replaceFolderPath`: replace an existing draft with the supplied content. The folder must be accessible under account restrictions, carry Thunderbird's Drafts flag, and match the selected `from` identity's configured drafts folder. Other folders and missing messages are rejected before saving. Supply all fields and attachments you want retained; content is not merged from the old draft.
- `includeSignature`: append the identity's signature, defaulting to `true` for a new draft and `false` for a replacement so re-saving a fetched draft does not duplicate its signature. Set it explicitly when needed. `sendMail` with `skipReview: true` accepts the same option, defaulting to `true`; review windows use Thunderbird's own signature preferences. The separator honors the identity's "suppress signature separator" setting and is omitted when a separator line already exists in the signature's plain-text rendition. Signature files are limited to 1 MiB; missing, unreadable, oversized, or image signature files are omitted without falling back to inline signature text.

On success, `saveDraft` returns `folderPath` when Thunderbird exposes the destination and that folder is accessible under account restrictions. Otherwise the save still succeeds without disclosing the folder URI. It does not return the saved message's ID; look up the draft in the returned folder when available.

### Delayed delivery with Send Later

When the [Send Later](https://github.com/Extended-Thunder/send-later) add-on is installed, `saveDraft` accepts `sendAt`, an ISO 8601 date-time (`2027-02-01T10:00` is read as local time; `Z` or an explicit offset such as `+02:00` are accepted). The draft carries the `X-Send-Later-At` and `X-Send-Later-Uuid` headers, and Send Later handles delivery without a review window. The result returns `sendAt` in UTC, at the same one-second precision as the saved header. Success confirms the draft was saved, not that it has been delivered.

Scheduling is refused before saving, signature processing or attachment preparation unless every condition holds:

- **Block `skipReview` is off and `sendMail` is enabled.** A scheduled draft is an unreviewed send.
- **At least one To, Cc or Bcc mailbox is supplied.** Empty address groups are refused; Cc-only and Bcc-only schedules are supported.
- **The selected identity passes the unreviewed-send encryption checks.** Required encryption, automatic encryption for a configured key/certificate, or unreadable encryption settings refuse scheduling. Signing by default alone does not refuse it; scheduled deliveries remain unsigned. Omit `sendAt` for a plain draft, or use `sendMail` with `skipReview: false` for Thunderbird's encryption and signing in the review window.
- **Send Later recently answered.** Its readiness is refreshed at startup and every minute, with a ten-second timeout. Failed or timed-out refreshes clear readiness; superseded and late answers are ignored. Information older than three minutes is unavailable for scheduling.
- **Send Later's delivery options are on:** both "send scheduled drafts" and "send unsent messages", with a check interval above zero.
- **The date is valid and at most 366 days ahead.** Impossible calendar dates and local times in daylight-saving gaps are refused. Supply an explicit offset to distinguish repeated fall-back times. A one-minute past-time tolerance accommodates clock skew; such a draft can be sent on the next check.

Thunderbird and Send Later must be running for delivery. After a restart, overdue drafts are subject to Send Later's late-delivery limit (48 hours by default); older drafts are held for rescheduling. Drafts in accounts deactivated in Send Later are not sent. Thunderbird MCP cannot detect that account setting through Send Later's external API. Servers that strip custom headers (Exchange through its IMAP gateway, Proton Mail Bridge) can prevent scheduling.

Before Send Later processes the draft, delete it to cancel, or replace it through `replaceMessageId` and `replaceFolderPath` with a new `sendAt` to reschedule. A replacement with omitted or `null` `sendAt` removes its scheduling headers. A preflight refusal leaves an existing draft unchanged, including its previous schedule. **Re-enabling Block `skipReview` does not cancel already scheduled drafts.** Disabling `sendMail` or changing identity preferences afterward also does not revoke them; these checks run when saving the schedule.

A timeout or lost connection after submitting the save has an unknown outcome. Check Drafts, Outbox and Sent before retrying a scheduled save to avoid duplicate delivery.

Compose tools validate the `from` identity strictly -- if the specified sender doesn't match any configured Thunderbird identity, the tool returns an error instead of silently substituting another account.

### Filters

| Tool | Description |
|------|-------------|
| `listFilters` | List all filter rules with human-readable conditions and actions |
| `createFilter` | Create filters with structured conditions (from, subject, date...) and actions (move, tag, flag...) |
| `updateFilter` | Modify a filter's name, enabled state, conditions, or actions |
| `deleteFilter` | Remove a filter by index |
| `reorderFilters` | Change filter execution priority |
| `applyFilters` | Start eligible enabled Manual filters on a folder; report submitted rules and skipped rules with reasons |

Your AI can create sorting rules, adjust priorities, and run them on existing mail. Changes persist after validation; all updates, including name or enabled-state edits, validate a complete candidate before replacing the existing rule. Unparseable rules cannot be updated, but can still be deleted. Filter names, conditions, and action text reject control characters (U+0000–U+001F and U+007F) and backslashes. Custom add-on actions cannot be created, preserved by an update, or submitted for execution.

Date conditions accept only `YYYY-MM-DD` as a local calendar day; date-times are rejected. `listFilters` reports a lone native ALL term as `matchAll: true` with an empty `terms` array. Compound rules retain their terms and Boolean operators, including ALL terms; an empty rule is not reported as matching all messages.

**Allow automatic Forward/Reply filter actions** is off by default in the extension settings. While off, `createFilter` and `updateFilter` reject any resulting rule containing Forward or Reply, including disabled rules or edits that retain an existing sending action. An update that only sets `enabled: false` can still disable a sending rule, and deletion remains available. Enabling this setting permits automatic sends without a compose review window, independently of **Block `skipReview`**. Turning it off does not disable saved filters or stop Thunderbird's own automatic filtering.

The sending-action setting governs rules MCP creates, modifies, or runs manually. Reordering or deleting a rule containing `StopExecution` can change which of your existing rules Thunderbird runs automatically, including Forward/Reply rules.

Move/Copy destinations must remain accessible under the current account restrictions, including actions retained by an update. An update that only sets `enabled: false`, or deletion, remains available for rules with inaccessible destinations.

New or replaced Move/Copy actions store the resolved folder's canonical URI, including when the supplied URI uses another accepted spelling.

`applyFilters` submits only enabled, parseable rules with the Manual type flag and skips rules with inaccessible, Outbox, or saved-search Move/Copy destinations, rules with address book conditions while address books are restricted, and Forward/Reply rules while the setting is off. Creation and updates also refuse saved-search destinations, including retained actions; disabling, deleting, or repairing an existing rule remains possible. It returns `submittedFilters`, submitted rule names, and skipped rule names with reasons. An eligible rule containing a Custom action rejects the call before any rule is submitted. A successful submission means processing has started, not completed; if no rules are eligible, nothing is submitted.

### Contacts

| Tool | Description |
|------|-------------|
| `searchContacts` | Search contacts across all address books by any of their email addresses or by name and return full contact details. Supports `maxResults`. |
| `getContact` | Read full contact details by UID, including all email addresses (`emails`, in preference order) |
| `createContact` | Create a contact with optional email addresses/name, phones, postal addresses, organization, title, note, and birthday. Phone-only contacts are supported. |
| `updateContact` | Update contact fields; omitted fields stay unchanged, while empty email/phone/address arrays clear those collections |
| `deleteContact` | Delete a contact by UID |

Contacts return `email` as Thunderbird's primary address and `emails` as all non-empty stored addresses in Thunderbird's preference order, with an optional `type: "work" | "home"`. Searches match any email address or name; multiple words must all match somewhere in those fields.

On create and update, `emails` accepts typed or untyped `{ address, type? }` entries. Local and CardDAV address books preserve all supplied addresses and types. Other address-book types retain the legacy limit of two addresses without types; a third address is refused. On update, omitted fields stay unchanged, `emails` replaces the whole collection in the supplied order, and `emails: []` clears it. If `email` is supplied too, it must equal the first trimmed address (or be `""` with an empty collection). Email values containing control characters are refused before changes are made.

Updating `email` alone with a non-empty value replaces the primary address and keeps the others. **`email: ""` clears all addresses on vCard-backed cards, but only the primary on legacy flat cards.** Use `emails: []` to explicitly clear the whole collection on either kind of card.

Reads preserve duplicate addresses already stored in Thunderbird; writes reject duplicates case-insensitively, so remove duplicates before sending a returned list as an update. The `emails` list trims address whitespace and omits empty entries, while `email` preserves Thunderbird's primary value. An empty preferred entry or surrounding whitespace can therefore make `emails[0].address` differ from `email`; reading a contact never changes its primary address to force agreement.

If a stored vCard cannot be read, contact reads and searches fall back to cached fields. Updates are refused before changing or saving the card, so they cannot overwrite unrelated data lost by the failed parse.

`createContact`, `updateContact`, and `deleteContact` are refused until you enable **Allow contact writes** in the extension settings (off by default, including after upgrading). Contact changes are saved immediately without a review step. Reading contacts does not require this permission; existing account and Tool Access restrictions still apply.

### Calendar

| Tool | Description |
|------|-------------|
| `listCalendars` | List all calendars with disabled, read-only, event, and task support flags |
| `refreshCalendar` | Refresh one calendar or all enabled refreshable calendars, using their providers (built-in Exchange, Owl, CalDAV, Google). Reports completion, provider errors, or an unconfirmed timeout. |
| `createEvent` | Create an event through a review dialog, optionally with RRULE recurrence and display `reminders`. Direct creation and non-empty `attendees` require disabling **Block `skipReview`**. Accepts `status: tentative \| confirmed \| cancelled`. |
| `listEvents` | Query events by date range with bounded recurrence expansion. Returns a plain array capped at `maxResults`, including status, recurrence, reminders, organizer, attendees, and your participation status. Series that cannot be expanded return a master marked `recurrenceNotExpanded: true`. |
| `updateEvent` | Modify an event or series; `recurrenceId` selects one occurrence. Meetings with other attendees are read-only while **Block `skipReview`** is on, even when `attendees` is omitted. |
| `deleteEvent` | Delete an event or series; `recurrenceId` excludes one occurrence. Meetings with other attendees cannot be deleted while **Block `skipReview`** is on. |
| `createTask` | Open a pre-filled task dialog for review; direct creation via `skipReview` requires explicitly disabling the default safety block |
| `listTasks` | List tasks/to-dos from calendars -- filter by completion status, due date, or calendar |
| `updateTask` | Update a task's title, due date, description, priority, completion status, or percent complete |

`recurrence` accepts a single RRULE, such as `FREQ=WEEKLY;BYDAY=MO,TU` or `RRULE:FREQ=DAILY;COUNT=10`. The optional prefix is case-insensitive. Control characters (including CR/LF), malformed rules, `SECONDLY`/`MINUTELY`, and `HOURLY` on all-day events are rejected; Thunderbird's recurrence parser validates the rule before saving. On update, `recurrence: ""` or `null` clears the rule. Replacing a rule discards existing EXDATEs and modified occurrences.

`refreshCalendar` returns `loaded` with a warning if a provider reports an error and then completes. If it reports an error but no completion within the wait, the result is `error` carrying that provider error. A silent provider returns `timeout`, meaning completion is unconfirmed. In all-calendar mode, inspect the individual results; disabled and non-refreshable calendars are excluded.

Both refresh tools refuse requests while Thunderbird is offline. `timeoutMs` defaults to 15000 and is clamped to 1000–25000. Concurrent calls and retries share an operation while it is younger than 25 seconds; joining does not renew that age, and a later call starts a fresh attempt once it is stale. A timeout does not cancel Thunderbird's work. Completion is best effort: a timer cannot interrupt a provider's synchronous work.

`reminders` is an array of whole minutes before the event starts (`0` is at the start), for example `[1440, 0]` for one display reminder a day before and one at the start. Duplicates are ignored, at most 10 entries are accepted, and each value must be between 0 and 40320 (four weeks). Invalid input is rejected before saving. On `createEvent`, omit it, or pass `null` or `[]`, to add no alarms.

`listEvents` and `updateEvent` manage the same subset: display alarms relative to the start, at or before it. On update, an array replaces that subset and `[]` removes it; omitting `reminders` or passing `null` preserves all alarms. End-relative, absolute, after-start, email, and other non-display alarms are not listed and are preserved by updates, including their provider metadata. Modified occurrences keep their own reminders when the series reminders change; update those with `recurrenceId`. Reminders-only writes to meetings with other attendees are refused while **Block `skipReview`** is on.

Only alarms stored on the event are listed. Returned minutes are unique and sorted largest first. Existing second-based offsets are rounded to the nearest whole minute (half-minutes round up) for listing; listing and updates with `reminders` omitted or null leave the stored offsets unchanged. Replacing the managed set writes whole-minute alarms. If an event's alarms cannot be read, its listed `reminders` is `[]`; an update that cannot read existing alarms fails without saving.

The selected calendar's declared alarm limit applies to the total of retained alarms plus the requested reminders; a limit of 0 permits no alarms. Requests that exceed it are refused, never truncated. New display reminders are also refused when the calendar declares supported alarm actions without DISPLAY. Retained alarms are never removed or filtered to satisfy these capabilities.

Thunderbird creates the initial **Home** calendar disabled. Check `listCalendars[].disabled` and enable the calendar in Thunderbird's calendar properties before using it. Event and task tools report an error for an explicitly selected disabled calendar; listings and default direct creation use enabled calendars, and report an error if all calendars are disabled. MCP leaves the enabled/disabled setting unchanged.

Use the `recurrenceId` returned by `listEvents` to update or delete one occurrence. Omitted or null IDs select the series; a non-null ID cannot be combined with `recurrence` on update. All-day occurrence IDs preserve their calendar date across timezones. Existing Date-compatible inputs remain accepted for event start/end dates; ISO 8601 is recommended.

Attendees use `{ "email": "alice@example.com", "name": "Alice", "role": "optional" }`; new attendees default to `required`, and a `mailto:` email prefix is accepted. Email addresses must be single mailboxes without control characters or URI headers; names cannot contain control characters either. On update, omitted/null attendees preserve the list, while `[]` removes everyone. Retained attendees match by case-insensitive email and keep their participation status and provider metadata. Only supplied name/role fields change; omitted/null fields preserve existing values, and an empty name clears it. New attendees start with `NEEDS-ACTION`.

Attendee management requires the calendar identity to match an existing organizer. Existing organizers are preserved; a missing organizer is initialized from the calendar identity, which may also set a missing calendar `organizerId` property. Creation with non-empty attendees requires disabling **Block `skipReview`**, even when `skipReview` is false: Exchange/Owl or CalDAV may email invitations containing the event title and description without review.

`listEvents` returns a plain array, capped at `maxResults` (default 100, maximum 500). RRULE generation is limited to `maxResults + 1` candidates per series (hard ceiling 501) and 5,000 per request before expansion can allocate an unbounded occurrence array. Generation limits can leave fewer than `maxResults` results even when more occurrences exist, especially when exclusions remove generated candidates. Use narrower date ranges to reduce this effect.

If a series cannot be expanded safely with a bounded count, expansion fails, or the shared generation budget is exhausted before that series, its master is included once with `recurrenceNotExpanded: true`, subject to the same output cap. This includes EXRULE series: incomplete exclusion expansion could expose excluded dates, so their rules are not expanded. A flagged master carries the series' original dates, which may be outside the requested range; it does not represent a matching occurrence. Ordinary EXDATEs and modified occurrences remain supported during bounded expansion.

Each event includes `organizer: { id, commonName }`, up to 100 `attendees`, `attendeeCount` with the full count, and `myParticipationStatus` (empty when unavailable). Do not use a truncated attendee list as an update replacement, since that would remove the omitted attendees.

### Access Control

| Tool | Description |
|------|-------------|
| `getAccountAccess` | View which accounts the MCP server can access |

Account and tool access are configured via the extension settings page (Tools > Add-ons > Thunderbird MCP > Options). Access control is not MCP-exposed -- only the user can change it.

Starting with 0.10.0, updates respect customized **Tool Access**: if you have disabled at least one tool, newly introduced tools start unchecked and cannot be called until you enable them in Options and save. Fresh installations and profiles that have never restricted tools keep all tools enabled. Once recorded, customization remains remembered even if you later enable every tool; only tools first encountered in a later update start disabled. A tool you explicitly enable stays enabled across restarts and updates.

If policy prevents saving the new restrictions, Options shows a note that Tool Access is managed by policy. Newly encountered tools stay disabled while existing tools keep following the policy's list. Migration completes after the preference is unlocked and its effective value can be updated. An untouched profile with a locked empty list keeps all tools enabled.

An observed restriction remains remembered if policy later empties the disabled list, including after a restart. If Tool Access history itself cannot be saved, newly encountered tools stay disabled until saving works; existing tools continue following their effective settings.

Two compatibility limits apply. A 0.9.x profile that disabled tools and then re-enabled everything looks identical to an untouched profile, so its first upgrade keeps all tools enabled. Downgrading to 0.9.3 and saving Tool Access there can rewrite the disabled list and discard restrictions for newer tools; review those choices after upgrading again.

`dryRunCompose` also requires `sendMail` to be enabled: disabling its parent hides it from `tools/list` and refuses calls. Options reflects this dependency immediately, including unsaved parent changes, while preserving the child's separate enabled/disabled choice.

The same settings page has a "Send Safety" section. **Block `skipReview`** is enabled by default and rejects `skipReview: true` for `sendMail`, `replyToMessage`, `forwardMessage`, `createEvent`, and `createTask`. It also refuses `saveDraft` with `sendAt`; re-enabling the block does not cancel already scheduled drafts. It also blocks adding attendees and makes meetings with attendees other than the calendar user read-only through MCP: `updateEvent` and `deleteEvent` reject every write, including time/description/recurrence edits, attendee removal, and occurrence deletion, even when `attendees` is omitted. Occurrence writes check both the series and selected occurrence; series writes also check stored exceptions. Events without attendees and events whose only attendee is the identified calendar user remain editable. If the calendar identity cannot be established, existing attendees cannot be treated as self-only. Disabling the setting permits these writes, which may email updates or cancellations without review. Event/task creation review dialogs remain available without attendees.

Even with **Block `skipReview`** disabled, unreviewed mail is refused before sending or scheduling when:

- The sending identity's effective encryption policy is greater than zero, or its encryption settings cannot be read reliably. An absent identity policy inherits Thunderbird's default identity policy.
- Thunderbird's automatic encryption setting (`mail.e2ee.auto_enable`) is on and the identity has an OpenPGP key or S/MIME encryption certificate configured. An unreadable automatic encryption setting also refuses for those identities; identities without a key or certificate are unaffected by that setting.
- A reply or forward's original is encrypted or protected, including opaque-signed S/MIME, or its MIME classification is missing or unknown. This check never requests decryption and applies even when encrypted-message reads are enabled. Clear-signed mail remains allowed when the existing MIME checks establish that its contents are clear.

For sends, these refusals direct callers to `skipReview: false`, where Thunderbird applies its own encryption in the review window. For `saveDraft`, omit `sendAt` to save a plain draft, or send through `sendMail`'s review window. The checks do not reproduce every compose-window encryption decision. Signing by default, attaching a public key, draft encryption, and subject protection settings alone do not cause refusals. Automatic reply drafts (`replyToMessage(saveAsDraft: true)`) also refuse protected originals, even with encrypted-message access enabled, because these drafts are saved unencrypted. Review windows and plain `saveDraft` calls without `sendAt` retain their existing behavior.

For an unreviewed reply or forward, the complete stored original must pass raw MIME checks before Thunderbird's structured parser runs. Preflight retrieves uncached IMAP originals as raw bytes without MIME conversion, allows at most 50 MiB of raw MIME and 50 MiB of structured body text, and refuses HTML above the existing 2 MiB parsing limit. If retrieval fails, use review or make the folder available offline and retry. Incomplete, truncated, oversized, or unclassifiable content requires review. Preflight has a 15-second deadline; a late parser callback cannot resume a refused call. Caller attachment files are accessed or created only after all encryption checks pass.

Additional guards apply to direct `skipReview: true` mail calls:

- `sendMail` refuses zero effective recipients, counting parsed mailboxes and enabled identity automatic Cc/Bcc. Empty review windows remain allowed.
- `sendMail`, `replyToMessage`, and `forwardMessage` refuse when the conservative encoded-size estimate exceeds a positive `mailnews.message_warning_size`. Equality is allowed. An unreadable threshold or an estimate that cannot be safely bounded also requires review; a threshold of zero or below disables this size gate. The add-on never changes the preference.

The estimate includes MIME/encoding overhead, dynamic headers, prepared body/signature and caller attachments. Replies and forwards additionally budget from the complete raw original, including original forward attachments, and check the assembled quote before materializing caller attachments. Estimates deliberately err high and can require review even when the eventual MIME file would fit. Use `skipReview: false` for these refusals. Thunderbird's native large-message dialog can outlive a send timeout, which is why the add-on checks before unreviewed delivery. These guards do not change Send Later scheduling checks or cancel existing schedules.

The separate "Filter Send Actions" section controls **Allow automatic Forward/Reply filter actions**, which defaults to off. Only the user can change this preference in the settings page; MCP clients cannot enable it.

The "Contact Writes" section controls **Allow contact writes**, which defaults to off and gates `createContact`, `updateContact`, and `deleteContact`. Only the user can change it; MCP clients cannot enable it. This permission is independent of Tool Access and **Allow all address books**. Under account restrictions, contact writes require both address-book access and this opt-in. Turning it off blocks the next write without undoing saved changes.

---

## Setup

### 1. Install the extension

```bash
git clone https://github.com/TKasperczyk/thunderbird-mcp.git
```

Install `dist/thunderbird-mcp.xpi` in Thunderbird (Tools > Add-ons > Install from File), then restart. A pre-built XPI is included in the repo -- no build step needed.

**Automatic updates:** From v0.7.3 on, the add-on auto-updates through Thunderbird's add-on update check. Thunderbird downloads updates in the background and applies them on the next restart; because this add-on uses an experiment API, updates are not live hot-swapped. v0.7.3 is the last build you need to install by hand because older builds have no `update_url` and cannot auto-discover it. Thunderbird ships with `xpinstall.signatures.required=false`, so unsigned auto-updates work out of the box; a profile hardened to require signatures blocks both manual and automatic installs. If updates do not arrive, check the Add-ons gear menu and make sure **Update Add-ons Automatically** is enabled.

### 2. Configure your MCP client

Add to your MCP client config (e.g. `~/.claude.json` for Claude Code):

```json
{
  "mcpServers": {
    "thunderbird-mail": {
      "command": "node",
      "args": ["/absolute/path/to/thunderbird-mcp/mcp-bridge.cjs"]
    }
  }
}
```

#### Connecting HTTP-capable clients directly

The extension serves MCP over HTTP itself (`POST /`: `initialize`, `tools/list`, `tools/call`, notifications), so clients that support HTTP transport can skip the bridge. Read the current `port` and `token` from `connection.json`, then replace `<port>` and `<token>` below:

```bash
claude mcp add --transport http thunderbird "http://127.0.0.1:<port>/" --header "Authorization: Bearer <token>"
```

- Read the port from `connection.json` after each server restart instead of pinning `8765`. The server tries ports 8765-8774; Options shows the selected port and a notice when it differs from the default.
- A stable token can survive restarts, but does not make the port stable. On shared machines, use the stdio bridge with automatic discovery instead of a pinned port with a stable token.
- Responses are plain JSON; there is no SSE stream and no `Mcp-Session-Id`, both optional in the Streamable HTTP spec.
- Path attachments are then read by the extension itself, under the same attachment policy as the bridge, and inside the Snap/Flatpak sandbox where applicable (the bridge is what makes host paths work for sandboxed installs).
- Keep this on localhost. For a client on another machine or container, prefer an SSH tunnel or a TLS proxy bound to loopback over "Listen on all interfaces": that mode is plain HTTP, so the token and mail content cross the network unencrypted.

### Sandbox-aware connection discovery

The bridge re-discovers `connection.json` on every cache miss. It tries these locations in order:

1. `THUNDERBIRD_MCP_CONNECTION_FILE`, if set
2. Native temp dir: `<os.tmpdir()>/thunderbird-mcp/connection.json`
3. macOS fallback: `/var/folders/*/*/T/thunderbird-mcp/connection.json` owned by the current user
4. Linux Snap: Thunderbird's live `TMPDIR` from `/proc/<pid>/environ`, plus the official snap fallback under `~/Downloads/thunderbird.tmp`
5. Linux Flatpak / Betterbird Flatpak: `$XDG_RUNTIME_DIR/app/<id>/thunderbird-mcp/connection.json` and `~/.var/app/<id>/cache/tmp/thunderbird-mcp/connection.json`

This covers native installs, the official Thunderbird snap, Thunderbird Flatpak, Thunderbird Beta Flatpak, and Betterbird Flatpak without changing the extension side. If multiple sandbox candidates exist at once, the bridge tries the newest file first. Flatpak discovery accepts only these application IDs: `org.mozilla.Thunderbird`, `org.mozilla.thunderbird`, `org.mozilla.thunderbird_esr`, `net.thunderbird.Thunderbird`, and `eu.betterbird.Betterbird`. Set `THUNDERBIRD_MCP_CONNECTION_FILE` to force a single explicit path.

Every connection file, including an explicit pin, must be a regular file of at most 4 KiB with a valid port and token. On Linux and macOS it must also belong to the bridge's user, have no group/other permissions (normally `0600`), and must not be a symlink. Rejected files are reported, never repaired or deleted. An invalid explicit pin does not fall back to discovery.

Example override:

```json
{
  "mcpServers": {
    "thunderbird-mail": {
      "command": "node",
      "args": ["/absolute/path/to/thunderbird-mcp/mcp-bridge.cjs"],
      "env": {
        "THUNDERBIRD_MCP_CONNECTION_FILE": "/absolute/path/to/connection.json"
      }
    }
  }
}
```

That's it. Your AI can now access Thunderbird.

---

### Outgoing attachments

For `sendMail`, `dryRunCompose`, `saveDraft`, `replyToMessage`, and `forwardMessage`, pass `attachments` as a JSON array, not a JSON-encoded string. The stdio bridge reads file-path attachments on the host and sends inline base64 to Thunderbird. The existing bridge limit is **18 MiB per path attachment**, now also applied to `saveDraft`. The extension's direct-HTTP path limit remains 50 MiB; inline base64 remains limited to 25 MiB of encoded data. Both transports enforce at most 20 attachments and 50 MiB of decoded attachments per operation.

For `dryRunCompose`, the extension inspects path metadata and inline base64 shape without decoding or materializing attachments. The stdio bridge still reads path attachments to inline them first; a refused path is a tool error rather than a report blocker. Disabling a tool does not prevent this pre-existing bridge path inspection. Hardening the bridge to consult tool availability before inspecting paths is outside this change.

UNC/network and device paths, dotfiles and dot-directories, application-data directories (Windows `AppData` and the macOS home `Library`), and credential/key filenames are refused. Windows alternate data streams and path components ending in dots or spaces are also refused. Files exported by `getMessage(saveAttachments: true)` can still be attached again, and ordinary local document paths continue to work. If any attachment is refused, missing, invalid, or over a limit, the entire operation fails before sending, saving a draft, or opening a review window. The error identifies refused entries; temporary files created during the failed conversion are removed. Callers that previously relied on partial attachment success must handle this error explicitly.

Files in hidden or application-data folders, including macOS `~/Library` and its cloud folders, must be copied to an ordinary folder before they can be attached by path.

External attachment links, including attachments removed with Thunderbird's **Detach** command and RSS/Atom feed enclosures, are refused by `getMessage` with `saveAttachments: true`, which reports "External attachment link refused" for the affected attachment. `includeInlineImages` also refuses them unless it can read a stored message part. A direct forward (`forwardMessage` with `skipReview: true`) is refused if the original includes one. Attachments stored in the message, including those from Exchange accounts, continue to work. A forward opened for review (`skipReview: false`) is built by Thunderbird and shows its attachments in the compose window.

## Security

- **Auth tokens**: The HTTP server requires a session-scoped bearer token. Generated on startup, written to `<TmpD>/thunderbird-mcp/connection.json` with 0600 permissions. The bridge re-discovers that file automatically across native installs, Snap, Flatpak, Betterbird Flatpak, and macOS temp directories.
- **Stable tokens**: Optional stable tokens persist across restarts. The stored token is cleared on uninstall only while the add-on is enabled (best effort). If you disable the add-on before uninstalling, regenerate the token before using it again.
- **Dynamic port**: Tries ports 8765-8774, records the actual port in the connection file. No hardcoded port dependency.
- **Account access control**: Restrict which email accounts are visible to MCP clients via the settings page. Changes take effect immediately. A corrupt stored setting is shown as blocked, and saving it unchanged does not lift the restriction.
- **Calendars and address books under account restrictions**: While any account restriction is active, calendar and address-book tools (listing, lookups by ID, and writes) are blocked unless you enable "Allow all calendars" or "Allow all address books" in the settings page. Both are off by default and grant access to all calendars or all address books respectively. Without an account restriction, these two permissions are not required. Tool Access and other safety settings still apply; "Allow all address books" does not grant contact writes.
- **Tool access control**: Disable specific tools via the settings page. Disabled tools are hidden from `tools/list` and blocked at dispatch.
- **Allow contact writes**: Off by default, including for existing users after upgrading. Enable it in Options to allow `createContact`, `updateContact`, and `deleteContact`. Reads remain available subject to existing access restrictions. Enabled contact tools stay listed while writes are blocked; enabling a tool in Tool Access does not grant this permission.
- **Encrypted-message access**: "Allow MCP clients to read encrypted messages" is off by default. While it is off, `getMessage` no longer lets Thunderbird decrypt an encrypted message, or ask for a key password. This also covers batch reads, raw-source reads, attachment saving, inline images, and `displayMessage`. Encrypted and opaque-signed S/MIME messages are withheld; mail that cannot be verified as unencrypted receives a distinct explanatory note. Clear-signed mail remains readable when its contents can be verified as unencrypted. Enabling the setting does not permit unreviewed replies, forwards, or automatic reply drafts of protected originals: open a reply or forward for review instead. While it is off, `searchMessages`, `getRecentMessages`, `getMessageHeaders`, and `batchGetMessageHeaders` return the outer subject (or a placeholder when unavailable) with `encryptedContentWithheld: true` for messages Thunderbird has recorded as OpenPGP-encrypted. Search and recent-message results omit previews; header searches do not match their cached protected subjects or previews, `searchBody` excludes them, and `displayMessage` refuses encrypted messages. This relies on Thunderbird's recorded OpenPGP status, so a message decrypted without that status being recorded may still show a cached decrypted subject.
- **Localhost only**: By default, the server binds to localhost only. The "Listen on all interfaces" option in settings binds to all IPv4 interfaces for WSL, Docker, or remote access. **This exposes the MCP server to every device on your local network.** Only enable on trusted networks. Auth token is always required.
- **Attachment-search privacy**: `searchAttachments` never decrypts or downloads a candidate, regardless of encrypted-message access. Signed mail requires the opt-in and strictly verified clear contents; protected, malformed or unverifiable candidates make the search incomplete and are counted without exposing their attachment names.
- **Auto-update integrity**: Auto-update is a code-delivery channel whose integrity depends on continued control of the GitHub repository, the GitHub Actions token, and the `tomaszkasperczyk.name` registration.

Reply and forward review windows follow Thunderbird's normal behavior when quoting an original for the user's review. Native filter execution follows Thunderbird's own settings. `searchBody` searches the existing index; it does not read full messages on demand.

---

## Troubleshooting

| Problem | Fix |
|---------|-----|
| Extension not loading | Check Tools > Add-ons and Themes. Errors: Tools > Developer Tools > Error Console |
| Connection refused | Make sure Thunderbird is running and the extension is enabled |
| Bridge can't find `connection.json` | Set `THUNDERBIRD_MCP_CONNECTION_FILE` explicitly if your environment uses a non-standard temp/runtime path |
| Server uses a port other than 8765 | Check Options or read `port` from `connection.json`; update direct HTTP clients to that port. The stdio bridge discovers it automatically. |
| AnkiConnect and Thunderbird MCP conflict | AnkiConnect also defaults to `127.0.0.1:8765`. Start Anki/AnkiConnect before Thunderbird so MCP can choose another port, then use the recorded MCP port or the stdio bridge. |
| Bridge reports that a port is not answered by Thunderbird MCP | Check the server's current port in Options and `connection.json`, and whether another application is using that port. Reconnect the client after correcting its configuration. |
| Missing recent emails | IMAP folders can be stale. Click the folder in Thunderbird to sync, or right-click > Properties > Repair Folder |
| Tool not found after update | If you restrict Tool Access, new tools start unchecked. Enable the tool in Thunderbird MCP Options, save, then reconnect MCP (`/mcp` in Claude Code). Also enable any required parent tool. |
| `searchBody` returns no results | IMAP accounts need offline sync enabled for Gloda to index message bodies |
| `rawSource` fails on IMAP | Messages without a complete offline copy are retrieved online. Check connectivity, or make the folder available offline and retry. Reads are limited to 50 MiB per message. |

### Release channel and Experiment API add-ons

This add-on uses Thunderbird's Experiment APIs. The add-on team has [announced plans to disable them on the Release channel](https://thunderbird.topicbox.com/groups/addons/T5426c1d2b0ba520c); Thunderbird has since [reported a postponement](https://blog.thunderbird.net/2026/06/thunderbird-monthly-development-digest-june-2026/), so do not assume every Release version blocks them.

If Thunderbird disables the add-on for this reason, use [Thunderbird ESR](https://www.thunderbird.net/thunderbird/all/) (Extended Support Release), which this add-on supports. It has been verified working on **140.16.0esr** and **153.3.1esr**. Check **Tools > Add-ons and Themes** to confirm the add-on is enabled, then open its Options page for the server status and any startup error.

---

## Development

See [CONTRIBUTING.md](CONTRIBUTING.md) for build, isolated test, and pull request instructions.

```bash
# Build the extension
./scripts/build.sh

# Test via the bridge (handles auth automatically)
echo '{"jsonrpc":"2.0","id":1,"method":"tools/list"}' | node mcp-bridge.cjs

# Test the HTTP API directly.
# On Snap / Flatpak / Betterbird Flatpak / macOS, point CONN_FILE at the
# real file or export THUNDERBIRD_MCP_CONNECTION_FILE first.
CONN_FILE="${THUNDERBIRD_MCP_CONNECTION_FILE:-/tmp/thunderbird-mcp/connection.json}"
TOKEN=$(jq -r .token "$CONN_FILE")
PORT=$(jq -r .port "$CONN_FILE")
curl -X POST http://127.0.0.1:$PORT \
  -H "Content-Type: application/json" \
  -H "Authorization: Bearer $TOKEN" \
  -d '{"jsonrpc":"2.0","id":1,"method":"tools/list"}'
```

**Dev-only extension reload:** After changing extension source locally, remove the add-on from Thunderbird, restart, reinstall the XPI, and restart again. Thunderbird caches aggressively. Regular users should install the latest release once and let auto-update handle later releases.

---

## Project structure

```
thunderbird-mcp/
├── mcp-bridge.cjs              # stdio <-> HTTP bridge (auth, port discovery)
├── extension/
│   ├── manifest.json
│   ├── background.js           # Extension entry point
│   ├── httpd.sys.mjs           # Embedded HTTP server (Mozilla)
│   ├── options.html            # Settings page UI
│   ├── options.js              # Settings page logic
│   ├── icons/                  # Extension icons
│   └── mcp_server/
│       ├── api.js              # All 49 MCP tools + auth + access control
│       └── schema.json
├── test/                       # Test suite (node:test, zero dependencies)
└── scripts/
    ├── build.sh
    └── install.sh
```

## Known issues

- IMAP folder databases can be stale until you click on them in Thunderbird
- HTML-only emails are converted to plain text (original formatting is lost)
- IMAP folder operations (rename, delete, move) are async -- verify with `listFolders` after
- Combining tags with move/trash on IMAP may not preserve tags on the moved copy -- use separate calls
- Thunderbird itself still runs pre-existing filters with cross-account move/copy targets automatically; the MCP filter tools refuse to update such rules (except disabling or deleting them) and `applyFilters` skips them
- `searchBody` on IMAP without offline sync only searches headers (Gloda limitation)
- IMAP messages without a complete offline copy may need an online connection for full-message reads

---

## License

This project uses the [MIT license](LICENSE). The bundled `extension/httpd.sys.mjs` is derived from Mozilla's HTTP server and remains under [MPL-2.0](https://mozilla.org/MPL/2.0/); its license notice is retained in the file.
