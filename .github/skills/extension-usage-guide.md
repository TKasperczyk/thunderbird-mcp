---
name: "thunderbird-mcp-extension-usage"
description: "Guide for using Thunderbird MCP with AI assistants to archive emails and attachments to the local file system"
version: "1.1.0"
tags: ["thunderbird-mcp", "mcp", "email", "ai", "copilot", "attachments", "archiving"]
applyTo: []
---

# Thunderbird MCP Extension - Usage Guide

The Thunderbird MCP Extension allows AI assistants (Claude, Copilot, etc.) to interact directly with Thunderbird — searching and reading emails, archiving messages, exporting attachments, and more.

## 📋 Architecture

```
AI Assistant (Claude/Copilot) 
    ↓ MCP Protocol (stdio)
mcp-bridge.cjs (Node.js)
    ↓ HTTP localhost:8765
Thunderbird Extension
    ↓
Email Database
```

## 🚀 Setup

### 1. Install the Extension

Download the latest `thunderbird-mcp.xpi` from [GitHub Releases](https://github.com/TKasperczyk/thunderbird-mcp/releases/latest) and install it in Thunderbird:

> **Thunderbird → Tools → Add-ons → ⚙️ → Install Add-on From File**

Restart Thunderbird afterwards.

### 2. Install the MCP Bridge

```bash
npm install -g thunderbird-mcp
```

### 3. Configure the MCP Client

**Claude Code** (`~/.claude.json`):
```json
{
  "mcpServers": {
    "thunderbird": {
      "command": "node",
      "args": ["/usr/local/lib/node_modules/thunderbird-mcp/mcp-bridge.cjs"]
    }
  }
}
```

**VS Code / GitHub Copilot** (`.vscode/mcp.json` or User Settings):
```json
{
  "servers": {
    "thunderbird": {
      "type": "stdio",
      "command": "node",
      "args": ["/usr/local/lib/node_modules/thunderbird-mcp/mcp-bridge.cjs"]
    }
  }
}
```

### 4. Install This Skill Globally (optional, recommended)

So Copilot knows the Thunderbird capabilities across all projects:

```bash
mkdir -p ~/.github/skills
curl -L https://github.com/TKasperczyk/thunderbird-mcp/releases/latest/download/extension-usage-guide.md \
  -o ~/.github/skills/thunderbird-mcp.md
```

---

## 🗂️ Primary Use Cases: Archiving Emails & Attachments

These are the most common and important scenarios when using Thunderbird MCP.

---

### 📁 Use Case 1: Archive Emails as EML Files

**Typical user prompt:**
> "Save all emails from Alice in the folder `Documents/Correspondence/Alice`."

**What Copilot/Claude does internally:**

```
1. searchMessages({ query: "from:Alice" })
   → Finds all emails from Alice, returns messageId + folderPath per message

2. Per email: getMessage({ messageId, folderPath, rawSource: true })
   → Returns the full RFC 2822 source text (= EML content)

3. Generate filename: "<date>_<subject>.eml"
   e.g. "2026-05-10_Project-Update.eml"

4. Write file to: ~/Documents/Correspondence/Alice/
   → Directory is created if it doesn't exist

5. Output summary
```

**Example Copilot output:**
```
✅ 7 emails from Alice archived → ~/Documents/Correspondence/Alice/

  2026-05-15_Project-Update.eml          (12 KB)
  2026-05-10_Invoice_April.eml           (8 KB)
  2026-04-28_Meeting_Thursday.eml        (4 KB)
  2026-04-20_Offer_Webdesign.eml         (21 KB)
  ...

Summary:
- Main topics: project progress, invoices, scheduling
- Latest message: May 15, 2026 ("Project Update – Phase 2 complete")
- Attachments found: 3 PDFs, 1 Excel file (archive separately?)
```

**Key parameters for `getMessage`:**
```javascript
getMessage({
  messageId: "msg-123",
  folderPath: "/mail/Local%20Folders/Inbox",
  rawSource: true  // ← Full RFC 2822 source = EML content
})
// rawSource contains: all headers, body (text + HTML), all attachments,
//                     signatures, inline images — everything in one string
```

**EML filename convention (best practice):**
```
Format:  YYYY-MM-DD_<subject-sanitized>.eml
Example: 2026-05-15_Project-Update-Phase-2.eml

Rules:
- Use the date from the email's "Date" header
- Replace special characters in subject with "_": / \ : * ? " < > |
- Truncate subject to ~60 characters
- On conflict (same subject + date): append suffix _2, _3, ...
```

---

### 📎 Use Case 2: Archive Attachments from Emails

**Typical user prompt:**
> "Save all attachments related to car insurance for `AA-BB 123` in `Documents/Car/Porsche`."

**What Copilot/Claude does internally:**

```
1. searchMessages({ query: "AA-BB 123", searchBody: true })
   → Searches the full email body for the license plate

2. Per email: getMessage({ messageId, folderPath, saveAttachments: true })
   → Saves attachments to /tmp/thunderbird-mcp/<messageId>/
   → Returns filePath for each attachment

3. Copy attachments to target folder: ~/Documents/Car/Porsche/
   → Directory is created if it doesn't exist
   → Optionally prefix with email date: "2026-03-01_Insurance-Policy.pdf"

4. Output summary
```

**Example Copilot output:**
```
✅ 9 attachments from 4 emails archived → ~/Documents/Car/Porsche/

  2026-03-01_Insurance-Policy-2026.pdf      (245 KB)  [from: allianz@insurance.com]
  2026-03-01_Premium-Invoice-Q1.pdf         (89 KB)   [from: allianz@insurance.com]
  2026-01-15_Damage-Form.pdf                (512 KB)  [from: claims@allianz.com]
  2026-01-15_Accident-Report-Photos.zip     (3.2 MB)  [from: claims@allianz.com]
  2025-12-20_Green-Card.pdf                 (34 KB)   [from: allianz@insurance.com]
  ...

Skipped (no attachments): 2 emails
Total: 9 files, 4.1 MB
```

**Key parameters for `getMessage`:**
```javascript
getMessage({
  messageId: "msg-456",
  folderPath: "/mail/Local%20Folders/Inbox",
  saveAttachments: true  // ← Attachments saved to /tmp/thunderbird-mcp/<msgId>/
})
// Each attachment in the response contains:
// {
//   name: "Insurance-Policy-2026.pdf",
//   contentType: "application/pdf",
//   size: 250880,
//   filePath: "/tmp/thunderbird-mcp/msg-456/Insurance-Policy-2026.pdf"  ← local path
// }
// Then: copy/move file from filePath to the target directory
```

**Attachment filename convention (best practice):**
```
Format:  YYYY-MM-DD_<original-filename>
Example: 2026-03-01_Insurance-Policy-2026.pdf

Rules:
- Use email date as prefix (for sortability)
- Keep the original filename
- On conflict (same name): append suffix _2, _3, ...
- Empty names: use "attachment_<index>" as fallback
```

---

## 🔎 Advanced Search for Archiving

Various operators can be combined for more precise queries:

```javascript
// All emails from a domain
searchMessages({ query: "from:allianz.com" })

// Search subject and body
searchMessages({ query: "AA-BB 123", searchBody: true })

// Restrict to a time range
searchMessages({ query: "from:Alice", startDate: "2026-01-01", endDate: "2026-12-31" })

// Only with attachments (search by keyword, then filter by attachments.length > 0)
searchMessages({ query: "Invoice", maxResults: 200 })

// In a specific folder
searchMessages({ query: "Insurance", folderPath: "mailbox://user@host/INBOX" })

// Unread only
searchMessages({ query: "from:Alice", unreadOnly: true })
```

---

## 📋 Workflow Template: Batch Archiving

Copilot typically follows this pattern for archiving tasks:

```
1. SEARCH
   searchMessages({ query: ..., maxResults: 200 })
   → List of all relevant emails

2. CONFIRM (optional, when uncertain)
   Show a short preview and ask for confirmation before writing many files

3. EXPORT (per email)
   For EML archiving:    getMessage({ ..., rawSource: true })
   For attachments:      getMessage({ ..., saveAttachments: true })

4. SAVE
   Create target directory, write files with meaningful names

5. SUMMARIZE
   Report count, total size, date range, and any errors
```

---

## 🔍 All Available Tools (Overview)

| Group | Tool | Description |
|-------|------|-------------|
| **Search** | `searchMessages` | Full-text, sender, date, tags, body |
| **Search** | `getRecentMessages` | Recent messages with pagination |
| **Read** | `getMessage` | Full email, `rawSource`, `saveAttachments` |
| **Read** | `displayMessage` | Open email in Thunderbird |
| **Manage** | `updateMessage` | Read/flag/tag/move |
| **Manage** | `deleteMessages` | Delete messages |
| **Manage** | `createFolder` / `moveFolder` | Manage folders |
| **Compose** | `sendMail` | New email |
| **Compose** | `replyToMessage` / `forwardMessage` | Reply/forward |
| **Compose** | `saveDraft` | Save draft |
| **Filters** | `listFilters` / `createFilter` / `applyFilters` | Filter rules |
| **Contacts** | `searchContacts` / `createContact` | Address book |
| **Calendar** | `listCalendars` / `createEvent` / `listEvents` | Events |

---

## ⚙️ Configuration & Security

### Control Account Access
In Thunderbird: **Tools > Add-ons > Thunderbird MCP > Options**
- Which email accounts are accessible via MCP?
- Which tools are enabled?

### Send Safety
```
□ Enable "Block skipReview"
  → AI can only send emails with confirmation in the Compose window
  → Prevents accidental direct sending
```

---

## 🐛 Known Limitations

| Issue | Cause | Solution |
|-------|-------|----------|
| Body only ~200 chars | IMAP without offline sync | Use `searchBody: true` or enable offline sync |
| `rawSource` fails | Message not cached locally | Open it once in Thunderbird (forces download) |
| Attachment size limit | Max. 25 MB per file | Download larger attachments manually |
| `filePath` empty | Attachment has no URL | Occurs with embedded (CID) images |

---

## 📚 References

- [Thunderbird MCP GitHub](https://github.com/TKasperczyk/thunderbird-mcp)
- [MCP Protocol Spec](https://modelcontextprotocol.io/)
- [RFC 2822 – Internet Message Format](https://tools.ietf.org/html/rfc2822)
