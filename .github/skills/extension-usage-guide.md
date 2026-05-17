---
name: "thunderbird-mcp-extension-usage"
description: "Guide for using Thunderbird MCP with AI assistants to archive emails and attachments to the local file system"
version: "1.1.0"
tags: ["thunderbird-mcp", "mcp", "email", "ai", "copilot", "attachments", "archiving"]
applyTo: []
---

# Thunderbird MCP Extension - Usage Guide

Die Thunderbird MCP Extension ermöglicht es AI-Assistenten (Claude, Copilot, etc.) direkt mit Thunderbird zu interagieren — E-Mails suchen, lesen, archivieren, Anhänge exportieren und vieles mehr.

## 📋 Architektur

```
AI Assistant (Claude/Copilot) 
    ↓ MCP Protocol (stdio)
mcp-bridge.cjs (Node.js)
    ↓ HTTP localhost:8765
Thunderbird Extension
    ↓
E-Mail Datenbank
```

## 🚀 Setup

### 1. Extension installieren

Die aktuelle `thunderbird-mcp.xpi` aus den [GitHub Releases](https://github.com/TKasperczyk/thunderbird-mcp/releases/latest) herunterladen und in Thunderbird installieren:

> **Thunderbird → Tools → Add-ons → ⚙️ → Add-on aus Datei installieren**

Danach Thunderbird neu starten.

### 2. MCP-Bridge installieren

```bash
npm install -g thunderbird-mcp
```

Oder alternativ nur die Bridge herunterladen:
```bash
# Einzelne Bridge-Datei aus dem Release herunterladen
curl -L https://github.com/TKasperczyk/thunderbird-mcp/releases/latest/download/mcp-bridge.cjs \
  -o ~/bin/thunderbird-mcp-bridge.cjs
```

### 3. MCP-Client konfigurieren

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

**VS Code / GitHub Copilot** (`.vscode/mcp.json` oder User-Settings):
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

### 4. Diesen Skill global ablegen (optional, empfohlen)

Damit Copilot die Thunderbird-Fähigkeiten in allen Projekten kennt:

```bash
mkdir -p ~/.github/skills
curl -L https://github.com/TKasperczyk/thunderbird-mcp/releases/latest/download/extension-usage-guide.md \
  -o ~/.github/skills/thunderbird-mcp.md
```

Oder die Datei manuell nach `~/.github/skills/thunderbird-mcp.md` kopieren.

---

## 🗂️ Primäre Use-Cases: E-Mails & Anhänge archivieren

Dies sind die häufigsten und wichtigsten Szenarien beim Einsatz von Thunderbird MCP.

---

### 📁 Use-Case 1: E-Mails als EML-Dateien archivieren

**Typischer Nutzer-Prompt:**
> „Lege alle E-Mails von Alice im Ordner `Dokumente/Korrespondenz/Alice` ab."

**Was Copilot/Claude intern tut:**

```
1. searchMessages({ query: "from:Alice" })
   → Findet alle E-Mails von Alice, liefert messageId + folderPath pro Nachricht

2. Pro E-Mail: getMessage({ messageId, folderPath, rawSource: true })
   → Liefert den vollständigen RFC 2822 Quelltext (= EML-Inhalt)

3. Dateiname generieren: "<Datum>_<Betreff>.eml"
   z.B. "2026-05-10_Projektupdate.eml"

4. Datei schreiben nach: ~/Dokumente/Korrespondenz/Alice/
   → Verzeichnis wird bei Bedarf angelegt

5. Zusammenfassung ausgeben
```

**Beispiel-Output von Copilot:**
```
✅ 7 E-Mails von Alice archiviert → ~/Dokumente/Korrespondenz/Alice/

  2026-05-15_Projektupdate.eml          (12 KB)
  2026-05-10_Rechnung_April.eml         (8 KB)
  2026-04-28_Meeting_Donnerstag.eml     (4 KB)
  2026-04-20_Angebot_Webdesign.eml      (21 KB)
  ...

Zusammenfassung der E-Mails:
- Hauptthemen: Projektfortschritt, Rechnungen, Terminabstimmungen
- Letzte Nachricht: 15. Mai 2026 ("Projektupdate – Phase 2 abgeschlossen")
- Anhänge enthalten: 3 PDFs, 1 Excel-Datei (separat archivieren?)
```

**Wichtige Parameter für `getMessage`:**
```javascript
getMessage({
  messageId: "msg-123",
  folderPath: "/mail/Local%20Folders/Inbox",
  rawSource: true  // ← Vollständiger RFC 2822 Source = EML-Inhalt
})
// rawSource enthält: alle Header, Body (Text + HTML), alle Anhänge,
//                    Signaturen, Inline-Bilder — alles in einem String
```

**EML-Dateiname generieren (Best Practice):**
```
Format:  YYYY-MM-DD_<Betreff-sanitized>.eml
Beispiel: 2026-05-15_Projektupdate_Phase_2.eml

Regeln:
- Datum aus dem "Date"-Header der E-Mail verwenden
- Sonderzeichen im Betreff durch "_" ersetzen: / \ : * ? " < > |
- Betreff auf ~60 Zeichen kürzen
- Bei Konflikten (gleicher Betreff+Datum): Suffix _2, _3, ... anhängen
```

---

### 📎 Use-Case 2: Anhänge aus E-Mails archivieren

**Typischer Nutzer-Prompt:**
> „Lege alle Anhänge der Autoversicherung für `AA-BB 123` im Ordner `Dokumente/Auto/Porsche` ab."

**Was Copilot/Claude intern tut:**

```
1. searchMessages({ query: "AA-BB 123", searchBody: true })
   → Sucht im gesamten E-Mail-Text nach dem Kennzeichen

2. Pro E-Mail: getMessage({ messageId, folderPath, saveAttachments: true })
   → Speichert Anhänge nach /tmp/thunderbird-mcp/<messageId>/
   → Gibt filePath für jeden Anhang zurück

3. Anhänge in den Zielordner kopieren: ~/Dokumente/Auto/Porsche/
   → Verzeichnis wird bei Bedarf angelegt
   → Ggf. E-Mail-Datum als Präfix: "2026-03-01_KFZ-Police.pdf"

4. Zusammenfassung ausgeben
```

**Beispiel-Output von Copilot:**
```
✅ 9 Anhänge aus 4 E-Mails archiviert → ~/Dokumente/Auto/Porsche/

  2026-03-01_KFZ-Police_2026.pdf            (245 KB)  [von: allianz@versicherung.de]
  2026-03-01_Beitragsrechnung_Q1.pdf        (89 KB)   [von: allianz@versicherung.de]
  2026-01-15_Schadensformular.pdf           (512 KB)  [von: schaden@allianz.de]
  2026-01-15_Unfallbericht_Fotos.zip        (3.2 MB)  [von: schaden@allianz.de]
  2025-12-20_Gruene_Karte.pdf               (34 KB)   [von: allianz@versicherung.de]
  ...

Übersprungen (keine Anhänge): 2 E-Mails
Gesamt: 9 Dateien, 4.1 MB
```

**Wichtige Parameter für `getMessage`:**
```javascript
getMessage({
  messageId: "msg-456",
  folderPath: "/mail/Local%20Folders/Inbox",
  saveAttachments: true  // ← Anhänge werden nach /tmp/thunderbird-mcp/<msgId>/ gespeichert
})
// Jeder Anhang in der Response enthält:
// {
//   name: "KFZ-Police_2026.pdf",
//   contentType: "application/pdf",
//   size: 250880,
//   filePath: "/tmp/thunderbird-mcp/msg-456/KFZ-Police_2026.pdf"  ← Lokaler Pfad
// }
// Danach: Datei von filePath in den Zielordner verschieben/kopieren
```

**Anhang-Dateiname generieren (Best Practice):**
```
Format:  YYYY-MM-DD_<original-filename>
Beispiel: 2026-03-01_KFZ-Police_2026.pdf

Regeln:
- Datum der E-Mail als Präfix verwenden (Sortierbarkeit)
- Originalen Dateinamen beibehalten
- Bei Konflikten (gleicher Name): Suffix _2, _3, ... anhängen
- Leere Namen: "attachment_<index>" als Fallback
```

---

## 🔎 Erweiterte Suche für Archivierungen

Für präzisere Suchanfragen können verschiedene Operatoren kombiniert werden:

```javascript
// Alle E-Mails von einer Domäne
searchMessages({ query: "from:allianz.de" })

// Betreff + Körper durchsuchen
searchMessages({ query: "AA-BB 123", searchBody: true })

// Zeitraum einschränken
searchMessages({ query: "from:Alice", startDate: "2026-01-01", endDate: "2026-12-31" })

// Nur mit Anhängen (nach Stichwort suchen, dann auf hasAttachments prüfen)
searchMessages({ query: "Rechnung", maxResults: 200 })
// → Im Response: attachments.length > 0 filtern

// In einem bestimmten Ordner
searchMessages({ query: "Versicherung", folderPath: "mailbox://user@host/INBOX" })

// Nur ungelesene
searchMessages({ query: "from:Alice", unreadOnly: true })
```

---

## 📋 Workflow-Vorlage: Batch-Archivierung

Copilot folgt bei Archivierungsaufgaben typischerweise diesem Muster:

```
1. SUCHEN
   searchMessages({ query: ..., maxResults: 200 })
   → Liste aller relevanten E-Mails

2. PRÜFEN (optional, bei Unsicherheit)
   Kurze Vorschau anzeigen und Bestätigung einholen, bevor viele Dateien gespeichert werden

3. EXPORTIEREN (pro E-Mail)
   Für EML-Archivierung:   getMessage({ ..., rawSource: true })
   Für Anhänge:            getMessage({ ..., saveAttachments: true })

4. ABLEGEN
   Zielverzeichnis anlegen, Dateien mit sinnvollem Namen speichern

5. ZUSAMMENFASSEN
   Anzahl, Gesamtgröße, Zeitraum, eventuelle Fehler berichten
```

---

## 🔍 Alle verfügbaren Tools (Überblick)

| Gruppe | Tool | Beschreibung |
|--------|------|-------------|
| **Suche** | `searchMessages` | Volltext, Absender, Datum, Tags, Body |
| **Suche** | `getRecentMessages` | Letzte Nachrichten mit Paginierung |
| **Lesen** | `getMessage` | Vollständige E-Mail, `rawSource`, `saveAttachments` |
| **Lesen** | `displayMessage` | E-Mail in Thunderbird öffnen |
| **Verwaltung** | `updateMessage` | Lesen/Flaggen/Tags/Verschieben |
| **Verwaltung** | `deleteMessages` | Nachrichten löschen |
| **Verwaltung** | `createFolder` / `moveFolder` | Ordner verwalten |
| **Verfassen** | `sendMail` | Neue E-Mail |
| **Verfassen** | `replyToMessage` / `forwardMessage` | Antworten/Weiterleiten |
| **Verfassen** | `saveDraft` | Entwurf speichern |
| **Filter** | `listFilters` / `createFilter` / `applyFilters` | Filterregeln |
| **Kontakte** | `searchContacts` / `createContact` | Adressbuch |
| **Kalender** | `listCalendars` / `createEvent` / `listEvents` | Termine |

---

## ⚙️ Konfiguration & Sicherheit

### Konto-Zugriff steuern
In Thunderbird: **Tools > Add-ons > Thunderbird MCP > Options**
- Welche E-Mail-Konten sind für MCP zugänglich?
- Welche Tools sind aktiviert?

### Versand-Sicherheit
```
□ "Block skipReview" aktivieren
  → AI kann E-Mails nur mit Bestätigung im Compose-Fenster versenden
  → Verhindert versehentliches direktes Versenden
```

---

## 🐛 Bekannte Einschränkungen

| Problem | Ursache | Lösung |
|---------|---------|--------|
| Body nur ~200 Zeichen | IMAP ohne Offline-Sync | `searchBody: true` oder Offline-Sync aktivieren |
| `rawSource` schlägt fehl | Nachricht nicht lokal gecacht | In Thunderbird einmal öffnen (erzwingt Download) |
| Anhang-Limit | Max. 25 MB pro Datei | Größere Anhänge manuell herunterladen |
| `filePath` leer | Anhang hat keine URL | Tritt bei eingebetteten (CID) Bildern auf |

---

## 📚 Referenzen

- [Thunderbird MCP GitHub](https://github.com/TKasperczyk/thunderbird-mcp)
- [MCP Protocol Spec](https://modelcontextprotocol.io/)
- [RFC 2822 – Internet Message Format](https://tools.ietf.org/html/rfc2822)
