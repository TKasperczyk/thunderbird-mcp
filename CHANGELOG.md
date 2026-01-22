# Changelog

## [Unreleased] - Enhanced Account and Folder Management

### Added
- **listAccounts** tool - Lists all configured email accounts with detailed information
  - Returns account key, name, type (IMAP/POP3/RSS), email address, username, and hostname
  - Useful for discovering available accounts before querying folders or messages
  
- **listFolders** tool - Lists all folders/mailboxes with filtering options
  - Recursively traverses folder hierarchy across all accounts or specific account
  - Returns folder name, path (URI), type (inbox/sent/drafts/trash/templates/folder)
  - Includes message counts (total and unread) for each folder
  - Properly detects folder types using Thunderbird's folder flags
  
- **getRecentMessages** tool - Efficiently retrieve recent messages with advanced filtering
  - Date-based filtering (default: last 30 days, configurable)
  - Returns messages sorted by date (newest first)
  - Optional folder-specific search or automatic inbox-wide search
  - Unread-only filter option
  - Configurable result limit (default: 20, max: 100 messages)
  - Automatically refreshes IMAP folders before reading
  - Uses proper MIME decoding for international characters

### Improved
- Added comprehensive JSDoc documentation to all new functions
  - Detailed parameter descriptions
  - Return value documentation
  - Implementation notes for folder flags and quirks
  
- Enhanced MCP bridge to properly handle resources/list and prompts/list requests
  - Returns empty arrays for unsupported resource/prompt types
  - Properly advertises capabilities in initialize response

### Technical Details
- New functions integrate seamlessly with existing MailServices API
- Folder type detection uses standard Thunderbird folder flags:
  - 0x00001000: Inbox
  - 0x00000200: Sent
  - 0x00000400: Drafts
  - 0x00000100: Trash
  - 0x00000800: Templates
- Skips virtual folders (search folders, unified inbox) to avoid duplicates
- Properly handles MIME-encoded headers using mime2Decoded* properties
- Efficient date-based filtering before sorting to minimize memory usage

### Performance
- getRecentMessages is significantly faster than searchMessages for recent mail
  - Direct date filtering avoids scanning entire mailbox
  - Sort-then-limit approach ensures newest messages are returned
  - No 50-message hard limit like searchMessages

## [0.1.0] - Initial Release (TKasperczyk)
- searchMessages - Find messages by text query
- getMessage - Read full email content
- sendMail - Compose new emails
- replyToMessage - Reply to messages
- searchContacts - Find contacts
- listCalendars - List calendars
