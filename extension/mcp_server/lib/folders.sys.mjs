/**
 * Folder operations for the Thunderbird MCP server.
 *
 * Covers folder enumeration (listFolders), opening with DB prep
 * (openFolder), CRUD (createFolder / renameFolder / deleteFolder /
 * moveFolder), and bulk message wiping for special folders
 * (emptyTrash / emptyJunk).
 *
 * Factory pattern: api.js calls makeFolders({ ... }) once per start()
 * invocation and gets back the operational object. MailServices / Cc /
 * Ci / Services are sandbox globals in api.js but not in .sys.mjs
 * files, so they MUST be passed in. Access-control helpers
 * (getAccessibleAccounts, getAccessibleFolder, isAccountAllowed) are
 * supplied by the access-control module created earlier in start().
 *
 * Note: openFolder is a SUPERSET of access-control's
 * getAccessibleFolder. getAccessibleFolder validates URI lookup +
 * account-policy and returns { folder } | { error }. openFolder calls
 * getAccessibleFolder first, then performs an IMAP folder refresh
 * (best-effort) and grabs folder.msgDatabase, returning
 * { folder, db } | { error }. It is the right entry point for any
 * tool that will read message headers.
 */

export function makeFolders({
  MailServices,
  Cc,
  Ci,
  Services,
  getAccessibleAccounts,
  getAccessibleFolder,
  isAccountAllowed,
}) {
  // ── Module-private helpers ──

  /**
   * Find the Trash folder for the account that owns the given folder.
   * Walks the account tree looking for a folder with the TRASH flag
   * (0x00000100) set; falls back to a folder named "trash" or
   * "deleted items" if no flagged folder is found.
   */
  function findTrashFolder(folder) {
    const TRASH_FLAG = 0x00000100;
    let account = null;
    try {
      account = MailServices.accounts.findAccountForServer(folder.server);
    } catch {
      return null;
    }
    const root = account?.incomingServer?.rootFolder;
    if (!root) return null;

    let fallback = null;
    const TRASH_NAMES = ["trash", "deleted items"];
    const stack = [root];
    while (stack.length > 0) {
      const current = stack.pop();
      try {
        if (current && typeof current.getFlag === "function" && current.getFlag(TRASH_FLAG)) {
          return current;
        }
      } catch {}
      if (!fallback && current?.prettyName && TRASH_NAMES.includes(current.prettyName.toLowerCase())) {
        fallback = current;
      }
      try {
        if (current?.hasSubFolders) {
          for (const sf of current.subFolders) stack.push(sf);
        }
      } catch {}
    }
    return fallback;
  }

  /**
   * Find a special folder by flag bit, searching the account's folder tree.
   */
  function findSpecialFolder(root, flagBit) {
    const search = (folder) => {
      try {
        if (folder.getFlag && folder.getFlag(flagBit)) return folder;
      } catch {}
      if (folder.hasSubFolders) {
        for (const sub of folder.subFolders) {
          const found = search(sub);
          if (found) return found;
        }
      }
      return null;
    };
    return search(root);
  }

  /**
   * Recursively delete all messages in a folder and its subfolders.
   * Returns total count of messages deleted.
   */
  function deleteAllMessagesRecursive(folder) {
    let count = 0;
    try {
      const db = folder.msgDatabase;
      if (db) {
        const hdrs = [];
        for (const hdr of db.enumerateMessages()) hdrs.push(hdr);
        if (hdrs.length > 0) {
          folder.deleteMessages(hdrs, null, true, false, null, false);
          count += hdrs.length;
        }
      }
    } catch (e) {
      // Continue traversal; log per-folder failures so partial empties are visible.
      console.error("thunderbird-mcp: deleteMessages failed for folder", folder?.URI || folder?.name, ":", e);
    }
    if (folder.hasSubFolders) {
      for (const sub of folder.subFolders) {
        count += deleteAllMessagesRecursive(sub);
      }
    }
    return count;
  }

  // ── Public API ──

  /**
   * Lists all folders (optionally limited to a single account).
   * Depth is 0 for root children, increasing for subfolders.
   */
  function listFolders(accountId, folderPath) {
    const results = [];

    function folderType(flags) {
      if (flags & 0x00001000) return "inbox";
      if (flags & 0x00000200) return "sent";
      if (flags & 0x00000400) return "drafts";
      if (flags & 0x00000100) return "trash";
      if (flags & 0x00400000) return "templates";
      if (flags & 0x00000800) return "queue";
      if (flags & 0x40000000) return "junk";
      if (flags & 0x00004000) return "archive";
      return "folder";
    }

    function walkFolder(folder, accountKey, depth) {
      try {
        // Skip virtual/search folders to avoid duplicates
        if (folder.flags & 0x00000020) return;

        const prettyName = folder.prettyName;
        results.push({
          name: prettyName || folder.name || "(unnamed)",
          path: folder.URI,
          type: folderType(folder.flags),
          accountId: accountKey,
          totalMessages: folder.getTotalMessages(false),
          unreadMessages: folder.getNumUnread(false),
          depth
        });
      } catch {
        // Skip inaccessible folders
      }

      try {
        if (folder.hasSubFolders) {
          for (const subfolder of folder.subFolders) {
            walkFolder(subfolder, accountKey, depth + 1);
          }
        }
      } catch {
        // Skip subfolder traversal errors
      }
    }

    // folderPath filter: list that folder and its subtree
    if (folderPath) {
      const result = getAccessibleFolder(folderPath);
      if (result.error) return result;
      const folder = result.folder;
      const accountKey = folder.server
        ? (MailServices.accounts.findAccountForServer(folder.server)?.key || "unknown")
        : "unknown";
      walkFolder(folder, accountKey, 0);
      return results;
    }

    if (accountId) {
      if (!isAccountAllowed(accountId)) {
        return { error: `Account not accessible: ${accountId}` };
      }
      let target = null;
      for (const account of MailServices.accounts.accounts) {
        if (account.key === accountId) {
          target = account;
          break;
        }
      }
      if (!target) {
        return { error: `Account not found: ${accountId}` };
      }
      try {
        const root = target.incomingServer.rootFolder;
        if (root && root.hasSubFolders) {
          for (const subfolder of root.subFolders) {
            walkFolder(subfolder, target.key, 0);
          }
        }
      } catch {
        // Skip inaccessible account
      }
      return results;
    }

    for (const account of getAccessibleAccounts()) {
      try {
        const root = account.incomingServer.rootFolder;
        if (!root) continue;
        if (root.hasSubFolders) {
          for (const subfolder of root.subFolders) {
            walkFolder(subfolder, account.key, 0);
          }
        }
      } catch {
        // Skip inaccessible accounts/folders
      }
    }

    return results;
  }

  /**
   * Opens a folder and its message database.
   * Best-effort refresh for IMAP folders (db may be stale).
   *
   * Superset of getAccessibleFolder: also refreshes IMAP and resolves
   * folder.msgDatabase. Use this (not getAccessibleFolder) when the
   * caller needs to enumerate or look up message headers.
   *
   * Returns { folder, db } or { error }.
   */
  function openFolder(folderPath) {
    try {
      const result = getAccessibleFolder(folderPath);
      if (result.error) return result;
      const folder = result.folder;

      // Attempt to refresh IMAP folders. This is async and may not
      // complete before we read, but helps with stale data.
      if (folder.server && folder.server.type === "imap") {
        try {
          folder.updateFolder(null);
        } catch {
          // updateFolder may fail, continue anyway
        }
      }

      const db = folder.msgDatabase;
      if (!db) {
        return { error: "Could not access folder database" };
      }

      return { folder, db };
    } catch (e) {
      return { error: e.toString() };
    }
  }

  function createFolder(parentFolderPath, name) {
    try {
      if (typeof parentFolderPath !== "string" || !parentFolderPath) {
        return { error: "parentFolderPath must be a non-empty string" };
      }
      if (typeof name !== "string" || !name) {
        return { error: "name must be a non-empty string" };
      }

      const parentResult = getAccessibleFolder(parentFolderPath);
      if (parentResult.error) return parentResult;
      const parent = parentResult.folder;

      parent.createSubfolder(name, null);

      // Try to return the new folder's URI
      let newPath = null;
      try {
        if (parent.hasSubFolders) {
          for (const sub of parent.subFolders) {
            if (sub.prettyName === name || sub.name === name) {
              newPath = sub.URI;
              break;
            }
          }
        }
      } catch {
        // Folder may not be immediately visible (IMAP)
      }

      return {
        success: true,
        message: `Folder "${name}" created`,
        path: newPath
      };
    } catch (e) {
      const msg = e.toString();
      if (msg.includes("NS_MSG_FOLDER_EXISTS")) {
        return { error: `Folder "${name}" already exists under this parent` };
      }
      return { error: msg };
    }
  }

  function renameFolder(folderPath, newName) {
    try {
      if (typeof folderPath !== "string" || !folderPath) {
        return { error: "folderPath must be a non-empty string" };
      }
      if (typeof newName !== "string" || !newName) {
        return { error: "newName must be a non-empty string" };
      }

      const renameResult = getAccessibleFolder(folderPath);
      if (renameResult.error) return renameResult;
      const folder = renameResult.folder;

      folder.rename(newName, null);
      return {
        success: true,
        message: `Folder renamed to "${newName}"`,
        oldPath: folderPath,
      };
    } catch (e) {
      return { error: e.toString() };
    }
  }

  function deleteFolder(folderPath) {
    try {
      if (typeof folderPath !== "string" || !folderPath) {
        return { error: "folderPath must be a non-empty string" };
      }

      const delResult = getAccessibleFolder(folderPath);
      if (delResult.error) return delResult;
      const folder = delResult.folder;
      const folderName = folder.prettyName || folder.name || folderPath;

      const parent = folder.parent;
      if (!parent) {
        return { error: "Cannot delete a root folder" };
      }

      // Check if folder is already in Trash — if so, permanently delete
      const TRASH_FLAG = 0x00000100;
      let inTrash = false;
      let ancestor = folder;
      while (ancestor) {
        try {
          if (ancestor.getFlag && ancestor.getFlag(TRASH_FLAG)) {
            inTrash = true;
            break;
          }
        } catch { /* ignore */ }
        ancestor = ancestor.parent;
      }

      if (inTrash) {
        // Permanently delete — deleteSelf requires a msgWindow
        const win = Services.wm.getMostRecentWindow("mail:3pane");
        folder.deleteSelf(win?.msgWindow ?? null);
        return { success: true, message: `Folder "${folderName}" permanently deleted` };
      } else {
        // Move to trash
        const trashFolder = findTrashFolder(folder);
        if (!trashFolder) {
          return { error: "Trash folder not found" };
        }
        MailServices.copy.copyFolder(folder, trashFolder, true, null, null);
        return { success: true, message: `Folder "${folderName}" moved to Trash` };
      }
    } catch (e) {
      return { error: e.toString() };
    }
  }

  function emptyTrash(accountId) {
    try {
      const TRASH_FLAG = 0x00000100;
      const accounts = accountId
        ? [MailServices.accounts.getAccount(accountId)].filter(Boolean)
        : Array.from(getAccessibleAccounts());
      if (accountId && accounts.length === 0) {
        return { error: `Account not found: ${accountId}` };
      }
      if (accountId && !isAccountAllowed(accountId)) {
        return { error: `Account not accessible: ${accountId}` };
      }

      const results = [];
      for (const account of accounts) {
        const root = account.incomingServer?.rootFolder;
        if (!root) continue;
        const trash = findSpecialFolder(root, TRASH_FLAG);
        if (!trash) {
          results.push({ account: account.key, status: "no Trash folder found" });
          continue;
        }
        // Use Thunderbird's native emptyTrash when available (handles
        // IMAP expunge, subfolders, and compaction correctly)
        if (typeof trash.emptyTrash === "function") {
          const win = Services.wm.getMostRecentWindow("mail:3pane");
          trash.emptyTrash(win?.msgWindow ?? null, null);
          results.push({ account: account.key, folder: trash.URI, status: "emptied" });
        } else {
          // Fallback: manually delete messages in folder + subfolders
          const deleted = deleteAllMessagesRecursive(trash);
          results.push({ account: account.key, folder: trash.URI, deleted });
        }
      }
      return { success: true, results };
    } catch (e) {
      return { error: e.toString() };
    }
  }

  function emptyJunk(accountId) {
    try {
      const JUNK_FLAG = 0x40000000;
      const accounts = accountId
        ? [MailServices.accounts.getAccount(accountId)].filter(Boolean)
        : Array.from(getAccessibleAccounts());
      if (accountId && accounts.length === 0) {
        return { error: `Account not found: ${accountId}` };
      }
      if (accountId && !isAccountAllowed(accountId)) {
        return { error: `Account not accessible: ${accountId}` };
      }

      const results = [];
      for (const account of accounts) {
        const root = account.incomingServer?.rootFolder;
        if (!root) continue;
        const junk = findSpecialFolder(root, JUNK_FLAG);
        if (!junk) {
          results.push({ account: account.key, status: "no Junk folder found" });
          continue;
        }
        // No native emptyJunk in Thunderbird, delete recursively
        const deleted = deleteAllMessagesRecursive(junk);
        results.push({ account: account.key, folder: junk.URI, deleted });
      }
      return { success: true, results };
    } catch (e) {
      return { error: e.toString() };
    }
  }

  function moveFolder(folderPath, newParentPath) {
    try {
      if (typeof folderPath !== "string" || !folderPath) {
        return { error: "folderPath must be a non-empty string" };
      }
      if (typeof newParentPath !== "string" || !newParentPath) {
        return { error: "newParentPath must be a non-empty string" };
      }

      const srcResult = getAccessibleFolder(folderPath);
      if (srcResult.error) return srcResult;
      const folder = srcResult.folder;
      const folderName = folder.prettyName || folder.name || folderPath;

      const destResult = getAccessibleFolder(newParentPath);
      if (destResult.error) return destResult;
      const newParent = destResult.folder;
      const parentName = newParent.prettyName || newParent.name || newParentPath;

      if (folder.parent && folder.parent.URI === newParentPath) {
        return { error: "Folder is already under this parent" };
      }

      MailServices.copy.copyFolder(folder, newParent, true, null, null);
      return {
        success: true,
        message: `Folder "${folderName}" moved to "${parentName}"`,
      };
    } catch (e) {
      return { error: e.toString() };
    }
  }

  return {
    listFolders,
    openFolder,
    createFolder,
    renameFolder,
    deleteFolder,
    moveFolder,
    emptyTrash,
    emptyJunk,
  };
}
