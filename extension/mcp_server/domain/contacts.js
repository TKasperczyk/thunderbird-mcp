/**
 * Contacts domain module for the Thunderbird MCP server.
 *
 * Implements the address-book tools: searchContacts, createContact,
 * updateContact, deleteContact, plus the findContactByUID helper.
 *
 * Loading
 * -------
 * api.js loads this file with Services.scriptloader.loadSubScript and a
 * { module: { exports: {} } } scope (the same mechanism it uses for
 * security_helpers.js), then calls the exported register(ctx) factory.
 *
 *     const scope = { module: { exports: {} } };
 *     Services.scriptloader.loadSubScript(
 *       "resource://thunderbird-mcp/mcp_server/domain/contacts.js?" + Date.now(),
 *       scope
 *     );
 *     scope.module.exports(ctx);
 *
 * register(ctx) destructures the shared dependencies it needs from ctx
 * (XPCOM globals, limit constants, and the infra helpers appendComposeAudit
 * and isContactWritesBlocked), defines the tool functions verbatim, and
 * assigns them back onto ctx so the dispatch switch in api.js can call them.
 *
 * The function bodies are unchanged from the original monolithic api.js: they
 * reference the same identifier names, now provided as locals destructured
 * from ctx, so behavior is identical.
 */
"use strict";

module.exports = function register(ctx) {
  const {
    Cc,
    Ci,
    MailServices,
    MAX_SEARCH_RESULTS_CAP,
    DEFAULT_MAX_RESULTS,
    appendComposeAudit,
    isContactWritesBlocked,
  } = ctx;

  function searchContacts(query, maxResults) {
    const results = [];
    const lowerQuery = query.toLowerCase();
    const requestedLimit = Number(maxResults);
    const limit = Number.isFinite(requestedLimit) && requestedLimit > 0
      ? Math.min(Math.floor(requestedLimit), MAX_SEARCH_RESULTS_CAP)
      : DEFAULT_MAX_RESULTS;
    let truncated = false;

    for (const book of MailServices.ab.directories) {
      for (const card of book.childCards) {
        if (card.isMailList) continue;

        const email = (card.primaryEmail || "").toLowerCase();
        const displayName = (card.displayName || "").toLowerCase();
        const firstName = (card.firstName || "").toLowerCase();
        const lastName = (card.lastName || "").toLowerCase();

        if (email.includes(lowerQuery) ||
            displayName.includes(lowerQuery) ||
            firstName.includes(lowerQuery) ||
            lastName.includes(lowerQuery)) {
          results.push({
            id: card.UID,
            displayName: card.displayName,
            email: card.primaryEmail,
            firstName: card.firstName,
            lastName: card.lastName,
            addressBook: book.dirName,
            addressBookId: book.URI,
          });
        }

        if (results.length >= limit) { truncated = true; break; }
      }
      if (truncated) break;
    }

    if (truncated) {
      return { contacts: results, hasMore: true, message: `Results limited to ${limit}. Refine your query to see more.` };
    }
    return results;
  }

  /**
   * Find a contact card by UID across all address books.
   * Returns { card, book } or { error }.
   */
  function findContactByUID(contactId) {
    for (const book of MailServices.ab.directories) {
      for (const card of book.childCards) {
        if (card.UID === contactId) {
          return { card, book };
        }
      }
    }
    return { error: `Contact not found: ${contactId}` };
  }

  function createContact(email, displayName, firstName, lastName, addressBookId) {
    try {
      if (isContactWritesBlocked()) {
        return { error: "User preference blocks contact writes via MCP. Enable 'Allow contact writes' in the extension options page if you trust this MCP client to manage your address book." };
      }
      if (typeof email !== "string" || !email) {
        return { error: "email must be a non-empty string" };
      }
      appendComposeAudit({
        tool: "createContact",
        email: typeof email === "string" ? email.slice(0, 200) : null,
        addressBookId: typeof addressBookId === "string" ? addressBookId : null,
      });

      // Find the target address book
      let targetBook = null;
      if (addressBookId) {
        for (const book of MailServices.ab.directories) {
          if (book.dirPrefId === addressBookId || book.UID === addressBookId || book.URI === addressBookId) {
            targetBook = book;
            break;
          }
        }
        if (!targetBook) {
          return { error: `Address book not found: ${addressBookId}` };
        }
      } else {
        // Use the first writable address book
        for (const book of MailServices.ab.directories) {
          if (!book.readOnly) {
            targetBook = book;
            break;
          }
        }
        if (!targetBook) {
          return { error: "No writable address book found" };
        }
      }

      const card = Cc["@mozilla.org/addressbook/cardproperty;1"]
        .createInstance(Ci.nsIAbCard);
      card.primaryEmail = email;
      if (displayName) card.displayName = displayName;
      if (firstName) card.firstName = firstName;
      if (lastName) card.lastName = lastName;

      const newCard = targetBook.addCard(card);
      return {
        success: true,
        id: newCard.UID,
        email: newCard.primaryEmail,
        displayName: newCard.displayName,
        addressBook: targetBook.dirName,
      };
    } catch (e) {
      return { error: e.toString() };
    }
  }

  function updateContact(contactId, email, displayName, firstName, lastName) {
    try {
      if (isContactWritesBlocked()) {
        return { error: "User preference blocks contact writes via MCP. Enable 'Allow contact writes' in the extension options page if you trust this MCP client to manage your address book." };
      }
      if (typeof contactId !== "string" || !contactId) {
        return { error: "contactId must be a non-empty string" };
      }

      const found = findContactByUID(contactId);
      if (found.error) return found;
      const { card, book } = found;

      // Log the change BEFORE mutating so the audit captures the old
      // email -- crucial when an attacker repoints "Boss" at their
      // own address. Persist as much identifying detail as we can
      // without storing the full contact card.
      appendComposeAudit({
        tool: "updateContact",
        contactId,
        bookName: book.dirName,
        bookURI: book.URI,
        oldEmail: card.primaryEmail || null,
        newEmail: typeof email === "string" ? email.slice(0, 200) : null,
      });

      if (email !== undefined) card.primaryEmail = email;
      if (displayName !== undefined) card.displayName = displayName;
      if (firstName !== undefined) card.firstName = firstName;
      if (lastName !== undefined) card.lastName = lastName;

      book.modifyCard(card);
      return {
        success: true,
        id: card.UID,
        email: card.primaryEmail,
        displayName: card.displayName,
        firstName: card.firstName,
        lastName: card.lastName,
      };
    } catch (e) {
      return { error: e.toString() };
    }
  }

  function deleteContact(contactId) {
    try {
      if (isContactWritesBlocked()) {
        return { error: "User preference blocks contact writes via MCP. Enable 'Allow contact writes' in the extension options page if you trust this MCP client to manage your address book." };
      }
      if (typeof contactId !== "string" || !contactId) {
        return { error: "contactId must be a non-empty string" };
      }

      const found = findContactByUID(contactId);
      if (found.error) return found;
      const { card, book } = found;

      appendComposeAudit({
        tool: "deleteContact",
        contactId,
        bookName: book.dirName,
        bookURI: book.URI,
        email: card.primaryEmail || null,
        displayName: card.displayName || null,
      });

      book.deleteCards([card]);
      return {
        success: true,
        message: `Contact "${card.displayName || card.primaryEmail}" deleted`,
      };
    } catch (e) {
      return { error: e.toString() };
    }
  }

  Object.assign(ctx, {
    searchContacts,
    findContactByUID,
    createContact,
    updateContact,
    deleteContact,
  });
};
