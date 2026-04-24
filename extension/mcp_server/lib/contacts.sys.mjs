/**
 * Address-book / contact CRUD for the Thunderbird MCP server.
 *
 * Wraps MailServices.ab directory iteration plus card create/modify/delete
 * for the searchContacts / createContact / updateContact / deleteContact
 * tool handlers.
 *
 * Factory: api.js calls makeContacts({ MailServices, Cc, Ci,
 * MAX_SEARCH_RESULTS_CAP, DEFAULT_MAX_RESULTS }). MailServices / Cc / Ci
 * are sandbox globals in api.js but not in .sys.mjs files, so they MUST
 * be passed in. The result-cap constants live alongside the other shared
 * limits in api.js and are passed through for parity with the original
 * inline behavior.
 *
 * findContactByUID is kept module-private -- it is only consumed by
 * updateContact and deleteContact in this group and was not exposed at
 * the api.js call sites.
 */

export function makeContacts({
  MailServices,
  Cc,
  Ci,
  MAX_SEARCH_RESULTS_CAP,
  DEFAULT_MAX_RESULTS,
}) {
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
      if (typeof email !== "string" || !email) {
        return { error: "email must be a non-empty string" };
      }

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
      if (typeof contactId !== "string" || !contactId) {
        return { error: "contactId must be a non-empty string" };
      }

      const found = findContactByUID(contactId);
      if (found.error) return found;
      const { card, book } = found;

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
      if (typeof contactId !== "string" || !contactId) {
        return { error: "contactId must be a non-empty string" };
      }

      const found = findContactByUID(contactId);
      if (found.error) return found;
      const { card, book } = found;

      book.deleteCards([card]);
      return {
        success: true,
        message: `Contact "${card.displayName || card.primaryEmail}" deleted`,
      };
    } catch (e) {
      return { error: e.toString() };
    }
  }

  return { searchContacts, createContact, updateContact, deleteContact };
}
