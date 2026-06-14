"use strict";

const { describe, it } = require("node:test");
const assert = require("node:assert/strict");

// Mirrors normalizeMessageIdForDedup from api.js.
function normalizeMessageIdForDedup(value) {
  if (value === undefined || value === null) return "";
  let normalized = String(value).trim();
  if (!normalized) return "";
  if (normalized.startsWith("<") && normalized.endsWith(">")) {
    normalized = normalized.slice(1, -1).trim();
  }
  return normalized;
}

// Mirrors dedupeSearchMessageResults from api.js.
function dedupeSearchMessageResults(results) {
  const seen = new Map();
  const deduped = [];

  function addDupLocation(survivor, folderPath) {
    if (!folderPath || folderPath === survivor.folderPath) return;
    if (!Array.isArray(survivor.dupLocations)) survivor.dupLocations = [];
    if (!survivor.dupLocations.includes(folderPath)) {
      survivor.dupLocations.push(folderPath);
    }
  }

  function mergeDupLocations(survivor, row) {
    addDupLocation(survivor, row.folderPath);
    if (Array.isArray(row.dupLocations)) {
      for (const folderPath of row.dupLocations) {
        addDupLocation(survivor, folderPath);
      }
    }
  }

  for (const row of results) {
    const normalizedId = normalizeMessageIdForDedup(row?.id);
    if (!normalizedId) {
      deduped.push(row);
      continue;
    }

    const survivor = seen.get(normalizedId);
    if (survivor) {
      mergeDupLocations(survivor, row);
      continue;
    }

    if (Array.isArray(row.dupLocations)) {
      const existingDupLocations = row.dupLocations;
      delete row.dupLocations;
      for (const folderPath of existingDupLocations) {
        addDupLocation(row, folderPath);
      }
    }
    seen.set(normalizedId, row);
    deduped.push(row);
  }

  return deduped;
}

function paginate(results, offset, effectiveLimit) {
  const effectiveOffset = offset > 0 ? Math.floor(offset) : 0;
  return {
    messages: results.slice(effectiveOffset, effectiveOffset + effectiveLimit),
    totalMatches: results.length,
    offset: effectiveOffset,
    limit: effectiveLimit,
    hasMore: effectiveOffset + effectiveLimit < results.length,
  };
}

describe("searchMessages dedup: Message-ID normalization", () => {
  it("strips one pair of angle brackets and trims whitespace, preserving case", () => {
    assert.equal(normalizeMessageIdForDedup("  <Case.ID@Example.COM>  "), "Case.ID@Example.COM");
    assert.equal(normalizeMessageIdForDedup("  Case.ID@Example.COM  "), "Case.ID@Example.COM");
    assert.equal(normalizeMessageIdForDedup("  < Case.ID@Example.COM >  "), "Case.ID@Example.COM");
  });

  it("keeps case-distinct Message-IDs separate (RFC 5322 local-part is case-sensitive)", () => {
    assert.notEqual(
      normalizeMessageIdForDedup("<ABC@example.com>"),
      normalizeMessageIdForDedup("<abc@example.com>")
    );
  });

  it("returns empty string for empty or missing values", () => {
    assert.equal(normalizeMessageIdForDedup(""), "");
    assert.equal(normalizeMessageIdForDedup("   "), "");
    assert.equal(normalizeMessageIdForDedup(null), "");
    assert.equal(normalizeMessageIdForDedup(undefined), "");
  });
});

describe("searchMessages dedup: result collapsing", () => {
  it("collapses the same normalized id across folders into the first survivor", () => {
    const inbox = { id: "<same@example.com>", folderPath: "imap://acct/INBOX", _dateTs: 30 };
    const important = { id: " same@example.com ", folderPath: "imap://acct/[Gmail]/Important", _dateTs: 30 };
    const allMail = { id: "same@example.com", folderPath: "imap://acct/[Gmail]/All Mail", _dateTs: 30 };

    const result = dedupeSearchMessageResults([inbox, important, allMail]);

    assert.equal(result.length, 1);
    assert.equal(result[0], inbox);
    assert.deepStrictEqual(result[0].dupLocations, [
      "imap://acct/[Gmail]/Important",
      "imap://acct/[Gmail]/All Mail",
    ]);
    assert.equal(result[0]._dateTs, 30);
  });

  it("leaves distinct ids alone without adding dupLocations", () => {
    const rows = [
      { id: "one@example.com", folderPath: "imap://acct/INBOX" },
      { id: "two@example.com", folderPath: "imap://acct/Archive" },
    ];

    const result = dedupeSearchMessageResults(rows);

    assert.deepStrictEqual(result, rows);
    assert.ok(!("dupLocations" in result[0]));
    assert.ok(!("dupLocations" in result[1]));
  });

  it("passes through rows with no Message-ID uncollapsed", () => {
    const rows = [
      { id: "", folderPath: "imap://acct/INBOX" },
      { id: null, folderPath: "imap://acct/[Gmail]/All Mail" },
      { folderPath: "imap://acct/Archive" },
    ];

    const result = dedupeSearchMessageResults(rows);

    assert.equal(result.length, 3);
    assert.deepStrictEqual(result, rows);
  });

  it("merges pre-existing dupLocations and excludes the survivor location", () => {
    const primary = {
      id: "merge@example.com",
      folderPath: "imap://acct/INBOX",
      dupLocations: ["imap://acct/Archive", "imap://acct/INBOX", "imap://acct/Archive"],
    };
    const duplicate = {
      id: "<merge@example.com>",
      folderPath: "imap://acct/[Gmail]/Important",
      dupLocations: ["imap://acct/Sent", "imap://acct/Archive"],
    };

    const result = dedupeSearchMessageResults([primary, duplicate]);

    assert.equal(result.length, 1);
    assert.deepStrictEqual(result[0].dupLocations, [
      "imap://acct/Archive",
      "imap://acct/[Gmail]/Important",
      "imap://acct/Sent",
    ]);
  });

  it("count reflects the deduped result length", () => {
    const result = dedupeSearchMessageResults([
      { id: "same@example.com", folderPath: "imap://acct/INBOX" },
      { id: "<same@example.com>", folderPath: "imap://acct/[Gmail]/All Mail" },
      { id: "other@example.com", folderPath: "imap://acct/Archive" },
      { id: "", folderPath: "imap://acct/NoId" },
    ]);

    assert.equal(result.length, 3);
  });
});

describe("searchMessages dedup: before pagination", () => {
  it("paginates and counts the deduped array, not the raw rows", () => {
    const rawRows = [
      { id: "same@example.com", folderPath: "imap://acct/INBOX" },
      { id: "<same@example.com>", folderPath: "imap://acct/[Gmail]/All Mail" },
      { id: "second@example.com", folderPath: "imap://acct/Archive" },
      { id: "third@example.com", folderPath: "imap://acct/Sent" },
    ];

    const finalResults = dedupeSearchMessageResults(rawRows);
    const firstPage = paginate(finalResults, 0, 10);
    const offsetPage = paginate(finalResults, 1, 1);

    assert.equal(firstPage.totalMatches, 3);
    assert.equal(firstPage.messages.length, 3);
    assert.deepStrictEqual(firstPage.messages.map(row => row.id), [
      "same@example.com",
      "second@example.com",
      "third@example.com",
    ]);
    assert.equal(offsetPage.totalMatches, 3);
    assert.deepStrictEqual(offsetPage.messages.map(row => row.id), ["second@example.com"]);
  });
});
