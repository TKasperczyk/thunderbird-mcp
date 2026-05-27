"use strict";

const { describe, it } = require("node:test");
const assert = require("node:assert/strict");

function isEmptyTrashArgError(e) {
  return (e && (e.result === 0x80570001 || e.result === 0x80570009)) ||
    String(e).includes("Not enough arguments") ||
    String(e).includes("Could not convert JavaScript argument");
}

function callEmptyTrashCompat(trash, getWindow) {
  try {
    trash.emptyTrash(null);
  } catch (e) {
    if (isEmptyTrashArgError(e)) {
      const win = getWindow();
      trash.emptyTrash(win?.msgWindow ?? null, null);
    } else {
      throw e;
    }
  }
}

describe("emptyTrash signature compatibility", () => {
  it("uses the modern listener-only signature first", () => {
    const calls = [];
    const trash = {
      emptyTrash(...args) {
        calls.push(args);
      },
    };

    callEmptyTrashCompat(trash, () => {
      throw new Error("legacy window should not be requested");
    });

    assert.deepStrictEqual(calls, [[null]]);
  });

  it("falls back to the legacy msgWindow signature for bad-convert errors", () => {
    const calls = [];
    const msgWindow = {};
    const trash = {
      emptyTrash(...args) {
        calls.push(args);
        if (calls.length === 1) {
          throw { result: 0x80570009 };
        }
      },
    };

    callEmptyTrashCompat(trash, () => ({ msgWindow }));

    assert.deepStrictEqual(calls, [[null], [msgWindow, null]]);
  });

  it("falls back to the legacy msgWindow signature for not-enough-args errors", () => {
    const calls = [];
    const trash = {
      emptyTrash(...args) {
        calls.push(args);
        if (calls.length === 1) {
          throw new Error("Not enough arguments");
        }
      },
    };

    callEmptyTrashCompat(trash, () => null);

    assert.deepStrictEqual(calls, [[null], [null, null]]);
  });

  it("surfaces non-signature errors from the modern call", () => {
    const error = new Error("IMAP failed");
    const trash = {
      emptyTrash() {
        throw error;
      },
    };

    assert.throws(
      () => callEmptyTrashCompat(trash, () => ({ msgWindow: {} })),
      error
    );
  });

  it("surfaces non-signature errors from the legacy fallback", () => {
    const error = new Error("legacy failed");
    let calls = 0;
    const trash = {
      emptyTrash() {
        calls += 1;
        if (calls === 1) {
          throw new Error("Could not convert JavaScript argument");
        }
        throw error;
      },
    };

    assert.throws(
      () => callEmptyTrashCompat(trash, () => ({ msgWindow: {} })),
      error
    );
    assert.equal(calls, 2);
  });
});
