/**
 * Stress tests and edge-case coverage for new tool features.
 *
 * These tests go beyond happy-path validation to exercise adversarial
 * and boundary inputs for:
 * - Tagging: addTags/removeTags arrays, prototype pollution, special chars
 * - Folder management: null/wrong-type params, long names, special chars
 * - Attachment sending: wrong types, mixed formats, large arrays
 * - Contact write: null/wrong-type params, prototype pollution, optional fields
 */

const { describe, it } = require('node:test');
const assert = require('node:assert/strict');

// ─── Helpers ───────────────────────────────────────────────────────

/**
 * Replicate validateToolArgs from api.js.
 */
function createValidator(tools) {
  const toolSchemas = Object.create(null);
  for (const t of tools) { toolSchemas[t.name] = t.inputSchema; }

  return function validateToolArgs(name, args) {
    const schema = toolSchemas[name];
    if (!schema) return [`Unknown tool: ${name}`];
    const errors = [];
    const props = schema.properties || {};
    const required = schema.required || [];

    for (const key of required) {
      if (args[key] === undefined || args[key] === null) {
        errors.push(`Missing required parameter: ${key}`);
      }
    }
    for (const [key, value] of Object.entries(args)) {
      const propSchema = Object.prototype.hasOwnProperty.call(props, key) ? props[key] : undefined;
      if (!propSchema) { errors.push(`Unknown parameter: ${key}`); continue; }
      if (value === undefined || value === null) continue;
      const expectedType = propSchema.type;
      if (expectedType === "array") {
        if (!Array.isArray(value)) errors.push(`Parameter '${key}' must be an array, got ${typeof value}`);
      } else if (expectedType === "object") {
        if (typeof value !== "object" || Array.isArray(value))
          errors.push(`Parameter '${key}' must be an object, got ${Array.isArray(value) ? "array" : typeof value}`);
      } else if (expectedType && typeof value !== expectedType) {
        errors.push(`Parameter '${key}' must be ${expectedType}, got ${typeof value}`);
      }
    }
    return errors;
  };
}

// ─── Tagging stress tests ─────────────────────────────────────────

describe('Tagging: validation edge cases', () => {
  const tagTools = [
    {
      name: "updateMessage",
      inputSchema: {
        type: "object",
        properties: {
          messageId: { type: "string" },
          messageIds: { type: "array", items: { type: "string" } },
          folderPath: { type: "string" },
          read: { type: "boolean" },
          flagged: { type: "boolean" },
          addTags: { type: "array", items: { type: "string" } },
          removeTags: { type: "array", items: { type: "string" } },
          moveTo: { type: "string" },
          trash: { type: "boolean" },
        },
        required: ["folderPath"],
      },
    },
    {
      name: "searchMessages",
      inputSchema: {
        type: "object",
        properties: {
          query: { type: "string" },
          tag: { type: "string" },
          maxResults: { type: "number" },
        },
        required: ["query"],
      },
    },
  ];
  const tagValidate = createValidator(tagTools);

  it('rejects addTags as object (not array)', () => {
    const errors = tagValidate('updateMessage', {
      folderPath: '/INBOX', addTags: { tag: '$label1' }
    });
    assert.equal(errors.length, 1);
    assert.match(errors[0], /must be an array/);
  });

  it('rejects addTags as boolean', () => {
    const errors = tagValidate('updateMessage', {
      folderPath: '/INBOX', addTags: true
    });
    assert.equal(errors.length, 1);
    assert.match(errors[0], /must be an array/);
  });

  it('rejects removeTags as string', () => {
    const errors = tagValidate('updateMessage', {
      folderPath: '/INBOX', removeTags: '$label1'
    });
    assert.equal(errors.length, 1);
    assert.match(errors[0], /must be an array/);
  });

  it('accepts empty addTags array', () => {
    const errors = tagValidate('updateMessage', {
      folderPath: '/INBOX', addTags: []
    });
    assert.equal(errors.length, 0);
  });

  it('accepts empty removeTags array', () => {
    const errors = tagValidate('updateMessage', {
      folderPath: '/INBOX', removeTags: []
    });
    assert.equal(errors.length, 0);
  });

  it('accepts addTags and removeTags simultaneously', () => {
    const errors = tagValidate('updateMessage', {
      folderPath: '/INBOX',
      messageId: 'msg1',
      addTags: ['$label1', '$label2'],
      removeTags: ['$label3'],
    });
    assert.equal(errors.length, 0);
  });

  it('accepts large number of tags in addTags', () => {
    const tags = Array.from({ length: 100 }, (_, i) => `custom-tag-${i}`);
    const errors = tagValidate('updateMessage', {
      folderPath: '/INBOX', addTags: tags
    });
    assert.equal(errors.length, 0);
  });

  it('rejects tag filter as array in searchMessages', () => {
    const errors = tagValidate('searchMessages', {
      query: 'test', tag: ['$label1']
    });
    assert.equal(errors.length, 1);
    assert.match(errors[0], /must be string/);
  });

  it('rejects tag filter as boolean in searchMessages', () => {
    const errors = tagValidate('searchMessages', {
      query: 'test', tag: true
    });
    assert.equal(errors.length, 1);
    assert.match(errors[0], /must be string/);
  });

  it('accepts empty string tag filter', () => {
    const errors = tagValidate('searchMessages', {
      query: 'test', tag: ''
    });
    assert.equal(errors.length, 0);
  });

  it('handles tags with special characters in validation', () => {
    const errors = tagValidate('updateMessage', {
      folderPath: '/INBOX',
      addTags: ['tag with spaces', '$label/slash', 'tag\twith\ttabs'],
    });
    assert.equal(errors.length, 0);
  });

  it('handles prototype pollution attempt via addTags', () => {
    const malicious = Object.create(null);
    malicious.folderPath = '/INBOX';
    malicious.addTags = ['$label1'];
    malicious.constructor = 'attack';
    const errors = tagValidate('updateMessage', malicious);
    assert.ok(errors.some(e => e.includes('Unknown parameter: constructor')));
  });
});

// ─── Folder management stress tests ──────────────────────────────

describe('Folder management: validation edge cases', () => {
  const folderTools = [
    {
      name: "renameFolder",
      inputSchema: {
        type: "object",
        properties: { folderPath: { type: "string" }, newName: { type: "string" } },
        required: ["folderPath", "newName"],
      },
    },
    {
      name: "deleteFolder",
      inputSchema: {
        type: "object",
        properties: { folderPath: { type: "string" } },
        required: ["folderPath"],
      },
    },
    {
      name: "moveFolder",
      inputSchema: {
        type: "object",
        properties: { folderPath: { type: "string" }, newParentPath: { type: "string" } },
        required: ["folderPath", "newParentPath"],
      },
    },
  ];
  const folderValidate = createValidator(folderTools);

  it('rejects renameFolder with null folderPath', () => {
    const errors = folderValidate('renameFolder', { folderPath: null, newName: 'test' });
    assert.ok(errors.some(e => e.includes('folderPath')));
  });

  it('rejects renameFolder with array newName', () => {
    const errors = folderValidate('renameFolder', { folderPath: '/INBOX', newName: ['a'] });
    assert.equal(errors.length, 1);
    assert.match(errors[0], /must be string/);
  });

  it('rejects deleteFolder with number folderPath', () => {
    const errors = folderValidate('deleteFolder', { folderPath: 12345 });
    assert.equal(errors.length, 1);
    assert.match(errors[0], /must be string/);
  });

  it('rejects moveFolder with boolean paths', () => {
    const errors = folderValidate('moveFolder', { folderPath: true, newParentPath: false });
    assert.equal(errors.length, 2);
  });

  it('rejects unknown params on renameFolder', () => {
    const errors = folderValidate('renameFolder', {
      folderPath: '/INBOX', newName: 'test', force: true
    });
    assert.equal(errors.length, 1);
    assert.match(errors[0], /Unknown parameter: force/);
  });

  it('handles very long folder name in renameFolder', () => {
    const longName = 'a'.repeat(10000);
    const errors = folderValidate('renameFolder', { folderPath: '/INBOX', newName: longName });
    assert.equal(errors.length, 0); // Validation passes; filesystem will reject
  });

  it('handles special characters in folder paths', () => {
    const errors = folderValidate('moveFolder', {
      folderPath: 'imap://user@server/INBOX/Über Spëcial & "Quotes"',
      newParentPath: 'imap://user@server/Archive/日本語'
    });
    assert.equal(errors.length, 0);
  });
});

// ─── Attachment sending stress tests ────────────────────────────

describe('Attachment sending: validation edge cases', () => {
  const mailTools = [
    {
      name: "sendMail",
      inputSchema: {
        type: "object",
        properties: {
          to: { type: "string" },
          subject: { type: "string" },
          body: { type: "string" },
          attachments: { type: "array" },
        },
        required: ["to", "subject", "body"],
      },
    },
  ];
  const mailValidate = createValidator(mailTools);

  it('rejects attachments as object (not array)', () => {
    const errors = mailValidate('sendMail', {
      to: 'a@b.com', subject: 's', body: 'b',
      attachments: { file: '/path' }
    });
    assert.equal(errors.length, 1);
    assert.match(errors[0], /must be an array/);
  });

  it('rejects attachments as boolean', () => {
    const errors = mailValidate('sendMail', {
      to: 'a@b.com', subject: 's', body: 'b',
      attachments: true
    });
    assert.equal(errors.length, 1);
    assert.match(errors[0], /must be an array/);
  });

  it('accepts empty attachments array', () => {
    const errors = mailValidate('sendMail', {
      to: 'a@b.com', subject: 's', body: 'b',
      attachments: []
    });
    assert.equal(errors.length, 0);
  });

  it('accepts large number of attachments in validation', () => {
    const attachments = Array.from({ length: 100 }, (_, i) => `/path/file${i}.pdf`);
    const errors = mailValidate('sendMail', {
      to: 'a@b.com', subject: 's', body: 'b',
      attachments
    });
    assert.equal(errors.length, 0);
  });

  it('accepts mixed file paths and inline objects in array', () => {
    const errors = mailValidate('sendMail', {
      to: 'a@b.com', subject: 's', body: 'b',
      attachments: [
        '/path/to/file.pdf',
        { name: 'inline.txt', contentType: 'text/plain', base64: 'SGVsbG8=' },
        '/another/file.doc'
      ]
    });
    assert.equal(errors.length, 0);
  });
});

// ─── Contact write support stress tests ─────────────────────────

describe('Contact write: validation edge cases', () => {
  const contactTools = [
    {
      name: "createContact",
      inputSchema: {
        type: "object",
        properties: {
          email: { type: "string" },
          displayName: { type: "string" },
          firstName: { type: "string" },
          lastName: { type: "string" },
          addressBookId: { type: "string" },
        },
        required: ["email"],
      },
    },
    {
      name: "updateContact",
      inputSchema: {
        type: "object",
        properties: {
          contactId: { type: "string" },
          email: { type: "string" },
          displayName: { type: "string" },
          firstName: { type: "string" },
          lastName: { type: "string" },
        },
        required: ["contactId"],
      },
    },
    {
      name: "deleteContact",
      inputSchema: {
        type: "object",
        properties: {
          contactId: { type: "string" },
        },
        required: ["contactId"],
      },
    },
  ];
  const contactValidate = createValidator(contactTools);

  it('rejects createContact with null email', () => {
    const errors = contactValidate('createContact', { email: null });
    assert.ok(errors.some(e => e.includes('email')));
  });

  it('rejects createContact with array email', () => {
    const errors = contactValidate('createContact', { email: ['a@b.com'] });
    assert.equal(errors.length, 1);
    assert.match(errors[0], /must be string/);
  });

  it('rejects updateContact with boolean contactId', () => {
    const errors = contactValidate('updateContact', { contactId: true });
    assert.equal(errors.length, 1);
    assert.match(errors[0], /must be string/);
  });

  it('rejects deleteContact with number contactId', () => {
    const errors = contactValidate('deleteContact', { contactId: 42 });
    assert.equal(errors.length, 1);
    assert.match(errors[0], /must be string/);
  });

  it('rejects unknown params on createContact', () => {
    const errors = contactValidate('createContact', {
      email: 'a@b.com', phone: '555-1234'
    });
    assert.equal(errors.length, 1);
    assert.match(errors[0], /Unknown parameter: phone/);
  });

  it('handles prototype pollution on createContact', () => {
    const malicious = Object.create(null);
    malicious.email = 'a@b.com';
    malicious.constructor = 'attack';
    const errors = contactValidate('createContact', malicious);
    assert.ok(errors.some(e => e.includes('Unknown parameter: constructor')));
  });

  it('accepts all optional fields on createContact', () => {
    const errors = contactValidate('createContact', {
      email: 'a@b.com',
      displayName: 'Test',
      firstName: 'First',
      lastName: 'Last',
      addressBookId: 'book-1',
    });
    assert.equal(errors.length, 0);
  });

  it('accepts contactId-only for updateContact (no changes)', () => {
    const errors = contactValidate('updateContact', { contactId: 'uid-1' });
    assert.equal(errors.length, 0);
  });

  it('handles very long email on createContact', () => {
    const longEmail = 'a'.repeat(5000) + '@example.com';
    const errors = contactValidate('createContact', { email: longEmail });
    assert.equal(errors.length, 0);
  });
});

// ─── Account access control stress tests ──────────────────────────

describe('Account access: validation edge cases', () => {
  const accessTools = [
    {
      name: "getAccountAccess",
      inputSchema: { type: "object", properties: {}, required: [] },
    },
  ];
  const accessValidate = createValidator(accessTools);

  it('rejects unknown parameters on getAccountAccess', () => {
    const errors = accessValidate('getAccountAccess', {
      admin: true,
    });
    assert.equal(errors.length, 1);
    assert.match(errors[0], /Unknown parameter: admin/);
  });

  it('rejects multiple unknown parameters on getAccountAccess', () => {
    const errors = accessValidate('getAccountAccess', {
      admin: true,
      inject: 'malicious',
    });
    assert.equal(errors.length, 2);
  });

  it('handles prototype pollution on getAccountAccess', () => {
    const malicious = Object.create(null);
    malicious.constructor = 'attack';
    const errors = accessValidate('getAccountAccess', malicious);
    assert.ok(errors.some(e => e.includes('Unknown parameter: constructor')));
  });
});

// ─── Pagination stress tests ──────────────────────────────────────

/**
 * Replicate pagination logic from api.js paginate() helper.
 * Backward-compatible: returns plain array when offset is not provided,
 * structured response when offset IS provided.
 */
function paginateResults(allResults, maxResults, offset) {
  const MAX_RESULTS_CAP = 200;
  const DEFAULT_MAX = 50;

  const requestedLimit = Number(maxResults);
  const effectiveLimit = Math.min(
    Number.isFinite(requestedLimit) && requestedLimit > 0 ? Math.floor(requestedLimit) : DEFAULT_MAX,
    MAX_RESULTS_CAP
  );
  const offsetProvided = offset !== undefined && offset !== null;
  // Matches production paginate() offset logic exactly
  const effectiveOffset = (offset > 0) ? Math.floor(offset) : 0;

  const page = allResults.slice(effectiveOffset, effectiveOffset + effectiveLimit);

  if (!offsetProvided) {
    return page;
  }
  return {
    messages: page,
    totalMatches: allResults.length,
    offset: effectiveOffset,
    limit: effectiveLimit,
    hasMore: effectiveOffset + effectiveLimit < allResults.length
  };
}

describe('Pagination: boundary conditions', () => {
  const mockData = Array.from({ length: 500 }, (_, i) => ({ id: `msg-${i}`, subject: `Message ${i}` }));

  it('offset=0 returns structured response', () => {
    const result = paginateResults(mockData, 50, 0);
    assert.equal(result.messages.length, 50);
    assert.equal(result.offset, 0);
    assert.equal(result.totalMatches, 500);
    assert.equal(result.hasMore, true);
    assert.equal(result.messages[0].id, 'msg-0');
  });

  it('offset at exact page boundary works', () => {
    const result = paginateResults(mockData, 50, 50);
    assert.equal(result.messages.length, 50);
    assert.equal(result.offset, 50);
    assert.equal(result.messages[0].id, 'msg-50');
    assert.equal(result.hasMore, true);
  });

  it('offset near end returns partial page', () => {
    const result = paginateResults(mockData, 50, 480);
    assert.equal(result.messages.length, 20);
    assert.equal(result.offset, 480);
    assert.equal(result.hasMore, false);
  });

  it('offset exactly at end returns empty page', () => {
    const result = paginateResults(mockData, 50, 500);
    assert.equal(result.messages.length, 0);
    assert.equal(result.offset, 500);
    assert.equal(result.hasMore, false);
    assert.equal(result.totalMatches, 500);
  });

  it('offset beyond end returns empty page', () => {
    const result = paginateResults(mockData, 50, 9999);
    assert.equal(result.messages.length, 0);
    assert.equal(result.offset, 9999);
    assert.equal(result.hasMore, false);
  });

  it('negative offset is treated as 0', () => {
    const result = paginateResults(mockData, 50, -10);
    assert.equal(result.offset, 0);
    assert.equal(result.messages.length, 50);
    assert.equal(result.messages[0].id, 'msg-0');
  });

  it('non-numeric offset is treated as 0', () => {
    const result = paginateResults(mockData, 50, "abc");
    assert.equal(result.offset, 0);
    assert.equal(result.messages.length, 50);
  });

  it('NaN offset is treated as 0', () => {
    const result = paginateResults(mockData, 50, NaN);
    assert.equal(result.offset, 0);
  });

  it('Infinity offset returns empty page (not valid JSON, cannot arrive via MCP)', () => {
    const result = paginateResults(mockData, 50, Infinity);
    assert.equal(result.offset, Infinity);
    assert.deepStrictEqual(result.messages, []);
  });

  it('fractional offset is floored', () => {
    const result = paginateResults(mockData, 50, 10.7);
    assert.equal(result.offset, 10);
    assert.equal(result.messages[0].id, 'msg-10');
  });
});

describe('Pagination: backward compatibility', () => {
  const mockData = Array.from({ length: 100 }, (_, i) => ({ id: `msg-${i}` }));

  it('returns plain array when offset is undefined', () => {
    const result = paginateResults(mockData, 50, undefined);
    assert.ok(Array.isArray(result));
    assert.equal(result.length, 50);
  });

  it('returns plain array when offset is null', () => {
    const result = paginateResults(mockData, 50, null);
    assert.ok(Array.isArray(result));
    assert.equal(result.length, 50);
  });

  it('returns structured response when offset is 0', () => {
    const result = paginateResults(mockData, 50, 0);
    assert.ok(!Array.isArray(result));
    assert.ok(Array.isArray(result.messages));
    assert.equal(result.offset, 0);
    assert.equal(result.totalMatches, 100);
  });

  it('returns structured response when offset is positive', () => {
    const result = paginateResults(mockData, 50, 25);
    assert.ok(!Array.isArray(result));
    assert.equal(result.messages.length, 50);
    assert.equal(result.offset, 25);
  });
});

describe('Pagination: maxResults edge cases', () => {
  const mockData = Array.from({ length: 300 }, (_, i) => ({ id: `msg-${i}` }));

  it('maxResults=0 uses default (50)', () => {
    const result = paginateResults(mockData, 0, 0);
    assert.equal(result.limit, 50);
    assert.equal(result.messages.length, 50);
  });

  it('maxResults=-1 uses default (50)', () => {
    const result = paginateResults(mockData, -1, 0);
    assert.equal(result.limit, 50);
  });

  it('maxResults=1 returns single result', () => {
    const result = paginateResults(mockData, 1, 0);
    assert.equal(result.messages.length, 1);
    assert.equal(result.limit, 1);
    assert.equal(result.hasMore, true);
  });

  it('maxResults exceeding cap is clamped to 200', () => {
    const result = paginateResults(mockData, 999, 0);
    assert.equal(result.limit, 200);
    assert.equal(result.messages.length, 200);
  });

  it('maxResults=200 (at cap) works', () => {
    const result = paginateResults(mockData, 200, 0);
    assert.equal(result.limit, 200);
    assert.equal(result.messages.length, 200);
    assert.equal(result.hasMore, true);
  });

  it('maxResults as string is treated as default', () => {
    const result = paginateResults(mockData, "fifty", 0);
    assert.equal(result.limit, 50);
  });

  it('maxResults=NaN uses default', () => {
    const result = paginateResults(mockData, NaN, 0);
    assert.equal(result.limit, 50);
  });

  it('empty results with pagination returns correct metadata', () => {
    const result = paginateResults([], 50, 0);
    assert.equal(result.messages.length, 0);
    assert.equal(result.totalMatches, 0);
    assert.equal(result.hasMore, false);
    assert.equal(result.offset, 0);
  });

  it('single result with offset=0 has hasMore=false', () => {
    const result = paginateResults([{ id: 'only' }], 50, 0);
    assert.equal(result.messages.length, 1);
    assert.equal(result.hasMore, false);
  });
});

describe('Pagination: sequential page traversal', () => {
  const mockData = Array.from({ length: 137 }, (_, i) => ({ id: `msg-${i}` }));

  it('can traverse all pages without gaps or duplicates', () => {
    const pageSize = 25;
    const allCollected = [];
    let offset = 0;
    let pages = 0;

    while (true) {
      const result = paginateResults(mockData, pageSize, offset);
      allCollected.push(...result.messages);
      pages++;

      if (!result.hasMore) break;
      offset += pageSize;

      if (pages > 100) throw new Error('Pagination infinite loop detected');
    }

    assert.equal(allCollected.length, 137);
    const ids = new Set(allCollected.map(m => m.id));
    assert.equal(ids.size, 137);
    assert.equal(pages, 6);
  });

  it('page traversal with maxResults=1 (worst case) collects all items', () => {
    const smallData = Array.from({ length: 10 }, (_, i) => ({ id: `msg-${i}` }));
    const allCollected = [];
    let offset = 0;
    let pages = 0;

    while (true) {
      const result = paginateResults(smallData, 1, offset);
      allCollected.push(...result.messages);
      pages++;
      if (!result.hasMore) break;
      offset += 1;
      if (pages > 100) throw new Error('Infinite loop');
    }

    assert.equal(allCollected.length, 10);
    assert.equal(pages, 10);
  });
});

// ─── Validation adversarial tests ─────────────────────────────────

describe('Validation: adversarial and edge-case inputs', () => {
  const advTools = [
    { name: "searchMessages", inputSchema: {
      type: "object",
      properties: {
        query: { type: "string" }, folderPath: { type: "string" },
        maxResults: { type: "number" }, offset: { type: "number" },
        unreadOnly: { type: "boolean" },
      },
      required: ["query"],
    }},
    { name: "getMessage", inputSchema: {
      type: "object",
      properties: {
        messageId: { type: "string" }, folderPath: { type: "string" },
        saveAttachments: { type: "boolean" },
      },
      required: ["messageId", "folderPath"],
    }},
    { name: "sendMail", inputSchema: {
      type: "object",
      properties: {
        to: { type: "string" }, subject: { type: "string" }, body: { type: "string" },
        attachments: { type: "array", items: { type: "string" } },
      },
      required: ["to", "subject", "body"],
    }},
  ];
  const advValidate = createValidator(advTools);

  it('handles empty string for required string param', () => {
    const errors = advValidate('getMessage', {
      messageId: '',
      folderPath: '',
    });
    assert.equal(errors.length, 0);
  });

  it('rejects constructor pollution attempt', () => {
    const errors = advValidate('searchMessages', {
      query: 'test',
      constructor: 'evil',
    });
    assert.ok(errors.some(e => e.includes('Unknown parameter: constructor')));
  });

  it('handles very long parameter names gracefully', () => {
    const longKey = 'x'.repeat(10000);
    const args = { query: 'test', [longKey]: 'value' };
    const errors = advValidate('searchMessages', args);
    assert.ok(errors.some(e => e.includes('Unknown parameter')));
  });

  it('handles very long parameter values', () => {
    const longValue = 'x'.repeat(100000);
    const errors = advValidate('searchMessages', { query: longValue });
    assert.equal(errors.length, 0);
  });

  it('handles deeply nested object where string expected', () => {
    const errors = advValidate('searchMessages', {
      query: { nested: { deep: { very: 'deep' } } },
    });
    assert.ok(errors.some(e => e.includes('must be string')));
  });

  it('handles array where string expected', () => {
    const errors = advValidate('getMessage', {
      messageId: ['id1', 'id2'],
      folderPath: '/INBOX',
    });
    assert.ok(errors.some(e => e.includes('must be string')));
  });

  it('handles 0 as valid number', () => {
    const errors = advValidate('searchMessages', { query: 'test', maxResults: 0 });
    assert.equal(errors.length, 0);
  });

  it('handles boolean false as valid required field type check', () => {
    const errors = advValidate('getMessage', {
      messageId: false,
      folderPath: '/INBOX',
    });
    assert.ok(errors.some(e => e.includes('must be string')));
  });

  it('handles massive number of unknown parameters', () => {
    const args = { query: 'test' };
    for (let i = 0; i < 100; i++) {
      args[`unknown_param_${i}`] = `value_${i}`;
    }
    const errors = advValidate('searchMessages', args);
    assert.equal(errors.length, 100);
  });

  it('handles all required params wrong type simultaneously', () => {
    const errors = advValidate('sendMail', {
      to: 123,
      subject: true,
      body: ['not', 'a', 'string'],
    });
    assert.equal(errors.length, 3);
  });
});

// ─── Coercion tests ──────────────────────────────────────────────

describe('Coercion: string-to-type conversion', () => {
  /**
   * Replicate coerceToolArgs from api.js.
   */
  function coerceToolArgs(name, args, toolSchemas) {
    const schema = toolSchemas[name];
    if (!schema) return args;
    const props = schema.properties || {};
    for (const [key, value] of Object.entries(args)) {
      if (value === undefined || value === null) continue;
      const propSchema = Object.prototype.hasOwnProperty.call(props, key) ? props[key] : undefined;
      if (!propSchema) continue;
      const expected = propSchema.type;
      if (expected === "boolean" && typeof value === "string") {
        if (value === "true") args[key] = true;
        else if (value === "false") args[key] = false;
      } else if (expected === "number" && typeof value === "string") {
        const n = Number(value);
        if (Number.isFinite(n)) args[key] = n;
      } else if (expected === "array" && typeof value === "string") {
        try {
          const parsed = JSON.parse(value);
          if (Array.isArray(parsed)) args[key] = parsed;
        } catch { /* leave as-is */ }
      }
    }
    return args;
  }

  const schemas = {
    searchMessages: {
      properties: {
        query: { type: "string" },
        maxResults: { type: "number" },
        offset: { type: "number" },
        unreadOnly: { type: "boolean" },
      },
      required: ["query"],
    },
    sendMail: {
      properties: {
        to: { type: "string" },
        subject: { type: "string" },
        body: { type: "string" },
        attachments: { type: "array" },
      },
      required: ["to", "subject", "body"],
    },
  };

  it('coerces "true" string to boolean true', () => {
    const args = { query: 'test', unreadOnly: 'true' };
    coerceToolArgs('searchMessages', args, schemas);
    assert.strictEqual(args.unreadOnly, true);
  });

  it('coerces "false" string to boolean false', () => {
    const args = { query: 'test', unreadOnly: 'false' };
    coerceToolArgs('searchMessages', args, schemas);
    assert.strictEqual(args.unreadOnly, false);
  });

  it('leaves non-boolean strings unchanged for boolean fields', () => {
    const args = { query: 'test', unreadOnly: 'yes' };
    coerceToolArgs('searchMessages', args, schemas);
    assert.strictEqual(args.unreadOnly, 'yes');
  });

  it('coerces numeric string to number', () => {
    const args = { query: 'test', maxResults: '50' };
    coerceToolArgs('searchMessages', args, schemas);
    assert.strictEqual(args.maxResults, 50);
  });

  it('coerces "0" string to number 0', () => {
    const args = { query: 'test', offset: '0' };
    coerceToolArgs('searchMessages', args, schemas);
    assert.strictEqual(args.offset, 0);
  });

  it('leaves non-numeric strings unchanged for number fields', () => {
    const args = { query: 'test', maxResults: 'fifty' };
    coerceToolArgs('searchMessages', args, schemas);
    assert.strictEqual(args.maxResults, 'fifty');
  });

  it('coerces JSON array string to array', () => {
    const args = { to: 'a@b.com', subject: 's', body: 'b', attachments: '["file1.pdf","file2.pdf"]' };
    coerceToolArgs('sendMail', args, schemas);
    assert.deepStrictEqual(args.attachments, ['file1.pdf', 'file2.pdf']);
  });

  it('leaves invalid JSON string unchanged for array fields', () => {
    const args = { to: 'a@b.com', subject: 's', body: 'b', attachments: 'not-json' };
    coerceToolArgs('sendMail', args, schemas);
    assert.strictEqual(args.attachments, 'not-json');
  });

  it('does not coerce JSON object string to array', () => {
    const args = { to: 'a@b.com', subject: 's', body: 'b', attachments: '{"key":"val"}' };
    coerceToolArgs('sendMail', args, schemas);
    assert.strictEqual(args.attachments, '{"key":"val"}');
  });

  it('does not touch unknown parameters', () => {
    const args = { query: 'test', unknown: 'true' };
    coerceToolArgs('searchMessages', args, schemas);
    assert.strictEqual(args.unknown, 'true');
  });
});
