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
      name: "setAccountAccess",
      inputSchema: {
        type: "object",
        properties: {
          allowedAccountIds: { type: "array", items: { type: "string" } },
        },
        required: ["allowedAccountIds"],
      },
    },
    {
      name: "getAccountAccess",
      inputSchema: { type: "object", properties: {}, required: [] },
    },
  ];
  const accessValidate = createValidator(accessTools);

  it('rejects allowedAccountIds as object', () => {
    const errors = accessValidate('setAccountAccess', {
      allowedAccountIds: { id: 'account1' }
    });
    assert.equal(errors.length, 1);
    assert.match(errors[0], /must be an array/);
  });

  it('rejects allowedAccountIds as boolean', () => {
    const errors = accessValidate('setAccountAccess', {
      allowedAccountIds: true
    });
    assert.equal(errors.length, 1);
    assert.match(errors[0], /must be an array/);
  });

  it('rejects allowedAccountIds as number', () => {
    const errors = accessValidate('setAccountAccess', {
      allowedAccountIds: 42
    });
    assert.equal(errors.length, 1);
    assert.match(errors[0], /must be an array/);
  });

  it('accepts large number of account IDs', () => {
    const ids = Array.from({ length: 50 }, (_, i) => `account${i}`);
    const errors = accessValidate('setAccountAccess', {
      allowedAccountIds: ids
    });
    assert.equal(errors.length, 0);
  });

  it('rejects unknown parameters on setAccountAccess', () => {
    const errors = accessValidate('setAccountAccess', {
      allowedAccountIds: ['acct1'],
      inject: 'malicious',
    });
    assert.equal(errors.length, 1);
    assert.match(errors[0], /Unknown parameter: inject/);
  });

  it('rejects unknown parameters on getAccountAccess', () => {
    const errors = accessValidate('getAccountAccess', {
      admin: true,
    });
    assert.equal(errors.length, 1);
    assert.match(errors[0], /Unknown parameter: admin/);
  });

  it('handles prototype pollution on setAccountAccess', () => {
    const malicious = Object.create(null);
    malicious.allowedAccountIds = ['acct1'];
    malicious.constructor = 'attack';
    const errors = accessValidate('setAccountAccess', malicious);
    assert.ok(errors.some(e => e.includes('Unknown parameter: constructor')));
  });
});
