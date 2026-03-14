/**
 * Tests for the schema validation logic used in api.js.
 *
 * The validator is shared via test/helpers.cjs which replicates the exact
 * logic from api.js. This ensures validation behavior is correct before
 * deploying to the extension.
 *
 * Covers:
 * - Required parameter enforcement
 * - Type checking (string, number, boolean, array, object)
 * - Unknown parameter rejection
 * - Edge cases (null values, empty args, nested schemas)
 */

const { describe, it } = require('node:test');
const assert = require('node:assert/strict');

const { createValidator } = require('./helpers.cjs');

// Sample tool definitions matching the shapes in api.js
const sampleTools = [
  {
    name: "listAccounts",
    inputSchema: { type: "object", properties: {}, required: [] },
  },
  {
    name: "getMessage",
    inputSchema: {
      type: "object",
      properties: {
        messageId: { type: "string" },
        folderPath: { type: "string" },
        saveAttachments: { type: "boolean" },
      },
      required: ["messageId", "folderPath"],
    },
  },
  {
    name: "searchMessages",
    inputSchema: {
      type: "object",
      properties: {
        query: { type: "string" },
        folderPath: { type: "string" },
        maxResults: { type: "number" },
        offset: { type: "number" },
        unreadOnly: { type: "boolean" },
        tag: { type: "string" },
      },
      required: ["query"],
    },
  },
  {
    name: "deleteMessages",
    inputSchema: {
      type: "object",
      properties: {
        messageIds: { type: "array", items: { type: "string" } },
        folderPath: { type: "string" },
      },
      required: ["messageIds", "folderPath"],
    },
  },
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
  {
    name: "renameFolder",
    inputSchema: {
      type: "object",
      properties: {
        folderPath: { type: "string" },
        newName: { type: "string" },
      },
      required: ["folderPath", "newName"],
    },
  },
  {
    name: "deleteFolder",
    inputSchema: {
      type: "object",
      properties: {
        folderPath: { type: "string" },
      },
      required: ["folderPath"],
    },
  },
  {
    name: "moveFolder",
    inputSchema: {
      type: "object",
      properties: {
        folderPath: { type: "string" },
        newParentPath: { type: "string" },
      },
      required: ["folderPath", "newParentPath"],
    },
  },
  {
    name: "getAccountAccess",
    inputSchema: { type: "object", properties: {}, required: [] },
  },
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
];

const validate = createValidator(sampleTools);

describe('Validation: unknown tools', () => {
  it('rejects unknown tool names', () => {
    const errors = validate('nonExistentTool', {});
    assert.equal(errors.length, 1);
    assert.match(errors[0], /Unknown tool/);
  });
});

describe('Validation: required parameters', () => {
  it('passes when all required params are present', () => {
    const errors = validate('getMessage', {
      messageId: '123',
      folderPath: 'imap://user@server/INBOX',
    });
    assert.equal(errors.length, 0);
  });

  it('fails when required params are missing', () => {
    const errors = validate('getMessage', {});
    assert.equal(errors.length, 2);
    assert.ok(errors.some(e => e.includes('messageId')));
    assert.ok(errors.some(e => e.includes('folderPath')));
  });

  it('fails when required param is null', () => {
    const errors = validate('getMessage', {
      messageId: null,
      folderPath: 'imap://user@server/INBOX',
    });
    assert.equal(errors.length, 1);
    assert.match(errors[0], /messageId/);
  });

  it('passes with no required params and empty args', () => {
    const errors = validate('listAccounts', {});
    assert.equal(errors.length, 0);
  });
});

describe('Validation: type checking', () => {
  it('accepts correct string type', () => {
    const errors = validate('searchMessages', { query: 'test' });
    assert.equal(errors.length, 0);
  });

  it('rejects number where string expected', () => {
    const errors = validate('searchMessages', { query: 123 });
    assert.equal(errors.length, 1);
    assert.match(errors[0], /must be string/);
  });

  it('accepts correct number type', () => {
    const errors = validate('searchMessages', { query: 'test', maxResults: 50 });
    assert.equal(errors.length, 0);
  });

  it('rejects string where number expected', () => {
    const errors = validate('searchMessages', { query: 'test', maxResults: '50' });
    assert.equal(errors.length, 1);
    assert.match(errors[0], /must be number/);
  });

  it('accepts correct boolean type', () => {
    const errors = validate('searchMessages', { query: 'test', unreadOnly: true });
    assert.equal(errors.length, 0);
  });

  it('rejects string where boolean expected', () => {
    const errors = validate('getMessage', {
      messageId: '123',
      folderPath: '/INBOX',
      saveAttachments: 'yes',
    });
    assert.equal(errors.length, 1);
    assert.match(errors[0], /must be boolean/);
  });

  it('accepts correct array type', () => {
    const errors = validate('deleteMessages', {
      messageIds: ['1', '2'],
      folderPath: '/INBOX',
    });
    assert.equal(errors.length, 0);
  });

  it('rejects string where array expected', () => {
    const errors = validate('deleteMessages', {
      messageIds: '1,2',
      folderPath: '/INBOX',
    });
    assert.equal(errors.length, 1);
    assert.match(errors[0], /must be an array/);
  });

  it('rejects object where array expected', () => {
    const errors = validate('deleteMessages', {
      messageIds: { id: '1' },
      folderPath: '/INBOX',
    });
    assert.equal(errors.length, 1);
    assert.match(errors[0], /must be an array/);
  });
});

describe('Validation: unknown parameters', () => {
  it('rejects unknown parameter names', () => {
    const errors = validate('listAccounts', { bogus: 'value' });
    assert.equal(errors.length, 1);
    assert.match(errors[0], /Unknown parameter: bogus/);
  });

  it('rejects multiple unknown parameters', () => {
    const errors = validate('searchMessages', {
      query: 'test',
      foo: 1,
      bar: 2,
    });
    assert.equal(errors.length, 2);
  });
});

describe('Validation: multiple errors', () => {
  it('reports both missing required and wrong type in one pass', () => {
    const errors = validate('sendMail', {
      to: 'user@example.com',
      // subject missing (required)
      body: 123, // wrong type
    });
    assert.ok(errors.length >= 2);
    assert.ok(errors.some(e => e.includes('subject')));
    assert.ok(errors.some(e => e.includes('must be string')));
  });
});

describe('Validation: optional parameters', () => {
  it('allows omitting optional parameters', () => {
    const errors = validate('searchMessages', { query: 'test' });
    assert.equal(errors.length, 0);
  });

  it('allows providing optional parameters with correct types', () => {
    const errors = validate('searchMessages', {
      query: 'test',
      folderPath: '/INBOX',
      maxResults: 10,
      unreadOnly: false,
    });
    assert.equal(errors.length, 0);
  });
});

describe('Validation: updateMessage with tags', () => {
  it('accepts addTags as array', () => {
    const errors = validate('updateMessage', {
      folderPath: '/INBOX',
      messageId: 'msg1',
      addTags: ['$label1', 'project-x'],
    });
    assert.equal(errors.length, 0);
  });

  it('accepts removeTags as array', () => {
    const errors = validate('updateMessage', {
      folderPath: '/INBOX',
      messageId: 'msg1',
      removeTags: ['$label1'],
    });
    assert.equal(errors.length, 0);
  });

  it('rejects addTags as string', () => {
    const errors = validate('updateMessage', {
      folderPath: '/INBOX',
      messageId: 'msg1',
      addTags: '$label1',
    });
    assert.equal(errors.length, 1);
    assert.match(errors[0], /must be an array/);
  });

  it('rejects removeTags as number', () => {
    const errors = validate('updateMessage', {
      folderPath: '/INBOX',
      messageId: 'msg1',
      removeTags: 123,
    });
    assert.equal(errors.length, 1);
    assert.match(errors[0], /must be an array/);
  });

  it('accepts tag filter as string in searchMessages', () => {
    const errors = validate('searchMessages', {
      query: 'test',
      tag: '$label1',
    });
    assert.equal(errors.length, 0);
  });

  it('rejects tag filter as number in searchMessages', () => {
    const errors = validate('searchMessages', {
      query: 'test',
      tag: 123,
    });
    assert.equal(errors.length, 1);
    assert.match(errors[0], /must be string/);
  });
});

describe('Validation: account access control', () => {
  it('getAccountAccess accepts no params', () => {
    const errors = validate('getAccountAccess', {});
    assert.equal(errors.length, 0);
  });

  it('getAccountAccess rejects unknown params', () => {
    const errors = validate('getAccountAccess', { bogus: 'value' });
    assert.equal(errors.length, 1);
    assert.match(errors[0], /Unknown parameter/);
  });
});

describe('Validation: folder management', () => {
  it('renameFolder requires both params', () => {
    const errors = validate('renameFolder', {});
    assert.equal(errors.length, 2);
    assert.ok(errors.some(e => e.includes('folderPath')));
    assert.ok(errors.some(e => e.includes('newName')));
  });

  it('renameFolder accepts valid params', () => {
    const errors = validate('renameFolder', {
      folderPath: 'imap://user@server/INBOX/Old',
      newName: 'New',
    });
    assert.equal(errors.length, 0);
  });

  it('renameFolder rejects number for newName', () => {
    const errors = validate('renameFolder', {
      folderPath: 'imap://user@server/INBOX/Old',
      newName: 123,
    });
    assert.equal(errors.length, 1);
    assert.match(errors[0], /must be string/);
  });

  it('deleteFolder requires folderPath', () => {
    const errors = validate('deleteFolder', {});
    assert.equal(errors.length, 1);
    assert.match(errors[0], /folderPath/);
  });

  it('deleteFolder accepts valid path', () => {
    const errors = validate('deleteFolder', {
      folderPath: 'imap://user@server/INBOX/ToDelete',
    });
    assert.equal(errors.length, 0);
  });

  it('moveFolder requires both params', () => {
    const errors = validate('moveFolder', {});
    assert.equal(errors.length, 2);
  });

  it('moveFolder accepts valid params', () => {
    const errors = validate('moveFolder', {
      folderPath: 'imap://user@server/INBOX/Source',
      newParentPath: 'imap://user@server/Archive',
    });
    assert.equal(errors.length, 0);
  });

  it('moveFolder rejects unknown params', () => {
    const errors = validate('moveFolder', {
      folderPath: '/Source',
      newParentPath: '/Dest',
      recursive: true,
    });
    assert.equal(errors.length, 1);
    assert.match(errors[0], /Unknown parameter/);
  });
});

describe('Validation: attachment sending', () => {
  it('accepts attachments as array of strings (file paths)', () => {
    const errors = validate('sendMail', {
      to: 'user@example.com',
      subject: 'test',
      body: 'hello',
      attachments: ['/path/to/file.pdf'],
    });
    assert.equal(errors.length, 0);
  });

  it('accepts attachments as array of objects (inline base64)', () => {
    const errors = validate('sendMail', {
      to: 'user@example.com',
      subject: 'test',
      body: 'hello',
      attachments: [{ name: 'file.pdf', contentType: 'application/pdf', base64: 'AAAA' }],
    });
    assert.equal(errors.length, 0);
  });

  it('accepts mixed attachment types', () => {
    const errors = validate('sendMail', {
      to: 'user@example.com',
      subject: 'test',
      body: 'hello',
      attachments: ['/path/to/file.pdf', { name: 'img.png', base64: 'AAAA' }],
    });
    assert.equal(errors.length, 0);
  });

  it('rejects attachments as string', () => {
    const errors = validate('sendMail', {
      to: 'user@example.com',
      subject: 'test',
      body: 'hello',
      attachments: '/path/to/file.pdf',
    });
    assert.equal(errors.length, 1);
    assert.match(errors[0], /must be an array/);
  });
});

describe('Validation: contact write operations', () => {
  it('createContact requires email', () => {
    const errors = validate('createContact', {});
    assert.equal(errors.length, 1);
    assert.match(errors[0], /email/);
  });

  it('createContact accepts valid params', () => {
    const errors = validate('createContact', {
      email: 'user@example.com',
      displayName: 'Test User',
      firstName: 'Test',
      lastName: 'User',
    });
    assert.equal(errors.length, 0);
  });

  it('createContact rejects number for email', () => {
    const errors = validate('createContact', { email: 123 });
    assert.equal(errors.length, 1);
    assert.match(errors[0], /must be string/);
  });

  it('updateContact requires contactId', () => {
    const errors = validate('updateContact', {});
    assert.equal(errors.length, 1);
    assert.match(errors[0], /contactId/);
  });

  it('updateContact accepts contactId with optional fields', () => {
    const errors = validate('updateContact', {
      contactId: 'uid-123',
      email: 'new@example.com',
    });
    assert.equal(errors.length, 0);
  });

  it('deleteContact requires contactId', () => {
    const errors = validate('deleteContact', {});
    assert.equal(errors.length, 1);
    assert.match(errors[0], /contactId/);
  });

  it('deleteContact accepts valid contactId', () => {
    const errors = validate('deleteContact', { contactId: 'uid-456' });
    assert.equal(errors.length, 0);
  });

  it('deleteContact rejects unknown params', () => {
    const errors = validate('deleteContact', {
      contactId: 'uid-456',
      force: true,
    });
    assert.equal(errors.length, 1);
    assert.match(errors[0], /Unknown parameter/);
  });
});
