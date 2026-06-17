/**
 * Validation tests for new tool schemas added in this PR.
 *
 * Since the validation function runs inside the Thunderbird extension context,
 * we replicate the exact same logic here to verify it independently.
 * This ensures validation behavior is correct before deploying to the extension.
 *
 * Covers:
 * - Tag operations (addTags/removeTags on updateMessage, tag filter on searchMessages)
 * - Folder management (renameFolder, deleteFolder, moveFolder)
 * - Attachment sending (file paths and inline base64 objects)
 * - Contact write operations (createContact, updateContact, deleteContact)
 */

const { describe, it } = require('node:test');
const assert = require('node:assert/strict');

/**
 * Exact copy of validateToolArgs from api.js.
 * Kept in sync manually — if the logic in api.js changes, update here too.
 */
function createValidator(tools) {
  const toolSchemas = Object.create(null);
  for (const t of tools) {
    toolSchemas[t.name] = t.inputSchema;
  }

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
      if (!propSchema) {
        errors.push(`Unknown parameter: ${key}`);
        continue;
      }
      if (value === undefined || value === null) continue;

      const expectedType = propSchema.type;
      if (expectedType === "array") {
        if (!Array.isArray(value)) {
          errors.push(`Parameter '${key}' must be an array, got ${typeof value}`);
        } else {
          if (propSchema.minItems !== undefined && value.length < propSchema.minItems) {
            errors.push(`Parameter '${key}' must contain at least ${propSchema.minItems} item(s)`);
          }
          if (propSchema.maxItems !== undefined && value.length > propSchema.maxItems) {
            errors.push(`Parameter '${key}' must contain at most ${propSchema.maxItems} item(s)`);
          }
        }
      } else if (expectedType === "object") {
        if (typeof value !== "object" || Array.isArray(value)) {
          errors.push(`Parameter '${key}' must be an object, got ${Array.isArray(value) ? "array" : typeof value}`);
        }
      } else if (expectedType && typeof value !== expectedType) {
        errors.push(`Parameter '${key}' must be ${expectedType}, got ${typeof value}`);
      }
    }

    return errors;
  };
}

// Tool definitions matching the schemas in api.js for new tools
const sampleTools = [
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
        dedupByMessageId: { type: "boolean" },
      },
      required: ["query"],
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
        attachments: { type: "array", items: { oneOf: [{ type: "string" }, { type: "object" }] } },
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
    name: "getMessages",
    inputSchema: {
      type: "object",
      properties: {
        messages: {
          type: "array",
          minItems: 1,
          maxItems: 10,
          items: {
            type: "object",
            properties: {
              messageId: { type: "string" },
              folderPath: { type: "string" },
            },
            required: ["messageId", "folderPath"],
          },
        },
        saveAttachments: { type: "boolean" },
        bodyFormat: { type: "string" },
        rawSource: { type: "boolean" },
      },
      required: ["messages"],
    },
  },
  {
    name: "getMessageHeaders",
    inputSchema: {
      type: "object",
      properties: {
        messageId: { type: "string" },
        folderPath: { type: "string" },
      },
      required: ["messageId", "folderPath"],
    },
  },
  {
    name: "batchGetMessageHeaders",
    inputSchema: {
      type: "object",
      properties: {
        messageIds: {
          type: "array",
          items: { type: "string" },
        },
        folderPath: { type: "string" },
      },
      required: ["messageIds", "folderPath"],
    },
  },
  {
    name: "getAccountAccess",
    inputSchema: { type: "object", properties: {}, required: [] },
  },
];

const validate = createValidator(sampleTools);

describe('Validation: updateMessage with tags', () => {
  it('accepts addTags as array', () => {
    const errors = validate('updateMessage', {
      folderPath: 'imap://user@server/INBOX',
      addTags: ['$label1', '$label2'],
    });
    assert.equal(errors.length, 0);
  });

  it('accepts removeTags as array', () => {
    const errors = validate('updateMessage', {
      folderPath: 'imap://user@server/INBOX',
      removeTags: ['$label1'],
    });
    assert.equal(errors.length, 0);
  });

  it('rejects addTags as string', () => {
    const errors = validate('updateMessage', {
      folderPath: 'imap://user@server/INBOX',
      addTags: '$label1',
    });
    assert.equal(errors.length, 1);
    assert.match(errors[0], /must be an array/);
  });

  it('rejects removeTags as number', () => {
    const errors = validate('updateMessage', {
      folderPath: 'imap://user@server/INBOX',
      removeTags: 42,
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
      tag: 42,
    });
    assert.equal(errors.length, 1);
    assert.match(errors[0], /must be string/);
  });

  it('accepts dedupByMessageId as boolean in searchMessages', () => {
    const errors = validate('searchMessages', {
      query: 'test',
      dedupByMessageId: false,
    });
    assert.equal(errors.length, 0);
  });
});

describe('Validation: pagination parameters', () => {
  it('accepts offset as number', () => {
    const errors = validate('searchMessages', {
      query: 'test',
      offset: 50,
    });
    assert.equal(errors.length, 0);
  });

  it('rejects offset as string', () => {
    const errors = validate('searchMessages', {
      query: 'test',
      offset: 'fifty',
    });
    assert.equal(errors.length, 1);
    assert.match(errors[0], /must be number/);
  });

  it('accepts offset=0', () => {
    const errors = validate('searchMessages', {
      query: 'test',
      offset: 0,
    });
    assert.equal(errors.length, 0);
  });

  it('accepts offset with maxResults', () => {
    const errors = validate('searchMessages', {
      query: 'test',
      offset: 100,
      maxResults: 50,
    });
    assert.equal(errors.length, 0);
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
    });
    assert.equal(errors.length, 0);
  });

  it('createContact rejects number for email', () => {
    const errors = validate('createContact', {
      email: 12345,
    });
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
      firstName: 'New',
    });
    assert.equal(errors.length, 0);
  });

  it('deleteContact requires contactId', () => {
    const errors = validate('deleteContact', {});
    assert.equal(errors.length, 1);
    assert.match(errors[0], /contactId/);
  });

  it('deleteContact accepts valid contactId', () => {
    const errors = validate('deleteContact', {
      contactId: 'uid-456',
    });
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

describe('Validation: getMessages batch size', () => {
  const messageRef = { messageId: "message-1", folderPath: "imap://user@server/INBOX" };

  it('accepts messages at the configured cap', () => {
    const errors = validate('getMessages', {
      messages: Array.from({ length: 10 }, (_, index) => ({
        ...messageRef,
        messageId: `message-${index}`,
      })),
    });
    assert.deepStrictEqual(errors, []);
  });

  it('rejects empty messages arrays', () => {
    const errors = validate('getMessages', { messages: [] });
    assert.equal(errors.length, 1);
    assert.match(errors[0], /at least 1/);
  });

  it('rejects messages arrays over the configured cap', () => {
    const errors = validate('getMessages', {
      messages: Array.from({ length: 11 }, (_, index) => ({
        ...messageRef,
        messageId: `message-${index}`,
      })),
    });
    assert.equal(errors.length, 1);
    assert.match(errors[0], /at most 10/);
  });
});

describe('Validation: getMessageHeaders', () => {
  it('accepts messageId + folderPath', () => {
    const errors = validate('getMessageHeaders', {
      messageId: '<abc@example.com>',
      folderPath: 'imap://user@server/INBOX',
    });
    assert.deepStrictEqual(errors, []);
  });

  it('rejects missing messageId', () => {
    const errors = validate('getMessageHeaders', {
      folderPath: 'imap://user@server/INBOX',
    });
    assert.equal(errors.length, 1);
    assert.match(errors[0], /Missing required parameter: messageId/);
  });

  it('rejects missing folderPath', () => {
    const errors = validate('getMessageHeaders', {
      messageId: '<abc@example.com>',
    });
    assert.equal(errors.length, 1);
    assert.match(errors[0], /Missing required parameter: folderPath/);
  });

  it('rejects messageId of the wrong type', () => {
    const errors = validate('getMessageHeaders', {
      messageId: 42,
      folderPath: 'imap://user@server/INBOX',
    });
    assert.equal(errors.length, 1);
    assert.match(errors[0], /must be string/);
  });

  it('rejects unknown parameters', () => {
    const errors = validate('getMessageHeaders', {
      messageId: '<abc@example.com>',
      folderPath: 'imap://user@server/INBOX',
      headers: ['Subject'],
    });
    assert.equal(errors.length, 1);
    assert.match(errors[0], /Unknown parameter: headers/);
  });
});

describe('Validation: batchGetMessageHeaders', () => {
  it('accepts a valid batch', () => {
    const errors = validate('batchGetMessageHeaders', {
      messageIds: Array.from({ length: 10 }, (_, index) => `<message-${index}@example.com>`),
      folderPath: 'imap://user@server/INBOX',
    });
    assert.deepStrictEqual(errors, []);
  });

  // No schema minItems: an empty array is valid at the dispatch layer; the
  // handler short-circuits it to { headers: {}, total: 0, failed: 0 }.
  it('accepts an empty messageIds array (handler returns the empty fast path)', () => {
    const errors = validate('batchGetMessageHeaders', {
      messageIds: [],
      folderPath: 'imap://user@server/INBOX',
    });
    assert.deepStrictEqual(errors, []);
  });

  // No schema maxItems: the 200-id hard cap lives in the handler (returning a
  // clear { error } value), not in validation, so large arrays pass validation.
  it('does not bound messageIds at the schema level', () => {
    const errors = validate('batchGetMessageHeaders', {
      messageIds: Array.from({ length: 250 }, (_, index) => `<message-${index}@example.com>`),
      folderPath: 'imap://user@server/INBOX',
    });
    assert.deepStrictEqual(errors, []);
  });

  it('rejects messageIds passed as a non-array', () => {
    const errors = validate('batchGetMessageHeaders', {
      messageIds: '<abc@example.com>',
      folderPath: 'imap://user@server/INBOX',
    });
    assert.equal(errors.length, 1);
    assert.match(errors[0], /must be an array/);
  });

  it('rejects missing folderPath', () => {
    const errors = validate('batchGetMessageHeaders', {
      messageIds: ['<abc@example.com>'],
    });
    assert.equal(errors.length, 1);
    assert.match(errors[0], /Missing required parameter: folderPath/);
  });

  // Validation is SHALLOW: it enforces the messageIds array type but does NOT
  // recurse into items, so a null element passes top-level validation (per-id
  // handling happens inside the handler, which records a per-item "not found in
  // this folder" error rather than aborting the batch).
  it('does not reject a null array item at the top level (shallow validation)', () => {
    const errors = validate('batchGetMessageHeaders', {
      messageIds: ['<abc@example.com>', null],
      folderPath: 'imap://user@server/INBOX',
    });
    assert.deepStrictEqual(errors, []);
  });
});
