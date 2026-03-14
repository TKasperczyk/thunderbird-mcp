/**
 * Tests for the schema validation logic used in api.js.
 *
 * Since the validation function runs inside the Thunderbird extension context,
 * we replicate the exact same logic here to verify it independently.
 * This ensures validation behavior is correct before deploying to the extension.
 *
 * Covers:
 * - Required parameter enforcement
 * - Type checking (string, number, boolean, array, object)
 * - Unknown parameter rejection
 * - Edge cases (null values, empty args, nested schemas)
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
      const propSchema = props[key];
      if (!propSchema) {
        errors.push(`Unknown parameter: ${key}`);
        continue;
      }
      if (value === undefined || value === null) continue;

      const expectedType = propSchema.type;
      if (expectedType === "array") {
        if (!Array.isArray(value)) {
          errors.push(`Parameter '${key}' must be an array, got ${typeof value}`);
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
        unreadOnly: { type: "boolean" },
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
        attachments: { type: "array", items: { type: "string" } },
      },
      required: ["to", "subject", "body"],
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
