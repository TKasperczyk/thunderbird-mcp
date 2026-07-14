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
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');

const apiSource = fs.readFileSync(
  path.resolve(__dirname, '../extension/mcp_server/api.js'),
  'utf8'
);

function getMarkedApiSnippet(startMarker, endMarker) {
  const start = apiSource.indexOf(startMarker);
  const end = apiSource.indexOf(endMarker, start + startMarker.length);
  assert.ok(start >= 0, `api.js marker missing: ${startMarker}`);
  assert.ok(end > start, `api.js marker missing: ${endMarker}`);
  return apiSource.slice(start, end);
}

function loadProductionAttachmentValidation() {
  const sandbox = {
    getConfiguredGetMessagesLimit: () => 20,
    MAX_BASE64_SIZE: 25 * 1024 * 1024,
  };
  vm.createContext(sandbox);
  vm.runInContext([
    getMarkedApiSnippet('// BEGIN INLINE ATTACHMENT BASE64 HELPERS', '// END INLINE ATTACHMENT BASE64 HELPERS'),
    getMarkedApiSnippet('// BEGIN TOOL SCHEMA BUILDER', '// END TOOL SCHEMA BUILDER'),
    getMarkedApiSnippet('// BEGIN TOOL SCHEMA VALIDATOR', '// END TOOL SCHEMA VALIDATOR'),
    getMarkedApiSnippet('// BEGIN OUTBOUND ATTACHMENT CONVERSION', '// END OUTBOUND ATTACHMENT CONVERSION'),
    'this.isValidBase64 = isValidBase64;',
    'this.buildTools = buildTools;',
    'this.validateAgainstSchema = validateAgainstSchema;',
    'this.filePathsToAttachDescs = filePathsToAttachDescs;',
  ].join('\n'), sandbox);
  return sandbox;
}

const productionAttachmentValidation = loadProductionAttachmentValidation();

function validateProductionToolArgs(name, args) {
  const tool = productionAttachmentValidation.buildTools().find(t => t.name === name);
  assert.ok(tool, `production tool schema missing: ${name}`);
  const schema = tool.inputSchema;
  const errors = [];
  for (const key of schema.required || []) {
    if (args[key] === undefined || args[key] === null) {
      errors.push(`Missing required parameter: ${key}`);
    }
  }
  for (const [key, value] of Object.entries(args)) {
    const propSchema = schema.properties?.[key];
    if (!propSchema) {
      errors.push(`Unknown parameter: ${key}`);
      continue;
    }
    productionAttachmentValidation.validateAgainstSchema(value, propSchema, key, errors);
  }
  return errors;
}

/**
 * Exact copy of validateToolArgs / validateAgainstSchema from api.js.
 * Kept in sync manually — if the logic in api.js changes, update here too.
 */
function validateAgainstSchema(value, schema, path, errors) {
  if (!schema || value === undefined || value === null) return;

  const expectedType = schema.type;
  if (expectedType === "array") {
    if (!Array.isArray(value)) {
      errors.push(`Parameter '${path}' must be an array, got ${typeof value}`);
      return;
    }
    if (schema.items) {
      for (let i = 0; i < value.length; i++) {
        if (value[i] === null || value[i] === undefined) {
          errors.push(`Parameter '${path}[${i}]' must not be null`);
          continue;
        }
        validateAgainstSchema(value[i], schema.items, `${path}[${i}]`, errors);
      }
    }
  } else if (expectedType === "object") {
    if (typeof value !== "object" || Array.isArray(value)) {
      errors.push(`Parameter '${path}' must be an object, got ${Array.isArray(value) ? "array" : typeof value}`);
      return;
    }
    const nestedProps = schema.properties || {};
    const nestedRequired = schema.required || [];
    for (const r of nestedRequired) {
      if (value[r] === undefined || value[r] === null) {
        errors.push(`Missing required parameter: ${path}.${r}`);
      }
    }
    for (const [k, v] of Object.entries(value)) {
      const has = Object.prototype.hasOwnProperty.call(nestedProps, k);
      if (!has) {
        if (schema.additionalProperties === false) {
          errors.push(`Unknown parameter: ${path}.${k}`);
        }
        continue;
      }
      validateAgainstSchema(v, nestedProps[k], `${path}.${k}`, errors);
    }
  } else if (expectedType === "integer") {
    if (typeof value !== "number" || !Number.isInteger(value)) {
      errors.push(`Parameter '${path}' must be an integer, got ${typeof value === "number" ? "non-integer number" : typeof value}`);
      return;
    }
  } else if (expectedType && typeof value !== expectedType) {
    errors.push(`Parameter '${path}' must be ${expectedType}, got ${typeof value}`);
    return;
  }

  if (expectedType === "string") {
    if (schema.minLength !== undefined && value.length < schema.minLength) {
      errors.push(`Parameter '${path}' must contain at least ${schema.minLength} character(s)`);
    }
    if (schema.contentEncoding === "base64" && !productionAttachmentValidation.isValidBase64(value)) {
      errors.push(`Parameter '${path}' must contain valid base64 data`);
    }
  }

  if (Array.isArray(schema.anyOf) && schema.anyOf.length > 0) {
    let matched = 0;
    const branchFailures = [];
    for (const branch of schema.anyOf) {
      const branchErrors = [];
      validateAgainstSchema(value, branch, path, branchErrors);
      if (branchErrors.length === 0) matched++;
      else branchFailures.push(branchErrors);
    }
    if (matched === 0) {
      const details = [...new Set(branchFailures.flat())].join("; ");
      errors.push(`Parameter '${path}' did not match any required schema alternative${details ? `: ${details}` : ""}`);
    }
  }

  if (Array.isArray(schema.oneOf) && schema.oneOf.length > 0) {
    let matched = 0;
    const branchFailures = [];
    for (const branch of schema.oneOf) {
      const branchErrors = [];
      validateAgainstSchema(value, branch, path, branchErrors);
      if (branchErrors.length === 0) matched++;
      else branchFailures.push(branchErrors);
    }
    if (matched === 0) {
      const details = [...new Set(branchFailures.flat())].join("; ");
      errors.push(`Parameter '${path}' did not match any allowed schema variant${details ? `: ${details}` : ""}`);
    } else if (matched > 1) {
      errors.push(`Parameter '${path}' matched more than one schema variant`);
    }
  }

  if (Array.isArray(schema.enum) && !schema.enum.includes(value)) {
    errors.push(`Parameter '${path}' must be one of ${JSON.stringify(schema.enum)}, got ${JSON.stringify(value)}`);
  }
}

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

      validateAgainstSchema(value, propSchema, key, errors);
      // Array length bounds (minItems/maxItems) sit outside validateAgainstSchema;
      // mirror the production validateToolArgs so getMessages-batch caps are tested.
      if (propSchema.type === "array" && Array.isArray(value)) {
        if (propSchema.minItems !== undefined && value.length < propSchema.minItems) {
          errors.push(`Parameter '${key}' must contain at least ${propSchema.minItems} item(s)`);
        }
        if (propSchema.maxItems !== undefined && value.length > propSchema.maxItems) {
          errors.push(`Parameter '${key}' must contain at most ${propSchema.maxItems} item(s)`);
        }
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
        attachments: {
          type: "array",
          items: {
            oneOf: [
              { type: "string" },
              {
                type: "object",
                properties: {
                  name: { type: "string" },
                  contentType: { type: "string" },
                  base64: { type: "string", minLength: 1, contentEncoding: "base64" },
                  content: { type: "string", minLength: 1, contentEncoding: "base64" },
                },
                required: ["name"],
                anyOf: [
                  { type: "object", required: ["base64"] },
                  { type: "object", required: ["content"] },
                ],
                additionalProperties: false,
              },
            ],
          },
        },
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

  // Regression: the runtime accepts `content` as an alias for `base64`
  // (entry.base64 || entry.content). Validation must not reject {name, content}.
  it('accepts inline attachment using `content` as a base64 alias', () => {
    const errors = validate('sendMail', {
      to: 'user@example.com',
      subject: 'test',
      body: 'hello',
      attachments: [{ name: 'file.pdf', content: 'AAAA' }],
    });
    assert.equal(errors.length, 0);
  });

  // Regression: contentType is optional at runtime; {name, base64} must pass.
  it('accepts inline attachment with base64 and no contentType', () => {
    const errors = validate('sendMail', {
      to: 'user@example.com',
      subject: 'test',
      body: 'hello',
      attachments: [{ name: 'file.pdf', base64: 'AAAA' }],
    });
    assert.equal(errors.length, 0);
  });

  // base64 + content together is accepted (runtime lets base64 win); validation
  // must mirror that and not reject the redundant shape.
  it('accepts inline attachment with both base64 and content set', () => {
    const errors = validate('sendMail', {
      to: 'user@example.com',
      subject: 'test',
      body: 'hello',
      attachments: [{ name: 'file.pdf', base64: 'AAAA', content: 'BBBB' }],
    });
    assert.equal(errors.length, 0);
  });

  it('rejects inline attachment missing name', () => {
    const errors = validate('sendMail', {
      to: 'user@example.com',
      subject: 'test',
      body: 'hello',
      attachments: [{ base64: 'AAAA' }],
    });
    assert.equal(errors.length, 1);
    assert.match(errors[0], /did not match any allowed schema variant/);
  });

  it('rejects inline attachment with unknown property', () => {
    const errors = validate('sendMail', {
      to: 'user@example.com',
      subject: 'test',
      body: 'hello',
      attachments: [{ name: 'file.pdf', base64: 'AAAA', evil: 'x' }],
    });
    assert.equal(errors.length, 1);
    assert.match(errors[0], /did not match any allowed schema variant/);
  });

  // Regression: validateAgainstSchema returns early on null, so a null array
  // item used to skip the item schema entirely. Must be rejected explicitly.
  it('rejects a null attachment item', () => {
    const errors = validate('sendMail', {
      to: 'user@example.com',
      subject: 'test',
      body: 'hello',
      attachments: [null],
    });
    assert.equal(errors.length, 1);
    assert.match(errors[0], /must not be null/);
  });

  const productionToolArgs = {
    sendMail: { to: 'user@example.com', subject: 'test', body: 'hello' },
    saveDraft: {},
    replyToMessage: { messageId: 'message-1', folderPath: 'imap://example/INBOX', body: 'hello' },
    forwardMessage: { messageId: 'message-1', folderPath: 'imap://example/INBOX', to: 'user@example.com' },
  };

  for (const [toolName, requiredArgs] of Object.entries(productionToolArgs)) {
    it(`production ${toolName} schema rejects payload-less and empty inline attachments`, () => {
      for (const attachment of [
        { name: 'empty.bin' },
        { name: 'empty.bin', base64: '' },
        { name: 'empty.bin', content: '' },
      ]) {
        const errors = validateProductionToolArgs(toolName, {
          ...requiredArgs,
          attachments: [attachment],
        });
        assert.ok(errors.length > 0, `${toolName} accepted ${JSON.stringify(attachment)}`);
      }
    });

    it(`production ${toolName} schema accepts non-empty base64 and content payloads`, () => {
      for (const attachment of [
        { name: 'file.bin', contentType: 'application/octet-stream', base64: 'AAAA' },
        { name: 'file.bin', contentType: 'application/octet-stream', content: 'AAAA' },
      ]) {
        const errors = validateProductionToolArgs(toolName, {
          ...requiredArgs,
          attachments: [attachment],
        });
        assert.deepEqual(errors, [], `${toolName} rejected ${JSON.stringify(attachment)}: ${errors.join('; ')}`);
      }
    });
  }

  it('production validation rejects malformed base64 with a clear error', () => {
    const errors = validateProductionToolArgs('sendMail', {
      ...productionToolArgs.sendMail,
      attachments: [{ name: 'garbage.bin', base64: '%%%%' }],
    });
    assert.ok(errors.some(error => /valid base64 data/.test(error)), errors.join('; '));
  });

  it('production runtime rejects malformed base64 before decoding it', () => {
    const result = productionAttachmentValidation.filePathsToAttachDescs([
      { name: 'garbage.bin', base64: '%%%%' },
    ]);
    assert.equal(result.descs.length, 0);
    assert.equal(result.failed.length, 1);
    assert.match(result.failed[0], /garbage\.bin \(invalid base64 data\)/);
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

// ─────────────────────────────────────────────────────────────────────────────
// Tests for the recursive validator additions: enum, oneOf branches, nested
// objects with additionalProperties:false, integer type checks. These cover
// schema keywords the previous shallow validator silently ignored.
// ─────────────────────────────────────────────────────────────────────────────

const richValidator = createValidator([
  {
    name: 'getMessage',
    inputSchema: {
      type: 'object',
      properties: {
        messageId: { type: 'string' },
        folderPath: { type: 'string' },
        bodyFormat: { type: 'string', enum: ['markdown', 'text', 'html'] },
        rawSource: { type: 'boolean' },
      },
      required: ['messageId', 'folderPath'],
    },
  },
  {
    name: 'sendMail',
    inputSchema: {
      type: 'object',
      properties: {
        to: { type: 'string' },
        subject: { type: 'string' },
        body: { type: 'string' },
        attachments: {
          type: 'array',
          items: {
            oneOf: [
              { type: 'string' },
              {
                type: 'object',
                properties: {
                  name: { type: 'string' },
                  contentType: { type: 'string' },
                  base64: { type: 'string' },
                },
                required: ['name', 'base64'],
                additionalProperties: false,
              },
            ],
          },
        },
      },
      required: ['to', 'subject', 'body'],
    },
  },
  {
    name: 'createTask',
    inputSchema: {
      type: 'object',
      properties: {
        title: { type: 'string' },
        priority: { type: 'integer' },
      },
      required: ['title'],
    },
  },
]);

describe('Validator: enum enforcement', () => {
  it('accepts an enum value that is in the list', () => {
    const errors = richValidator('getMessage', {
      messageId: 'm-1',
      folderPath: 'imap://x/INBOX',
      bodyFormat: 'markdown',
    });
    assert.equal(errors.length, 0);
  });

  it('rejects an enum value that is not in the list', () => {
    const errors = richValidator('getMessage', {
      messageId: 'm-1',
      folderPath: 'imap://x/INBOX',
      bodyFormat: 'docx',
    });
    assert.equal(errors.length, 1);
    assert.match(errors[0], /must be one of/);
    assert.match(errors[0], /bodyFormat/);
  });

  it('enum check still rejects when string type is satisfied but value is off-list', () => {
    const errors = richValidator('getMessage', {
      messageId: 'm-1',
      folderPath: 'imap://x/INBOX',
      bodyFormat: 'Markdown', // wrong case
    });
    assert.equal(errors.length, 1);
    assert.match(errors[0], /must be one of/);
  });
});

describe('Validator: oneOf items (attachments)', () => {
  it('accepts pure file-path strings', () => {
    const errors = richValidator('sendMail', {
      to: 'a@b.c',
      subject: 's',
      body: 'b',
      attachments: ['/tmp/a.txt', '/tmp/b.pdf'],
    });
    assert.equal(errors.length, 0);
  });

  it('accepts well-formed inline base64 objects', () => {
    const errors = richValidator('sendMail', {
      to: 'a@b.c',
      subject: 's',
      body: 'b',
      attachments: [{ name: 'a.txt', base64: 'aGk=' }],
    });
    assert.equal(errors.length, 0);
  });

  it('rejects attachment objects missing required fields', () => {
    const errors = richValidator('sendMail', {
      to: 'a@b.c',
      subject: 's',
      body: 'b',
      attachments: [{ contentType: 'application/pdf' }],
    });
    // both branches fail (string branch on type, object branch on missing
    // required), so oneOf records "did not match any allowed variant"
    assert.ok(errors.some(e => /did not match any allowed schema variant/.test(e)),
      `expected oneOf failure, got: ${JSON.stringify(errors)}`);
  });

  it('rejects attachment objects with extra properties (additionalProperties:false)', () => {
    const errors = richValidator('sendMail', {
      to: 'a@b.c',
      subject: 's',
      body: 'b',
      attachments: [{ name: 'a.txt', base64: 'aGk=', evil: '../../../etc/passwd' }],
    });
    assert.ok(errors.some(e => /did not match any allowed schema variant/.test(e)),
      `expected oneOf failure when extra props present, got: ${JSON.stringify(errors)}`);
  });

  it('rejects array entries that are neither strings nor valid objects (e.g. numbers)', () => {
    const errors = richValidator('sendMail', {
      to: 'a@b.c',
      subject: 's',
      body: 'b',
      attachments: [123],
    });
    assert.ok(errors.some(e => /did not match any allowed schema variant/.test(e)));
  });

  it('flags the failing index in the error path', () => {
    const errors = richValidator('sendMail', {
      to: 'a@b.c',
      subject: 's',
      body: 'b',
      attachments: ['/tmp/ok.txt', 123, '/tmp/also-ok.txt'],
    });
    assert.ok(errors.some(e => /attachments\[1\]/.test(e)),
      `expected attachments[1] path, got: ${JSON.stringify(errors)}`);
  });
});

describe('Validator: integer type', () => {
  it('accepts whole numbers for integer fields', () => {
    const errors = richValidator('createTask', { title: 't', priority: 5 });
    assert.equal(errors.length, 0);
  });

  it('rejects floats for integer fields', () => {
    const errors = richValidator('createTask', { title: 't', priority: 1.5 });
    assert.equal(errors.length, 1);
    assert.match(errors[0], /must be an integer/);
  });

  it('rejects numeric strings for integer fields (no auto-coerce here)', () => {
    const errors = richValidator('createTask', { title: 't', priority: '5' });
    assert.equal(errors.length, 1);
    assert.match(errors[0], /must be an integer/);
  });
});

// ─────────────────────────────────────────────────────────────────────────────
// isSensitiveFilePath: deny-list defending against the LLM-confused-deputy
// chain (attacker email content → assistant calls sendMail with sensitive
// attachment path). Tests execute the marked production helper from api.js and
// run cross-platform against its normalized (lower-cased, forward-slash) form.
// ─────────────────────────────────────────────────────────────────────────────

function loadProductionSensitivePathHelper() {
  const sandbox = {};
  vm.createContext(sandbox);
  vm.runInContext([
    getMarkedApiSnippet('// BEGIN SENSITIVE ATTACHMENT PATH HELPERS', '// END SENSITIVE ATTACHMENT PATH HELPERS'),
    'this.isSensitiveFilePath = isSensitiveFilePath;',
  ].join('\n'), sandbox);
  return sandbox.isSensitiveFilePath;
}

const isSensitiveFilePath = loadProductionSensitivePathHelper();

describe('isSensitiveFilePath: credential and key files', () => {
  it('blocks SSH private keys in ~/.ssh/', () => {
    assert.equal(isSensitiveFilePath('/home/user/.ssh/id_rsa'), true);
    assert.equal(isSensitiveFilePath('/Users/jordan/.ssh/id_ed25519'), true);
    assert.equal(isSensitiveFilePath('C:\\Users\\jordan\\.ssh\\id_rsa'), true);
  });

  it('blocks SSH public keys (still sensitive, contains fingerprint info)', () => {
    assert.equal(isSensitiveFilePath('/home/user/.ssh/id_rsa.pub'), true);
  });

  it('blocks the SSH directory listing itself', () => {
    assert.equal(isSensitiveFilePath('/home/user/.ssh'), true);
    assert.equal(isSensitiveFilePath('/home/user/.ssh/'), true);
  });

  it('blocks GnuPG, AWS, Azure, GCloud, kube, docker config dirs', () => {
    assert.equal(isSensitiveFilePath('/home/user/.gnupg/secring.gpg'), true);
    assert.equal(isSensitiveFilePath('/home/user/.aws/credentials'), true);
    assert.equal(isSensitiveFilePath('/home/user/.azure/accessTokens.json'), true);
    assert.equal(isSensitiveFilePath('/home/user/.config/gcloud/credentials.db'), true);
    assert.equal(isSensitiveFilePath('/home/user/.kube/config'), true);
    assert.equal(isSensitiveFilePath('/home/user/.docker/config.json'), true);
  });

  it('blocks .netrc / .npmrc / .pypirc credential files', () => {
    assert.equal(isSensitiveFilePath('/home/user/.netrc'), true);
    assert.equal(isSensitiveFilePath('/home/user/.npmrc'), true);
    assert.equal(isSensitiveFilePath('/home/user/.pypirc'), true);
  });

  it('blocks PEM/PFX/P12/KDBX/KEY/ASC/GPG files anywhere on disk', () => {
    assert.equal(isSensitiveFilePath('/tmp/server.pem'), true);
    assert.equal(isSensitiveFilePath('/home/user/wildcard.pfx'), true);
    assert.equal(isSensitiveFilePath('/data/cert.p12'), true);
    assert.equal(isSensitiveFilePath('/Users/x/Passwords.kdbx'), true);
    assert.equal(isSensitiveFilePath('/etc/ssl/private.key'), true);
    assert.equal(isSensitiveFilePath('/home/user/key.asc'), true);
    assert.equal(isSensitiveFilePath('/home/user/secret.gpg'), true);
  });
});

describe('isSensitiveFilePath: system directories', () => {
  it('blocks Linux/macOS system dirs', () => {
    assert.equal(isSensitiveFilePath('/etc/shadow'), true);
    assert.equal(isSensitiveFilePath('/etc/passwd'), true);
    assert.equal(isSensitiveFilePath('/proc/self/environ'), true);
    assert.equal(isSensitiveFilePath('/sys/class/net/eth0/address'), true);
    assert.equal(isSensitiveFilePath('/root/.bash_history'), true);
    assert.equal(isSensitiveFilePath('/var/log/auth.log'), true);
    assert.equal(isSensitiveFilePath('/var/lib/sudo/lectured/user'), true);
  });

  it('blocks macOS keychains', () => {
    assert.equal(isSensitiveFilePath('/Users/x/Library/Keychains/login.keychain-db'), true);
  });

  it('blocks Windows system dirs (forward and back slashes)', () => {
    assert.equal(isSensitiveFilePath('C:\\Windows\\System32\\config\\SAM'), true);
    assert.equal(isSensitiveFilePath('C:/Windows/System32/config/SAM'), true);
    assert.equal(isSensitiveFilePath('D:/Windows/System32/notepad.exe'), true);
  });

  it('blocks Windows DPAPI / credential vault locations', () => {
    assert.equal(isSensitiveFilePath('C:\\ProgramData\\Microsoft\\Crypto\\RSA\\MachineKeys\\x'), true);
    assert.equal(isSensitiveFilePath('C:\\Users\\jordan\\AppData\\Local\\Microsoft\\Credentials\\x'), true);
    assert.equal(isSensitiveFilePath('C:\\Users\\jordan\\AppData\\Roaming\\Microsoft\\Vault'), true);
  });
});

describe('isSensitiveFilePath: browser and mail data', () => {
  it('blocks browser credential stores', () => {
    assert.equal(isSensitiveFilePath('/home/user/.mozilla/firefox/abc.default/logins.json'), true);
    assert.equal(isSensitiveFilePath('/home/user/.mozilla/firefox/abc.default/key4.db'), true);
    assert.equal(isSensitiveFilePath('/home/user/.config/google-chrome/Default/Cookies'), true);
    assert.equal(
      isSensitiveFilePath('C:\\Users\\x\\AppData\\Local\\Google\\Chrome\\User Data\\Default\\Login Data'),
      true
    );
  });

  it('blocks real Linux Thunderbird profile layouts and root files', () => {
    assert.equal(isSensitiveFilePath('/home/user/.thunderbird/abc.default-release/Mail/Local Folders'), true);
    assert.equal(isSensitiveFilePath('/home/user/.thunderbird/abc.default-release/prefs.js'), true);
    assert.equal(isSensitiveFilePath('/home/user/.thunderbird/abc.default-release/key4.db'), true);
    assert.equal(isSensitiveFilePath('/home/user/.thunderbird/abc.default-release/logins.json'), true);
    assert.equal(isSensitiveFilePath('/home/user/.thunderbird/profiles.ini'), true);
    assert.equal(isSensitiveFilePath('/home/user/.thunderbird'), true);
  });

  it('blocks Icedove and platform-specific Thunderbird profile roots', () => {
    assert.equal(isSensitiveFilePath('/home/user/.icedove/abc.default/prefs.js'), true);
    assert.equal(isSensitiveFilePath('/Users/x/Library/Thunderbird/Profiles/abc/INBOX'), true);
    assert.equal(isSensitiveFilePath('C:\\Users\\jordan\\AppData\\Roaming\\Thunderbird\\Profiles\\abc'), true);
  });
});

describe('isSensitiveFilePath: benign paths pass through', () => {
  it('allows typical user documents and downloads', () => {
    assert.equal(isSensitiveFilePath('/home/user/Documents/report.pdf'), false);
    assert.equal(isSensitiveFilePath('/home/user/Downloads/photo.jpg'), false);
    assert.equal(isSensitiveFilePath('C:\\Users\\jordan\\Downloads\\invoice.xlsx'), false);
    assert.equal(isSensitiveFilePath('/tmp/scratch.txt'), false);
  });

  it('allows files whose names merely contain substrings of patterns', () => {
    // "ssh" inside a filename is not the /.ssh/ directory boundary
    assert.equal(isSensitiveFilePath('/home/user/notes/ssh-cheatsheet.md'), false);
    // .pemphigus is not .pem
    assert.equal(isSensitiveFilePath('/home/user/medical/pemphigus.txt'), false);
    // "etc" inside a path that doesn't start at /etc/
    assert.equal(isSensitiveFilePath('/home/user/etc-notes.md'), false);
    // Profile-root names require a full path-component boundary.
    assert.equal(isSensitiveFilePath('/home/user/.thunderbird-notes/report.txt'), false);
  });

  it('returns false on non-string / empty input rather than throwing', () => {
    assert.equal(isSensitiveFilePath(''), false);
    assert.equal(isSensitiveFilePath(null), false);
    assert.equal(isSensitiveFilePath(undefined), false);
    assert.equal(isSensitiveFilePath(123), false);
    assert.equal(isSensitiveFilePath({}), false);
  });
});

describe('isSensitiveFilePath: case insensitivity and slash normalization', () => {
  it('matches regardless of letter case', () => {
    assert.equal(isSensitiveFilePath('/HOME/USER/.SSH/ID_RSA'), true);
    assert.equal(isSensitiveFilePath('/Etc/Shadow'), true);
    assert.equal(isSensitiveFilePath('C:\\WINDOWS\\system32\\drivers'), true);
  });

  it('treats backslashes and forward slashes as equivalent boundaries', () => {
    assert.equal(isSensitiveFilePath('C:/Users/x/.ssh/id_rsa'), true);
    assert.equal(isSensitiveFilePath('C:\\Users\\x\\.ssh\\id_rsa'), true);
  });
});
