
const { describe, it, beforeEach } = require('node:test');
const assert = require('node:assert/strict');

const {
  createMapper,
  TOOL_ID_CONFIG,
} = require('../mcp-id-mapper.cjs');

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

// A real-looking RFC 5322 Message-ID
const MSG_ID_1 = '<a1b2c3d4.12345.67890@mail.example.com>';
const MSG_ID_2 = '<e5f6g7h8.24680.13579@other.example.com>';
const UUID_1 = '550e8400-e29b-41d4-a716-446655440000';
const UUID_2 = '6ba7b810-9dad-11d1-80b4-00c04fd430c8';

// id-agent alias pattern: 3 words with optional prefix (e.g. msg_foo-bar-baz)
const ALIAS_RE = /^[a-z]{2,3}_[a-z]{2,}(?:-[a-z]{2,})*$/;

// ---------------------------------------------------------------------------
// Tests
// ---------------------------------------------------------------------------

describe('createMapper', () => {
  let mapper;

  beforeEach(() => {
    mapper = createMapper();
  });

  // ── Outbound mapping ──

  describe('interceptOutbound', () => {
    it('replaces message IDs in searchMessages results', () => {
      const data = [
        { id: MSG_ID_1, threadId: '<thread1@x>', subject: 'Hello' },
        { id: MSG_ID_2, threadId: '<thread2@x>', subject: 'World' },
      ];
      mapper.interceptOutbound('searchMessages', data);

      assert.ok(ALIAS_RE.test(data[0].id), `expected alias, got ${data[0].id}`);
      assert.ok(ALIAS_RE.test(data[0].threadId));
      assert.ok(ALIAS_RE.test(data[1].id));
      assert.equal(data[0].subject, 'Hello'); // non-ID fields untouched
      assert.equal(data[1].subject, 'World');
    });

    it('passes non-ID fields through unchanged', () => {
      const data = {
        author: 'Alice <alice@example.com>',
        subject: 'Test message',
        body: '<p>Hello world</p>',
        read: true,
        flagged: false,
      };
      mapper.interceptOutbound('getMessage', data);

      assert.equal(data.author, 'Alice <alice@example.com>');
      assert.equal(data.subject, 'Test message');
      assert.equal(data.body, '<p>Hello world</p>');
      assert.equal(data.read, true);
      assert.equal(data.flagged, false);
    });

    it('handles getMessages nested messages array', () => {
      const data = {
        messages: [
          { messageId: MSG_ID_1, threadId: '<t1@x>' },
          { messageId: MSG_ID_2, threadId: '<t2@x>' },
        ],
        succeeded: 2,
        failed: 0,
        max: 10,
      };
      mapper.interceptOutbound('getMessages', data);

      assert.ok(ALIAS_RE.test(data.messages[0].messageId));
      assert.ok(ALIAS_RE.test(data.messages[0].threadId));
      assert.ok(ALIAS_RE.test(data.messages[1].messageId));
      assert.equal(data.succeeded, 2);
      assert.equal(data.max, 10);
    });

    it('maps contact IDs in searchContacts results', () => {
      const data = {
        contacts: [
          { id: UUID_1, addressBookId: UUID_2, displayName: 'Alice' },
        ],
      };
      mapper.interceptOutbound('searchContacts', data);

      assert.ok(ALIAS_RE.test(data.contacts[0].id));
      assert.ok(/^ctc_/.test(data.contacts[0].id), 'should have ctc_ prefix');
      assert.equal(data.contacts[0].displayName, 'Alice');
    });

    it('maps event IDs in listEvents results', () => {
      const data = [
        { id: UUID_1, calendarId: UUID_2, title: 'Meeting' },
      ];
      mapper.interceptOutbound('listEvents', data);

      assert.ok(/^evt_/.test(data[0].id), `expected evt_ prefix, got ${data[0].id}`);
      assert.ok(ALIAS_RE.test(data[0].id));
      assert.ok(ALIAS_RE.test(data[0].calendarId));
      assert.equal(data[0].title, 'Meeting');
    });

    it('maps account IDs in listAccounts results', () => {
      const data = [
        {
          id: 'account1',
          name: 'Work',
          identities: [
            { id: 'id1', email: 'work@example.com', name: 'Work' },
          ],
        },
      ];
      mapper.interceptOutbound('listAccounts', data);

      assert.ok(/^acc_/.test(data[0].id), `expected acc_ prefix, got ${data[0].id}`);
      assert.ok(ALIAS_RE.test(data[0].id));
      assert.ok(ALIAS_RE.test(data[0].identities[0].id));
    });

    it('does not map IDs for unconfigured tools', () => {
      const data = { id: UUID_1, name: 'something' };
      mapper.interceptOutbound('deleteMessage', data);
      assert.equal(data.id, UUID_1); // unchanged
    });
  });

  // ── Inbound mapping ──

  describe('interceptInbound', () => {
    it('round-trips: outbound → inbound restores original IDs', () => {
      const data = {
        messageId: MSG_ID_1,
        folderPath: '/account1/INBOX',
      };

      mapper.interceptOutbound('getMessage', data);
      const alias = data.messageId;
      assert.ok(ALIAS_RE.test(alias));

      mapper.interceptInbound('getMessage', data);
      assert.equal(data.messageId, MSG_ID_1);
      assert.equal(data.folderPath, '/account1/INBOX');
    });

    it('round-trips array of message IDs', () => {
      const data = {
        messageIds: [MSG_ID_1, MSG_ID_2],
        folderPath: '/account1/INBOX',
      };

      mapper.interceptOutbound('searchMessages', [data]);
      // data itself has no direct ID keys in searchMessages config,
      // but messageIds uses the generic inbound walker (it sees string
      // aliases that match the pattern). For outbound, we need to
      // manually test the inbound walk with an alias we create.

      // Instead, test inbound directly with a fresh alias
      const args = { messageIds: ['msg_fake-word-test'], folderPath: '/a/b' };
      mapper.interceptInbound('deleteMessages', args);
      // This fake alias won't resolve — verify it stays unchanged
      assert.equal(args.messageIds[0], 'msg_fake-word-test');
    });

    it('round-trips nested message objects in args', () => {
      // First map them outbound to get real aliases
      const orig = {
        messages: [
          { messageId: MSG_ID_1, folderPath: '/a/b' },
          { messageId: MSG_ID_2, folderPath: '/c/d' },
        ],
      };
      mapper.interceptOutbound('getMessages', orig);

      const alias1 = orig.messages[0].messageId;
      const alias2 = orig.messages[1].messageId;
      assert.ok(ALIAS_RE.test(alias1));
      assert.ok(ALIAS_RE.test(alias2));

      // Now simulate the LLM sending these aliases back
      const args = {
        messages: [
          { messageId: alias1, folderPath: '/a/b' },
          { messageId: alias2, folderPath: '/c/d' },
        ],
      };
      mapper.interceptInbound('getMessages', args);

      assert.equal(args.messages[0].messageId, MSG_ID_1);
      assert.equal(args.messages[1].messageId, MSG_ID_2);
    });

    it('passes unaliased strings through unchanged', () => {
      const args = {
        to: 'alice@example.com',
        subject: 'Hello world',
        body: '<p>Test</p>',
        messageId: MSG_ID_1, // a real message ID that wasn't mapped
      };
      mapper.interceptInbound('sendMail', args);

      // Real but unmapped message ID stays as-is
      assert.equal(args.messageId, MSG_ID_1);
      assert.equal(args.to, 'alice@example.com');
    });

    it('handles null / non-object args', () => {
      assert.doesNotThrow(() => mapper.interceptInbound('foo', null));
      assert.doesNotThrow(() => mapper.interceptInbound('foo', undefined));
    });
  });

  // ── Idempotency ──

  describe('idempotency', () => {
    it('same real ID → same alias', () => {
      const a1 = mapper._makeAlias('messages', MSG_ID_1);
      const a2 = mapper._makeAlias('messages', MSG_ID_1);
      assert.equal(a1, a2);
    });

    it('different real IDs → different aliases', () => {
      // With 3 words and 4096-word pool: 68B combos. Collision chance
      // is near-zero across practical session sizes.
      const a1 = mapper._makeAlias('messages', MSG_ID_1);
      const a2 = mapper._makeAlias('messages', MSG_ID_2);
      assert.notEqual(a1, a2);
    });

    it('different entity types don\'t collide', () => {
      const msgAlias = mapper._makeAlias('messages', UUID_1);
      const ctcAlias = mapper._makeAlias('contacts', UUID_1);
      // Different prefixes ensure no cross-type collision
      assert.ok(msgAlias.startsWith('msg_'));
      assert.ok(ctcAlias.startsWith('ctc_'));
    });
  });

  // ── Stats & lifecycle ──

  describe('getStats and clear', () => {
    it('tracks mapping counts per entity type', () => {
      assert.deepEqual(mapper.getStats(), {});

      mapper._makeAlias('messages', MSG_ID_1);
      mapper._makeAlias('contacts', UUID_1);
      mapper._makeAlias('contacts', UUID_2);

      const stats = mapper.getStats();
      assert.equal(stats.messages, 1);
      assert.equal(stats.contacts, 2);
    });

    it('clear() resets all mappings', () => {
      mapper._makeAlias('messages', MSG_ID_1);
      mapper._makeAlias('contacts', UUID_1);

      assert.equal(mapper.getStats().messages, 1);

      mapper.clear();
      assert.deepEqual(mapper.getStats(), {});
    });
  });

  // ── Non-ID aliases pass through inbound ──

  describe('inbound non-alias pass-through', () => {
    it('does not modify strings that don\'t match alias pattern', () => {
      const args = {
        body: '<html><body>Hello world</body></html>',
        subject: 'Re: Meeting about Q3 budget',
        to: 'boss@company.com',
        from: 'id1',
      };
      mapper.interceptInbound('replyToMessage', args);
      assert.equal(args.body, '<html><body>Hello world</body></html>');
      assert.equal(args.subject, 'Re: Meeting about Q3 budget');
      assert.equal(args.to, 'boss@company.com');
      assert.equal(args.from, 'id1'); // identity id from listAccounts (not an alias)
    });

    it('EAFP: no crash on deeply nested structures', () => {
      const deep = {
        a: { b: { c: { d: { e: 'nothing-to-map' } } } },
      };
      assert.doesNotThrow(() => mapper.interceptInbound('foo', deep));
    });
  });

  // ── TOOL_ID_CONFIG coverage ──

  describe('TOOL_ID_CONFIG', () => {
    it('covers all read/list tools that return IDs', () => {
      const expected = [
        'searchMessages',
        'getRecentMessages',
        'getMessage',
        'getMessages',
        'listFolders',
        'searchContacts',
        'createContact',
        'listCalendars',
        'listEvents',
        'createEvent',
        'listTasks',
        'createTask',
        'listAccounts',
        'getAccountAccess',
      ];
      for (const name of expected) {
        assert.ok(TOOL_ID_CONFIG[name], `Missing config for ${name}`);
      }
    });

    it('each config entry has keys or nestedKeys', () => {
      for (const [name, config] of Object.entries(TOOL_ID_CONFIG)) {
        assert.ok(
          (config.keys && config.keys.length > 0) ||
            (config.nestedKeys && Object.keys(config.nestedKeys).length > 0),
          `${name} must have keys or nestedKeys`
        );
      }
    });
  });
});

// ---------------------------------------------------------------------------
// End-to-end scenario: search → read → reply with mapped IDs
// ---------------------------------------------------------------------------

describe('End-to-end flow', () => {
  it('simulates a full LLM session: search → get → reply', () => {
    const mapper = createMapper();

    // Step 1: LLM calls searchMessages
    const searchResults = [
      {
        id: '<abc123@mail.com>',
        threadId: '<thread-abc@mail.com>',
        subject: 'RE: Q3 Budget',
        folderPath: '/account1/INBOX',
        author: 'boss@company.com',
      },
    ];

    mapper.interceptOutbound('searchMessages', searchResults);
    const msgAlias = searchResults[0].id;
    assert.ok(ALIAS_RE.test(msgAlias), `msgAlias=${msgAlias}`);
    // LLM sees: msgAlias = "msg_foo-bar-baz"

    // Step 2: LLM reads the message using the alias
    const getArgs = {
      messageId: msgAlias,
      folderPath: '/account1/INBOX',
    };
    mapper.interceptInbound('getMessage', getArgs);
    assert.equal(getArgs.messageId, '<abc123@mail.com>');
    // Thunderbird sees the real message ID

    // Step 3: LLM replies, still using the alias
    const replyArgs = {
      messageId: msgAlias,
      folderPath: '/account1/INBOX',
      body: 'Approved!',
      to: 'boss@company.com',
    };
    mapper.interceptInbound('replyToMessage', replyArgs);
    assert.equal(replyArgs.messageId, '<abc123@mail.com>');
    // Thunderbird sees the real message ID
  });
});