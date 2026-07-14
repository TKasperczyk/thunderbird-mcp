'use strict';

/**
 * MCP ID Mapper — Surrogate ID Translation Layer
 *
 * Replaces long, token-wasteful real IDs (RFC 5322 Message-IDs, UUIDs)
 * with short 1-token-per-word surrogate IDs from id-agent.
 *
 * Outbound (Thunderbird → LLM):  replace real IDs with short aliases
 * Inbound  (LLM → Thunderbird):  replace short aliases with real IDs
 *
 * Per entity type, a separate AliasMap avoids cross-type collisions.
 * Disable with THUNDERBIRD_MCP_DISABLE_ID_MAPPING=1.
 */

const { createAliasMap } = require('id-agent');

// ---------------------------------------------------------------------------
// Configuration — which JSON keys to remap, per tool
// ---------------------------------------------------------------------------

/**
 * For each tool that produces entities with IDs, define:
 *   keys: string[]        — ID-bearing field names in the response JSON
 *   nestedKeys: object    — field → child-key mappings for array fields
 *
 * Only tools listed here have their responses mapped.
 * Inbound args are mapped for any tool whose args carry ID keys.
 */
const TOOL_ID_CONFIG = {
  // ── Messages ──
  searchMessages: {
    keys: ['id', 'threadId', 'dupLocations'],
    nestedKeys: { dupLocations: ['id'] },
  },
  getRecentMessages: {
    keys: ['id', 'threadId'],
  },
  getMessage: {
    keys: ['id', 'threadId'],
  },
  getMessages: {
    keys: ['messageId', 'threadId'],
    nestedKeys: { messages: ['messageId', 'threadId'] },
  },
  // ── Folders ──
  listFolders: {
    keys: ['accountId'],
  },
  // ── Contacts ──
  searchContacts: {
    keys: ['id', 'addressBookId'],
    nestedKeys: { contacts: ['id', 'addressBookId'] },
  },
  createContact: {
    keys: ['id', 'addressBookId'],
  },
  // ── Calendar ──
  listCalendars: {
    keys: ['id'],
  },
  listEvents: {
    keys: ['id', 'calendarId'],
  },
  createEvent: {
    keys: ['id', 'calendarId'],
  },
  listTasks: {
    keys: ['id', 'calendarId'],
  },
  createTask: {
    keys: ['id', 'calendarId'],
  },
  // ── Accounts ──
  listAccounts: {
    keys: ['id'],
    nestedKeys: { identities: ['id'] },
  },
  getAccountAccess: {
    keys: ['id'],
  },
};

/**
 * Entity-type to id-agent prefix mapping.
 * Each entity type gets its own AliasMap. Short prefixes reserve
 * as few tokens as possible while providing self-documenting type
 * information and preventing cross-type collisions.
 *
 *   msg = message, ctc = contact, evt = event, tsk = task,
 *   cal = calendar, fld = folder, acc = account
 */
const TYPE_PREFIX = {
  messages: 'msg',
  contacts: 'ctc',
  events: 'evt',
  tasks: 'tsk',
  calendars: 'cal',
  folders: 'fld',
  accounts: 'acc',
};

// Field name → entity type (for unambiguous keys)
const FIELD_TO_TYPE = Object.freeze({
  messageId: 'messages',
  threadId: 'messages',
  contactId: 'contacts',
  eventId: 'events',
  taskId: 'tasks',
  calendarId: 'calendars',
  addressBookId: 'contacts',
  accountId: 'accounts',
});

// Tool name → dominant entity type (for ambiguous 'id' fields)
const TOOL_TO_DOMINANT_TYPE = Object.freeze({
  searchMessages: 'messages',
  getRecentMessages: 'messages',
  getMessage: 'messages',
  getMessages: 'messages',
  searchContacts: 'contacts',
  createContact: 'contacts',
  listCalendars: 'calendars',
  listEvents: 'events',
  createEvent: 'events',
  listTasks: 'tasks',
  createTask: 'tasks',
  listAccounts: 'accounts',
  getAccountAccess: 'accounts',
  listFolders: 'folders',
});

// ── Map creation ──

function createMapper() {
  /** Maps entity-type → AliasMap */
  const maps = {};
  /** Short alias → entity type (for inbound lookup when prefix stripped) */
  const aliasToType = new Map();

  function getMap(type) {
    if (!maps[type]) {
      maps[type] = createAliasMap({ words: 3 });
    }
    return maps[type];
  }

  /**
   * Builds the alias for a real ID: prefix_word1-word2-word3
   */
  function makeAlias(type, realId) {
    const prefix = TYPE_PREFIX[type];
    const aliasMap = getMap(type);
    const bareAlias = aliasMap.set(realId);
    const full = prefix ? `${prefix}_${bareAlias}` : bareAlias;
    aliasToType.set(full, type);
    return full;
  }

  /**
   * Extract the real ID from an alias prefix_word1-word2-word3.
   *
   * aliasMap stores:  realId → bare-alias (e.g., "attern-agre-divide").
   * makeAlias returns:  `${prefix}_${bare-alias}` (e.g., "msg_attern-agre-divide").
   *
   * On inbound: we receive the full prefixed alias, strip the prefix,
   * and look up the bare alias in the appropriate AliasMap.
   */
  function resolveAlias(alias) {
    // Try aliasToType for a direct lookup first
    const type = aliasToType.get(alias);
    if (type && maps[type]) {
      const bare = alias.startsWith(`${TYPE_PREFIX[type]}_`)
        ? alias.slice(TYPE_PREFIX[type].length + 1)
        : alias;
      return maps[type].get(bare);
    }
    // Fallback: strip known prefix and try all maps
    for (const t of Object.keys(TYPE_PREFIX)) {
      const pfx = TYPE_PREFIX[t];
      if (!pfx) continue;
      if (alias.startsWith(`${pfx}_`)) {
        const bare = alias.slice(pfx.length + 1);
        if (maps[t]) {
          const resolved = maps[t].get(bare);
          if (resolved !== undefined) return resolved;
        }
      }
    }
    return undefined;
  }

  /**
   * Determine the entity type for a value based on surrounding context.
   * Returns null when the value doesn't look like an ID we should map.
   */
  function inferType(key, value, contextToolName) {
    if (typeof value !== 'string' || value.length < 3) return null;

    // 1. Direct field-name mapping
    const directType = FIELD_TO_TYPE[key];
    if (directType) return directType;

    // 2. Ambiguous 'id' — infer from tool context
    if (key === 'id' && TOOL_TO_DOMINANT_TYPE[contextToolName]) {
      return TOOL_TO_DOMINANT_TYPE[contextToolName];
    }

    // 3. 'ids' plural (not yet used but future-proof)
    if (key === 'ids') return null;

    return null;
  }

  // ── Deep JSON walker ──

  /**
   * Walk a parsed JSON value (object, array, or scalar), replacing
   * real IDs with short aliases. Mutates in place.
   *
   * Handles:
   *   - Direct fields: { messageId: "<long@id>" }
   *   - Array fields: { messages: [{ messageId: "..." }, ...] }
   *   - Nested keys: specified per-tool in TOOL_ID_CONFIG
   */
  function walkOutbound(toolName, value, path = '') {
    if (value === null || typeof value !== 'object') return;

    if (Array.isArray(value)) {
      for (let i = 0; i < value.length; i++) {
        walkOutbound(toolName, value[i], `${path}[${i}]`);
      }
      return;
    }

    for (const key of Object.keys(value)) {
      const child = value[key];
      const newPath = path ? `${path}.${key}` : key;

      // If this key holds an array of sub-objects, check nestedKeys
      const config = TOOL_ID_CONFIG[toolName];
      if (Array.isArray(child)) {
        if (config && config.nestedKeys && config.nestedKeys[key]) {
          // Array of entities — walk each with the nested key list
          for (let i = 0; i < child.length; i++) {
            if (child[i] && typeof child[i] === 'object' && !Array.isArray(child[i])) {
              for (const nk of config.nestedKeys[key]) {
                if (typeof child[i][nk] === 'string') {
                  const type = inferType(nk, child[i][nk], toolName);
                  if (type) {
                    child[i][nk] = makeAlias(type, child[i][nk]);
                  }
                }
              }
            }
          }
          continue;
        }
        // Plain array (not nestedKeys) — walk normally
        walkOutbound(toolName, child, newPath);
        continue;
      }

      // Scalar or object — map scalar IDs, recurse into objects
      if (typeof child === 'string') {
        const type = inferType(key, child, toolName);
        if (type) {
          value[key] = makeAlias(type, child);
        }
      } else if (child !== null && typeof child === 'object') {
        walkOutbound(toolName, child, newPath);
      }
    }
  }

  /**
   * Reverse walk — replace short aliases with real IDs before
   * forwarding args to Thunderbird. Only touches string fields
   * that look like id-agent aliases (lowercase, hyphenated words).
   */
  function walkInbound(value, path = '') {
    if (value === null || typeof value !== 'object') return;

    if (Array.isArray(value)) {
      for (let i = 0; i < value.length; i++) {
        walkInbound(value[i], `${path}[${i}]`);
      }
      return;
    }

    for (const key of Object.keys(value)) {
      const child = value[key];

      if (typeof child === 'string') {
        // Check if this string looks like an id-agent alias
        if (/^[a-z]{2,3}_[a-z]{2,}(?:-[a-z]{2,})*$/.test(child)) {
          const real = resolveAlias(child);
          if (real !== undefined) {
            value[key] = real;
          }
        }
      } else if (Array.isArray(child)) {
        for (let i = 0; i < child.length; i++) {
          if (typeof child[i] === 'string') {
            if (/^[a-z]{2,3}_[a-z]{2,}(?:-[a-z]{2,})*$/.test(child[i])) {
              const real = resolveAlias(child[i]);
              if (real !== undefined) {
                child[i] = real;
              }
            }
          } else if (child[i] !== null && typeof child[i] === 'object') {
            walkInbound(child[i], `${path}[${i}]`);
          }
        }
      } else if (child !== null && typeof child === 'object') {
        walkInbound(child, path ? `${path}.${key}` : key);
      }
    }
  }

  // ── Public API ──

  return {
    /**
     * interceptOutbound(toolName, data)
     *
     * Replace all real IDs in a tool result with short aliases.
     * `data` is the parsed JSON result from the Thunderbird extension
     * (the content of the `text` field inside a tool response).
     * Mutates `data` in place; returns `data` for chaining.
     */
    interceptOutbound(toolName, data) {
      if (!TOOL_ID_CONFIG[toolName]) return data;
      walkOutbound(toolName, data);
      return data;
    },

    /**
     * interceptInbound(toolName, args)
     *
     * Replace all short aliases in tool arguments with real IDs
     * before forwarding to Thunderbird. Mutates in place.
     */
    interceptInbound(toolName, args) {
      if (typeof args !== 'object' || args === null) return args;
      walkInbound(args);
      return args;
    },

    /**
     * getStats() → { type: count, ... }
     */
    getStats() {
      const stats = {};
      for (const type of Object.keys(maps)) {
        stats[type] = maps[type].size;
      }
      return stats;
    },

    /** Clear all mappings (useful between tests) */
    clear() {
      for (const key of Object.keys(maps)) {
        delete maps[key];
      }
      aliasToType.clear();
    },

    // Exposed for testing
    _maps: maps,
    _aliasToType: aliasToType,
    _makeAlias: makeAlias,
    _resolveAlias: resolveAlias,
  };
}

const DISABLED = process.env.THUNDERBIRD_MCP_DISABLE_ID_MAPPING === '1';

/**
 * Singleton mapper instance. Created lazily so the module can be
 * loaded in test environments without id-agent disk access issues.
 */
let _mapper = null;
function getMapper() {
  if (DISABLED) return null;
  if (!_mapper) {
    _mapper = createMapper();
  }
  return _mapper;
}

module.exports = {
  createMapper,
  getMapper,
  TOOL_ID_CONFIG,
  TYPE_PREFIX,
  FIELD_TO_TYPE,
  TOOL_TO_DOMINANT_TYPE,
};