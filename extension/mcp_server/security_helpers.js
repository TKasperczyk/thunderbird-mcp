/**
 * Shared pure helpers used by both the extension (api.js) and the Node test
 * suite. No Mozilla XPCOM dependencies, no Thunderbird globals; everything in
 * this file must be evaluable from a plain Node.js context so the tests can
 * exercise the same code that ships in production instead of maintaining
 * parallel re-implementations that drift.
 *
 * Loading
 * -------
 * Inside the extension, api.js loads this file with
 *
 *     const helperScope = { module: { exports: {} } };
 *     Services.scriptloader.loadSubScript(
 *       "resource://thunderbird-mcp/mcp_server/security_helpers.js",
 *       helperScope
 *     );
 *     const helpers = helperScope.module.exports;
 *
 * In the test suite this file is `require()`-d directly because the
 * `module.exports = { ... }` at the bottom is standard CommonJS.
 */
'use strict';

// ─── Attachment deny-list (S1b) ──────────────────────────────────────────────

const SENSITIVE_ATTACHMENT_PATTERNS = [
  /\/\.ssh(\/|$)/,
  /\/\.gnupg(\/|$)/,
  /\/\.aws(\/|$)/,
  /\/\.azure(\/|$)/,
  /\/\.config\/gcloud(\/|$)/,
  /\/\.kube(\/|$)/,
  /\/\.docker(\/|$)/,
  /\/\.netrc$/,
  /\/\.npmrc$/,
  /\/\.pypirc$/,
  /\/id_(rsa|dsa|ecdsa|ed25519)(\.pub)?$/,
  /\.pem$/,
  /\.pfx$/,
  /\.p12$/,
  /\.kdbx$/,
  /\.key$/,
  /\.asc$/,
  /\.gpg$/,
  /^\/etc\//,
  /^\/proc\//,
  /^\/sys\//,
  /^\/root\//,
  /^\/var\/log\//,
  /^\/var\/lib\/sudo\//,
  /\/library\/keychains\//,
  /^[a-z]:\/windows\//,
  /^[a-z]:\/programdata\/microsoft\/(crypto|protect)\//,
  /\/appdata\/(local|roaming)\/microsoft\/(credentials|crypto|protect|vault)(\/|$)/,
  /\/(logins\.json|key3\.db|key4\.db|cookies(\.sqlite)?|login data)$/,
  /\/\.?thunderbird\/profiles?(\/|$)/,
  /\/library\/thunderbird(\/|$)/,
  /\/appdata\/roaming\/thunderbird(\/|$)/,
];

function isSensitiveFilePath(attachmentPath) {
  if (typeof attachmentPath !== 'string' || !attachmentPath) return false;
  const normalized = attachmentPath.replace(/\\/g, '/').toLowerCase();
  return SENSITIVE_ATTACHMENT_PATTERNS.some(re => re.test(normalized));
}

// ─── Single-line header sanitization (F2) ────────────────────────────────────

function sanitizeHeaderLine(value) {
  if (typeof value !== 'string') return value;
  return value.replace(/[\r\n\0]+/g, ' ');
}

// ─── Untrusted-content delimiters for email bodies / previews ────────────────

const UNTRUSTED_OPEN = '<untrusted_email_body>';
const UNTRUSTED_CLOSE = '</untrusted_email_body>';
// Zero-width-space inserted into a defanged close marker so a sender-embedded
// `</untrusted_email_body>` cannot terminate our wrap and escape into the
// instruction layer of the consuming LLM.
const UNTRUSTED_CLOSE_DEFANGED = '</untrusted_email_body​>';

function wrapUntrustedBody(text) {
  if (typeof text !== 'string' || !text) return text;
  const safe = text.split(UNTRUSTED_CLOSE).join(UNTRUSTED_CLOSE_DEFANGED);
  return `${UNTRUSTED_OPEN}\n${safe}\n${UNTRUSTED_CLOSE}`;
}

function wrapUntrustedPreview(text) {
  if (typeof text !== 'string' || !text) return text;
  const safe = text.split(UNTRUSTED_CLOSE).join(UNTRUSTED_CLOSE_DEFANGED);
  return `${UNTRUSTED_OPEN} ${safe} ${UNTRUSTED_CLOSE}`;
}

// ─── Audit-log helpers ───────────────────────────────────────────────────────

function countRecipients(s) {
  if (typeof s !== 'string' || !s.trim()) return 0;
  return s.split(',').map(p => p.trim()).filter(Boolean).length;
}

function summarizeAttachmentsForAudit(descs) {
  if (!Array.isArray(descs)) return { count: 0, names: [], totalBytes: 0 };
  let total = 0;
  const names = [];
  for (const d of descs) {
    if (d && typeof d.size === 'number') total += d.size;
    if (d && d.name) names.push(String(d.name).slice(0, 200));
  }
  return { count: descs.length, names, totalBytes: total };
}

// ─── Markdown URL allow-list + injection escape (F4 / F5) ────────────────────

const SAFE_HREF_SCHEMES = new Set(['http', 'https', 'mailto', 'tel', 'cid', 'ftp', 'ftps']);

function isSafeMarkdownHref(url) {
  if (typeof url !== 'string') return false;
  const trimmed = url.trim();
  if (!trimmed) return false;
  const cleaned = trimmed.replace(/^[\s -]+/, '');
  const colon = cleaned.indexOf(':');
  const slash = cleaned.indexOf('/');
  const question = cleaned.indexOf('?');
  const hash = cleaned.indexOf('#');
  if (colon === -1) return true;
  if (slash !== -1 && slash < colon) return true;
  if (question !== -1 && question < colon) return true;
  if (hash !== -1 && hash < colon) return true;
  const scheme = cleaned.slice(0, colon).toLowerCase();
  return SAFE_HREF_SCHEMES.has(scheme);
}

function isSafeImageSrc(url) {
  if (isSafeMarkdownHref(url)) return true;
  if (typeof url !== 'string') return false;
  const cleaned = url.trim().replace(/^[\s -]+/, '').toLowerCase();
  return cleaned.startsWith('data:image/');
}

function escapeMarkdownLinkText(s) {
  return String(s).replace(/[\[\]]/g, m => (m === '[' ? '\\[' : '\\]'));
}

function renderMarkdownLink(text, url) {
  const safeText = escapeMarkdownLinkText(text);
  if (url.includes('>')) return safeText;
  if (/[()\s]/.test(url)) return `[${safeText}](<${url}>)`;
  return `[${safeText}](${url})`;
}

// ─── System-principal fetch scheme allow-list ────────────────────────────────

const SYSTEM_PRINCIPAL_FETCH_SCHEMES = new Set([
  'mailbox',
  'mailbox-message',
  'imap',
  'imap-message',
  'news',
  'news-message',
]);

function isSystemPrincipalFetchAllowed(url) {
  if (typeof url !== 'string' || !url) return false;
  const colon = url.indexOf(':');
  if (colon <= 0) return false;
  const scheme = url.slice(0, colon).toLowerCase();
  return SYSTEM_PRINCIPAL_FETCH_SCHEMES.has(scheme);
}

// ─── Recursive schema walker (S4) ────────────────────────────────────────────

function validateAgainstSchema(value, schema, path, errors) {
  if (!schema || value === undefined || value === null) return;

  const expectedType = schema.type;
  if (expectedType === 'array') {
    if (!Array.isArray(value)) {
      errors.push(`Parameter '${path}' must be an array, got ${typeof value}`);
      return;
    }
    if (schema.items) {
      for (let i = 0; i < value.length; i++) {
        validateAgainstSchema(value[i], schema.items, `${path}[${i}]`, errors);
      }
    }
  } else if (expectedType === 'object') {
    if (typeof value !== 'object' || Array.isArray(value)) {
      errors.push(`Parameter '${path}' must be an object, got ${Array.isArray(value) ? 'array' : typeof value}`);
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
  } else if (expectedType === 'integer') {
    if (typeof value !== 'number' || !Number.isInteger(value)) {
      errors.push(`Parameter '${path}' must be an integer, got ${typeof value === 'number' ? 'non-integer number' : typeof value}`);
      return;
    }
  } else if (expectedType && typeof value !== expectedType) {
    errors.push(`Parameter '${path}' must be ${expectedType}, got ${typeof value}`);
    return;
  }

  if (Array.isArray(schema.oneOf) && schema.oneOf.length > 0) {
    let matched = 0;
    for (const branch of schema.oneOf) {
      const branchErrors = [];
      validateAgainstSchema(value, branch, path, branchErrors);
      if (branchErrors.length === 0) matched++;
    }
    if (matched === 0) {
      errors.push(`Parameter '${path}' did not match any allowed schema variant`);
    } else if (matched > 1) {
      errors.push(`Parameter '${path}' matched more than one schema variant`);
    }
  }

  if (Array.isArray(schema.enum) && !schema.enum.includes(value)) {
    errors.push(`Parameter '${path}' must be one of ${JSON.stringify(schema.enum)}, got ${JSON.stringify(value)}`);
  }
}

// ─── Rate limiter (sliding window per tool name) ─────────────────────────────

/**
 * Defaults are conservative: outbound mail / contact writes / filter create
 * are the operations a confused LLM agent could do real harm with. Read tools
 * are not rate-limited here -- the HTTP body cap + per-request validation
 * already bound them, and the agent legitimately fans out many reads.
 *
 * Format: tool name -> { limit, windowMs }
 */
const RATE_LIMIT_DEFAULTS = {
  sendMail:         { limit: 10, windowMs: 5 * 60 * 1000 },
  replyToMessage:   { limit: 10, windowMs: 5 * 60 * 1000 },
  forwardMessage:   { limit: 10, windowMs: 5 * 60 * 1000 },
  saveDraft:        { limit: 30, windowMs: 5 * 60 * 1000 },
  createContact:    { limit:  5, windowMs: 60 * 1000 },
  updateContact:    { limit:  5, windowMs: 60 * 1000 },
  deleteContact:    { limit:  5, windowMs: 60 * 1000 },
  createFilter:     { limit:  5, windowMs: 5 * 60 * 1000 },
  updateFilter:     { limit:  5, windowMs: 5 * 60 * 1000 },
  deleteFilter:     { limit: 10, windowMs: 5 * 60 * 1000 },
};

/**
 * Sliding-window rate limiter. Each tool keeps a ring of recent call
 * timestamps; on consume(), we drop timestamps outside the window and check
 * the remaining count against `limit`.
 *
 * Stateless from the caller's view -- pass `now()` so tests can advance time.
 * Lives in security_helpers so tests exercise the same code; production
 * passes Date.now and a singleton state object created at extension startup.
 */
function createRateLimiterState(limits) {
  const config = { ...RATE_LIMIT_DEFAULTS, ...(limits || {}) };
  return { config, hits: Object.create(null) };
}

/**
 * Try to consume one rate-limit slot for `tool`. Returns:
 *   { allowed: true,  remaining, resetAfterMs }  - call is permitted
 *   { allowed: false, remaining: 0, resetAfterMs, limit, windowMs } - blocked
 *
 * If `tool` is not in the config map the call is always allowed (no limit
 * defined means no limit enforced).
 */
function consumeRateLimit(state, tool, nowMs) {
  if (!state || !state.config) return { allowed: true, remaining: Infinity, resetAfterMs: 0 };
  const cfg = state.config[tool];
  if (!cfg) return { allowed: true, remaining: Infinity, resetAfterMs: 0 };
  const now = typeof nowMs === "number" ? nowMs : Date.now();
  const cutoff = now - cfg.windowMs;
  const ring = state.hits[tool] || [];
  // Drop expired timestamps from the front.
  let firstAlive = 0;
  while (firstAlive < ring.length && ring[firstAlive] <= cutoff) firstAlive++;
  const fresh = firstAlive === 0 ? ring : ring.slice(firstAlive);
  if (fresh.length >= cfg.limit) {
    // Earliest fresh hit is the one whose expiry resets a slot.
    const resetAfterMs = (fresh[0] + cfg.windowMs) - now;
    state.hits[tool] = fresh;
    return {
      allowed: false,
      remaining: 0,
      resetAfterMs: Math.max(0, resetAfterMs),
      limit: cfg.limit,
      windowMs: cfg.windowMs,
    };
  }
  fresh.push(now);
  state.hits[tool] = fresh;
  return {
    allowed: true,
    remaining: cfg.limit - fresh.length,
    resetAfterMs: cfg.windowMs,
  };
}

/**
 * Snapshot the current bucket state without mutating it. Used by
 * getServerCapabilities so the LLM can plan around how many slots it has
 * left before the next call.
 */
function inspectRateLimits(state, nowMs) {
  const now = typeof nowMs === "number" ? nowMs : Date.now();
  const result = {};
  if (!state || !state.config) return result;
  for (const tool of Object.keys(state.config)) {
    const cfg = state.config[tool];
    const ring = state.hits[tool] || [];
    const cutoff = now - cfg.windowMs;
    const fresh = ring.filter(t => t > cutoff);
    result[tool] = {
      limit: cfg.limit,
      windowMs: cfg.windowMs,
      used: fresh.length,
      remaining: Math.max(0, cfg.limit - fresh.length),
    };
  }
  return result;
}

module.exports = {
  SENSITIVE_ATTACHMENT_PATTERNS,
  isSensitiveFilePath,
  sanitizeHeaderLine,
  UNTRUSTED_OPEN,
  UNTRUSTED_CLOSE,
  UNTRUSTED_CLOSE_DEFANGED,
  wrapUntrustedBody,
  wrapUntrustedPreview,
  countRecipients,
  summarizeAttachmentsForAudit,
  SAFE_HREF_SCHEMES,
  isSafeMarkdownHref,
  isSafeImageSrc,
  escapeMarkdownLinkText,
  renderMarkdownLink,
  SYSTEM_PRINCIPAL_FETCH_SCHEMES,
  isSystemPrincipalFetchAllowed,
  validateAgainstSchema,
  RATE_LIMIT_DEFAULTS,
  createRateLimiterState,
  consumeRateLimit,
  inspectRateLimits,
};
