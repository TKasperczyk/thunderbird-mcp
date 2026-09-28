#!/usr/bin/env node
/**
 * MCP Bridge for Thunderbird
 *
 * Converts stdio MCP protocol to HTTP requests for the Thunderbird MCP extension.
 * The extension exposes an HTTP endpoint on localhost:8765.
 */

const http = require('http');
const fs = require('fs');
const path = require('path');
const os = require('os');

const THUNDERBIRD_HOSTS = ['127.0.0.1'];
const REQUEST_TIMEOUT = 30000;
const CONNECTION_RETRY_DELAY_MS = 1000;
const CONNECTION_MAX_RETRIES = 5;
const CONNECTION_CACHE_TTL_MS = 5000; // 5 seconds

const DEFAULT_PROC_ROOT = '/proc';
const DEFAULT_DARWIN_FOLDERS_ROOT = '/var/folders';
const THUNDERBIRD_MCP_SUBDIR = 'thunderbird-mcp';
const CONNECTION_FILE_BASENAME = 'connection.json';
const AUTH_TOKEN_PATTERN = /^[0-9a-f]{64}$/;

// MCP protocol versions the bridge knows how to speak. Per lifecycle spec the
// server MUST respond with the requested version if it supports it, otherwise
// with the latest version it supports. The bridge is a transparent JSON-RPC
// relay -- behavior never changes by version -- so it accepts every published
// version, but it does NOT echo unknown future versions back as if it knew them.
const SUPPORTED_PROTOCOL_VERSIONS = new Set([
  '2024-10-07',
  '2024-11-05',
  '2025-03-26',
  '2025-06-18',
  '2025-11-25',
]);
const LATEST_PROTOCOL_VERSION = '2025-11-25';
const BRIDGE_VERSION = (() => {
  try {
    return require('./package.json').version || '0.0.0';
  } catch {
    return '0.0.0';
  }
})();
const SERVER_INFO = Object.freeze({
  name: 'thunderbird-mcp',
  version: BRIDGE_VERSION,
});

const DEBUG = !!process.env.THUNDERBIRD_MCP_DEBUG;

function debugLog(message) {
  if (DEBUG) {
    process.stderr.write('[thunderbird-mcp] ' + message + '\n');
  }
}

function isValidAuthToken(token) {
  return typeof token === 'string' && AUTH_TOKEN_PATTERN.test(token);
}

let cachedConnectionInfo = null;
let connectionCacheExpiry = 0;
let lastDiscoveryAttempts = [];
// Full set of valid connection candidates from the last discovery, in priority
// order. forwardToThunderbird advances through this list when a candidate's
// HTTP endpoint refuses or returns 403, so a stale connection file can't
// permanently mask a live one further down the list.
let cachedCandidateList = [];
let cachedCandidateIndex = 0;

function normalizeFsError(err) {
  if (!err) {
    return 'unknown error';
  }
  if (err.code === 'ENOENT') {
    return 'file not found';
  }
  if (err.code === 'EACCES' || err.code === 'EPERM') {
    return 'permission denied';
  }
  return err.message || String(err);
}

function getCurrentUid(processImpl = process) {
  return typeof processImpl.getuid === 'function' ? processImpl.getuid() : null;
}

function createDiscoveryContext(options = {}) {
  const fsImpl = options.fsImpl || fs;
  const pathImpl = options.pathImpl || path;
  const osImpl = options.osImpl || os;
  const processImpl = options.processImpl || process;
  const env = options.env || processImpl.env || {};
  const uid = Object.prototype.hasOwnProperty.call(options, 'uid')
    ? options.uid
    : getCurrentUid(processImpl);

  return {
    fsImpl,
    pathImpl,
    osImpl,
    processImpl,
    env,
    uid,
    platform: options.platform || processImpl.platform,
    homeDir: Object.prototype.hasOwnProperty.call(options, 'homeDir')
      ? options.homeDir
      : osImpl.homedir(),
    procRoot: options.procRoot || DEFAULT_PROC_ROOT,
    darwinFoldersRoot: options.darwinFoldersRoot || DEFAULT_DARWIN_FOLDERS_ROOT,
    runtimeDir: Object.prototype.hasOwnProperty.call(options, 'runtimeDir')
      ? options.runtimeDir
      : getRuntimeDir({ env, pathImpl, uid }),
  };
}

function getRuntimeDir({ env, pathImpl, uid }) {
  if (env.XDG_RUNTIME_DIR) {
    return env.XDG_RUNTIME_DIR;
  }
  if (uid !== null && uid !== undefined) {
    return pathImpl.join('/run/user', String(uid));
  }
  return null;
}

function getDefaultConnectionFile(context) {
  return context.pathImpl.join(
    context.osImpl.tmpdir(),
    THUNDERBIRD_MCP_SUBDIR,
    CONNECTION_FILE_BASENAME
  );
}

function makeAttempt(label, filePath, reason) {
  return { label, path: filePath, reason };
}

function makeCandidate(label, filePath, mtimeMs = Number.NEGATIVE_INFINITY) {
  return { label, path: filePath, mtimeMs };
}

function addUniqueCandidate(candidates, seenPaths, candidate) {
  if (!candidate.path || seenPaths.has(candidate.path)) {
    return;
  }
  seenPaths.add(candidate.path);
  candidates.push(candidate);
}

function sortCandidatesByMtime(candidates) {
  // When a sandbox scan yields multiple connection files, try the newest file
  // first so selection is deterministic without silently ignoring other paths.
  return candidates.sort((a, b) => {
    if (a.mtimeMs !== b.mtimeMs) {
      return b.mtimeMs - a.mtimeMs;
    }
    return a.path.localeCompare(b.path);
  });
}

function buildScanGroup(label, pattern, candidates, noMatchReason) {
  const notes = [];
  if (candidates.length === 0) {
    notes.push(makeAttempt(label, pattern, noMatchReason));
    return { notes, candidates };
  }
  if (candidates.length > 1) {
    notes.push(makeAttempt(label, pattern, `multiple matches found, trying newest first (${candidates.length} files)`));
  }
  return { notes, candidates: sortCandidatesByMtime(candidates) };
}

function findMacOsConnectionCandidates(context) {
  const { fsImpl, pathImpl, darwinFoldersRoot, uid } = context;
  const pattern = pathImpl.join(
    darwinFoldersRoot,
    '*',
    '*',
    'T',
    THUNDERBIRD_MCP_SUBDIR,
    CONNECTION_FILE_BASENAME
  );

  let firstLevel;
  try {
    firstLevel = fsImpl.readdirSync(darwinFoldersRoot, { withFileTypes: true });
  } catch (err) {
    return {
      notes: [makeAttempt('macOS temp scan', pattern, normalizeFsError(err))],
      candidates: [],
    };
  }

  const candidates = [];
  const seenPaths = new Set();

  for (const firstDir of firstLevel) {
    if (!firstDir.isDirectory()) {
      continue;
    }

    let secondLevel;
    const firstPath = pathImpl.join(darwinFoldersRoot, firstDir.name);
    try {
      secondLevel = fsImpl.readdirSync(firstPath, { withFileTypes: true });
    } catch {
      continue;
    }

    for (const secondDir of secondLevel) {
      if (!secondDir.isDirectory()) {
        continue;
      }

      const candidatePath = pathImpl.join(
        firstPath,
        secondDir.name,
        'T',
        THUNDERBIRD_MCP_SUBDIR,
        CONNECTION_FILE_BASENAME
      );

      try {
        const stat = fsImpl.statSync(candidatePath);
        if (!stat.isFile()) {
          continue;
        }
        if (uid !== null && uid !== undefined && stat.uid !== uid) {
          continue;
        }
        addUniqueCandidate(candidates, seenPaths, makeCandidate('macOS temp scan', candidatePath, stat.mtimeMs));
      } catch (err) {
        if (err.code !== 'ENOENT' && err.code !== 'ENOTDIR') {
          continue;
        }
      }
    }
  }

  const ownerText = uid !== null && uid !== undefined
    ? `no matching files owned by uid ${uid}`
    : 'no matching files';

  return buildScanGroup('macOS temp scan', pattern, candidates, ownerText);
}

function findSnapConnectionCandidates(context) {
  const { fsImpl, pathImpl, homeDir, procRoot } = context;
  const snapDir = homeDir ? pathImpl.join(homeDir, 'snap', 'thunderbird') : null;
  const pattern = pathImpl.join(procRoot, '<pid>', 'environ');

  if (!snapDir) {
    return {
      notes: [makeAttempt('Snap detection', pattern, 'home directory unavailable')],
      candidates: [],
    };
  }

  try {
    fsImpl.accessSync(snapDir, fs.constants.F_OK);
  } catch {
    return {
      notes: [makeAttempt('Snap detection', pattern, 'snap install not detected')],
      candidates: [],
    };
  }

  const candidates = [];
  const seenPaths = new Set();

  try {
    const procDirs = fsImpl.readdirSync(procRoot).filter((entry) => /^\d+$/.test(entry));
    for (const pid of procDirs) {
      try {
        const cmdline = fsImpl.readFileSync(pathImpl.join(procRoot, pid, 'cmdline'), 'utf8');
        // Match argv[0] basename precisely -- not any occurrence of 'thunderbird'
        // in argv. A text editor opened on 'thunderbird.txt' would have the
        // substring in argv[1], and we do NOT want to read its TMPDIR.
        const argv0 = cmdline.split('\0')[0] || '';
        const argv0Basename = pathImpl.basename(argv0);
        if (!/^(thunderbird|betterbird)(-.+)?$/.test(argv0Basename)) {
          continue;
        }

        const environ = fsImpl.readFileSync(pathImpl.join(procRoot, pid, 'environ'), 'utf8');
        const tmpEntry = environ.split('\0').find((entry) => entry.startsWith('TMPDIR='));
        if (!tmpEntry) {
          continue;
        }

        const tmpDir = tmpEntry.slice('TMPDIR='.length);
        const candidatePath = pathImpl.join(tmpDir, THUNDERBIRD_MCP_SUBDIR, CONNECTION_FILE_BASENAME);
        let mtimeMs = Number.NEGATIVE_INFINITY;
        try {
          mtimeMs = fsImpl.statSync(candidatePath).mtimeMs;
        } catch {
          // Missing file is handled later when the candidate is read.
        }
        addUniqueCandidate(
          candidates,
          seenPaths,
          makeCandidate(`Snap TMPDIR from /proc/${pid}/environ`, candidatePath, mtimeMs)
        );
      } catch {
        // Processes can disappear or deny access while we scan /proc.
      }
    }
  } catch (err) {
    return {
      notes: [makeAttempt('Snap detection', pattern, normalizeFsError(err))],
      candidates: [],
    };
  }

  // Match the official snap tmpdir helper as a best-effort fallback when /proc
  // cannot tell us the runtime TMPDIR.
  const fallbackPath = pathImpl.join(
    homeDir,
    'Downloads',
    'thunderbird.tmp',
    THUNDERBIRD_MCP_SUBDIR,
    CONNECTION_FILE_BASENAME
  );
  let fallbackMtime = Number.NEGATIVE_INFINITY;
  try {
    fallbackMtime = fsImpl.statSync(fallbackPath).mtimeMs;
  } catch {
    // Missing file is handled later when the candidate is read.
  }
  addUniqueCandidate(
    candidates,
    seenPaths,
    makeCandidate('Snap Downloads fallback', fallbackPath, fallbackMtime)
  );

  return buildScanGroup('Snap detection', pattern, candidates, 'no thunderbird TMPDIR candidates found');
}

function findFlatpakConnectionCandidates(context) {
  const { fsImpl, pathImpl, runtimeDir } = context;
  const patternBase = runtimeDir || '$XDG_RUNTIME_DIR';
  const pattern = pathImpl.join(
    patternBase,
    'app',
    '*',
    THUNDERBIRD_MCP_SUBDIR,
    CONNECTION_FILE_BASENAME
  );

  if (!runtimeDir) {
    return {
      notes: [makeAttempt('Flatpak scan', pattern, 'runtime dir unavailable')],
      candidates: [],
    };
  }

  const appRoot = pathImpl.join(runtimeDir, 'app');
  let appEntries;
  try {
    appEntries = fsImpl.readdirSync(appRoot, { withFileTypes: true });
  } catch (err) {
    return {
      notes: [makeAttempt('Flatpak scan', pattern, normalizeFsError(err))],
      candidates: [],
    };
  }

  const candidates = [];
  const seenPaths = new Set();

  for (const appEntry of appEntries) {
    if (!appEntry.isDirectory()) {
      continue;
    }

    const candidatePath = pathImpl.join(
      appRoot,
      appEntry.name,
      THUNDERBIRD_MCP_SUBDIR,
      CONNECTION_FILE_BASENAME
    );

    try {
      const stat = fsImpl.statSync(candidatePath);
      if (!stat.isFile()) {
        continue;
      }
      addUniqueCandidate(candidates, seenPaths, makeCandidate('Flatpak runtime scan', candidatePath, stat.mtimeMs));
    } catch (err) {
      if (err.code !== 'ENOENT' && err.code !== 'ENOTDIR') {
        continue;
      }
    }
  }

  return buildScanGroup('Flatpak scan', pattern, candidates, 'no matching files');
}

function buildCandidateGroups(options = {}) {
  const context = createDiscoveryContext(options);
  const groups = [];

  if (context.env.THUNDERBIRD_MCP_CONNECTION_FILE) {
    groups.push({
      notes: [],
      candidates: [
        makeCandidate(
          'THUNDERBIRD_MCP_CONNECTION_FILE',
          context.env.THUNDERBIRD_MCP_CONNECTION_FILE
        )
      ],
      stopOnFailure: true,
      context,
    });
    return groups;
  }

  groups.push({
    notes: [],
    candidates: [makeCandidate('native tmp', getDefaultConnectionFile(context))],
    stopOnFailure: false,
    context,
  });

  if (context.platform === 'darwin') {
    groups.push({ ...findMacOsConnectionCandidates(context), stopOnFailure: false, context });
  }

  if (context.platform === 'linux') {
    groups.push({ ...findSnapConnectionCandidates(context), stopOnFailure: false, context });
    groups.push({ ...findFlatpakConnectionCandidates(context), stopOnFailure: false, context });
  }

  return groups;
}

function tryReadConnectionCandidate(candidate, context) {
  try {
    const raw = context.fsImpl.readFileSync(candidate.path, 'utf8');
    let data;
    try {
      data = JSON.parse(raw);
    } catch (err) {
      return {
        ok: false,
        attempt: makeAttempt(candidate.label, candidate.path, `malformed JSON (${err.message})`)
      };
    }

    if (!data.port || !data.token) {
      return {
        ok: false,
        attempt: makeAttempt(candidate.label, candidate.path, 'missing port or token')
      };
    }

    return {
      ok: true,
      data,
      attempt: makeAttempt(candidate.label, candidate.path, 'ok')
    };
  } catch (err) {
    return {
      ok: false,
      attempt: makeAttempt(candidate.label, candidate.path, normalizeFsError(err))
    };
  }
}

function discoverConnectionInfo(options = {}) {
  const groups = buildCandidateGroups(options);
  const attempts = [];
  const candidates = [];

  for (const group of groups) {
    attempts.push(...group.notes);

    for (const candidate of group.candidates) {
      const result = tryReadConnectionCandidate(candidate, group.context);
      attempts.push(result.attempt);
      if (result.ok) {
        candidates.push({ data: result.data, path: candidate.path });
        if (group.stopOnFailure) {
          // Hard pin (e.g. THUNDERBIRD_MCP_CONNECTION_FILE): user explicitly named
          // this candidate; honor it and don't fall through to autodiscovery.
          return { candidates, attempts };
        }
      } else if (group.stopOnFailure) {
        // Pinned path failed; do not fall through to autodiscovery candidates.
        return { candidates, attempts };
      }
    }
  }

  return { candidates, attempts };
}

// Max raw bytes for an attachment read from a path before base64 encoding.
// Encoded size grows ~33%, so 18 MB raw → ~24 MB base64, staying under the
// extension's 25 MB MAX_BASE64_SIZE limit.
const MAX_ATTACHMENT_BYTES = 18 * 1024 * 1024;
// Keep these message-wide limits in sync with extension/mcp_server/api.js.
const MAX_TOTAL_ATTACHMENT_BYTES = 50 * 1024 * 1024;
const MAX_ATTACHMENTS_PER_MESSAGE = 20;

// File paths that an MCP caller must never be allowed to attach to outbound
// mail. Keep the pattern list and helper behavior identical to the extension so
// neither transport can bypass the LLM-confused-deputy defense.
// Keep in sync with extension/mcp_server/api.js isSensitiveFilePath.
const SENSITIVE_ATTACHMENT_PATTERNS = [
  // SSH / PGP / cloud / kube / docker credentials
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
  // Common key / secret file extensions anywhere on disk
  /\/id_(rsa|dsa|ecdsa|ed25519)(\.pub)?$/,
  /\.pem$/,
  /\.pfx$/,
  /\.p12$/,
  /\.kdbx$/,
  /\.key$/,
  /\.asc$/,
  /\.gpg$/,
  // Linux / macOS system directories
  /^\/etc\//,
  /^\/proc\//,
  /^\/sys\//,
  /^\/root\//,
  /^\/var\/log\//,
  /^\/var\/lib\/sudo\//,
  // macOS keychain locations
  /\/library\/keychains\//,
  // Windows system directories
  /^[a-z]:\/windows\//,
  /^[a-z]:\/programdata\/microsoft\/(crypto|protect)\//,
  /\/appdata\/(local|roaming)\/microsoft\/(credentials|crypto|protect|vault)(\/|$)/,
  // Browser credential stores (Firefox / Chrome / Edge)
  /\/(logins\.json|key3\.db|key4\.db|cookies(\.sqlite)?|login data)$/,
  // Thunderbird's own profile (contains the user's entire mail store + prefs).
  // Linux profile directories and profiles.ini live directly under
  // ~/.thunderbird (or ~/.icedove), while macOS and Windows use the platform
  // application-data directories below. Block each profile root in full.
  /\/\.(?:thunderbird|icedove)(\/|$)/,
  /\/library\/thunderbird(\/|$)/,
  /\/appdata\/roaming\/thunderbird(\/|$)/,
];

function isSensitiveFilePath(attachmentPath) {
  if (typeof attachmentPath !== 'string' || !attachmentPath) return false;
  const normalized = attachmentPath.replace(/\\/g, '/').toLowerCase();
  return SENSITIVE_ATTACHMENT_PATTERNS.some(re => re.test(normalized));
}

// Tools whose `attachments` array may contain string file paths that this
// bridge resolves on the host filesystem before forwarding. Needed because the
// Thunderbird snap (and other sandboxed installs) cannot see arbitrary host
// paths like /data/... or the host's /tmp; passing those paths through to the
// extension results in silent "failed to attach" warnings since file.exists()
// returns false inside the sandbox. Reading on the bridge side and shipping
// inline base64 sidesteps the sandbox entirely.
const ATTACHMENT_TOOLS = new Set(['sendMail', 'replyToMessage', 'forwardMessage']);

// Minimal MIME map covering common attachment types (documents, images,
// archives, A/V). Falls back to application/octet-stream which Thunderbird
// handles fine.
const MIME_BY_EXT = {
  pdf: 'application/pdf',
  png: 'image/png',
  jpg: 'image/jpeg',
  jpeg: 'image/jpeg',
  gif: 'image/gif',
  webp: 'image/webp',
  svg: 'image/svg+xml',
  bmp: 'image/bmp',
  tif: 'image/tiff',
  tiff: 'image/tiff',
  heic: 'image/heic',
  txt: 'text/plain',
  csv: 'text/csv',
  html: 'text/html',
  htm: 'text/html',
  md: 'text/markdown',
  json: 'application/json',
  xml: 'application/xml',
  yml: 'application/yaml',
  yaml: 'application/yaml',
  zip: 'application/zip',
  tar: 'application/x-tar',
  gz: 'application/gzip',
  '7z': 'application/x-7z-compressed',
  doc: 'application/msword',
  docx: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
  xls: 'application/vnd.ms-excel',
  xlsx: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  ppt: 'application/vnd.ms-powerpoint',
  pptx: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
  odt: 'application/vnd.oasis.opendocument.text',
  ods: 'application/vnd.oasis.opendocument.spreadsheet',
  odp: 'application/vnd.oasis.opendocument.presentation',
  ics: 'text/calendar',
  eml: 'message/rfc822',
  mp3: 'audio/mpeg',
  wav: 'audio/wav',
  ogg: 'audio/ogg',
  m4a: 'audio/mp4',
  mp4: 'video/mp4',
  webm: 'video/webm',
  mov: 'video/quicktime'
};

function guessContentType(filePath) {
  const ext = path.extname(filePath).toLowerCase().replace(/^\./, '');
  return MIME_BY_EXT[ext] || 'application/octet-stream';
}

function attachmentError(action, filePath, error) {
  if (error?.code === 'ENOENT') {
    return new Error(`Attachment not found: ${filePath}`, { cause: error });
  }
  if (error?.code === 'EACCES' || error?.code === 'EPERM') {
    return new Error(`Attachment unreadable (permission denied): ${filePath}`, { cause: error });
  }
  return new Error(`Attachment ${action} failed (${error?.code || 'unknown'}): ${filePath}`, { cause: error });
}

function validateAttachmentStat(filePath, stat) {
  if (stat.isSymbolicLink()) {
    throw new Error(`Attachment path is a symlink and is not allowed: ${filePath}`);
  }
  if (!stat.isFile()) {
    throw new Error(`Attachment is not a regular file: ${filePath}`);
  }
  if (!Number.isSafeInteger(stat.size) || stat.size < 0) {
    throw new Error(`Attachment has an invalid file size: ${filePath}`);
  }
  if (stat.size > MAX_ATTACHMENT_BYTES) {
    throw new Error(
      `Attachment too large: ${filePath} is ${stat.size} bytes ` +
      `(limit ${MAX_ATTACHMENT_BYTES} bytes / ${MAX_ATTACHMENT_BYTES / 1024 / 1024} MB raw before base64)`
    );
  }
}

async function inspectAttachmentPath(filePath) {
  // Check both the supplied path and its lexical normalization before any
  // filesystem access. The latter catches paths such as /tmp/../etc/passwd.
  if (isSensitiveFilePath(filePath) || isSensitiveFilePath(path.resolve(filePath))) {
    throw new Error(`Sensitive attachment path blocked: ${filePath}`);
  }

  let stat;
  try {
    // lstat is deliberate: stat would follow the final symlink before policy
    // could reject it.
    stat = await fs.promises.lstat(filePath);
  } catch (e) {
    throw attachmentError('lstat', filePath, e);
  }
  validateAttachmentStat(filePath, stat);
  return { filePath, stat };
}

function sameFile(left, right) {
  return left.dev === right.dev && left.ino === right.ino;
}

async function readFileHandleExactly(handle, filePath, size) {
  const buffer = Buffer.allocUnsafe(size);
  let offset = 0;
  while (offset < size) {
    const { bytesRead } = await handle.read(buffer, offset, size - offset, offset);
    if (bytesRead === 0) {
      throw new Error(`Attachment changed while being read: ${filePath}`);
    }
    offset += bytesRead;
  }

  // Do not let a file that grew after fstat trigger an unbounded read.
  const extra = Buffer.allocUnsafe(1);
  const { bytesRead } = await handle.read(extra, 0, 1, size);
  if (bytesRead !== 0) {
    throw new Error(`Attachment changed while being read: ${filePath}`);
  }

  return buffer;
}

// Read a preflighted file path off the host filesystem and convert it to the
// inline { name, contentType, base64 } shape the extension supports. Opening
// with O_NOFOLLOW where available and comparing the opened file to the lstat
// snapshot prevents a path swap from redirecting the read to a symlink/other
// inode between policy validation and I/O.
async function readAttachmentFromPath(fileInfo) {
  const { filePath, stat: preflightStat } = fileInfo;
  const freshInfo = await inspectAttachmentPath(filePath);
  if (!sameFile(preflightStat, freshInfo.stat) || preflightStat.size !== freshInfo.stat.size) {
    throw new Error(`Attachment changed after validation: ${filePath}`);
  }

  const noFollow = fs.constants.O_NOFOLLOW || 0;
  let handle;
  try {
    handle = await fs.promises.open(filePath, fs.constants.O_RDONLY | noFollow);
  } catch (e) {
    if (e?.code === 'ELOOP') {
      throw new Error(`Attachment path is a symlink and is not allowed: ${filePath}`, { cause: e });
    }
    throw attachmentError('open', filePath, e);
  }

  try {
    let openedStat;
    try {
      openedStat = await handle.stat();
    } catch (e) {
      throw attachmentError('fstat', filePath, e);
    }
    validateAttachmentStat(filePath, openedStat);
    if (!sameFile(freshInfo.stat, openedStat) || freshInfo.stat.size !== openedStat.size) {
      throw new Error(`Attachment changed after validation: ${filePath}`);
    }
    const buffer = await readFileHandleExactly(handle, filePath, openedStat.size);
    return {
      name: path.basename(filePath),
      contentType: guessContentType(filePath),
      base64: buffer.toString('base64')
    };
  } finally {
    await handle.close();
  }
}

// Replace every string entry in `args.attachments` (= file path) with an
// inline { name, contentType, base64 } object read off the host filesystem.
// Inline objects pass through unchanged. All paths and message-wide limits are
// preflighted before the first read, then files are read sequentially so a
// caller cannot force many large buffers to be resident at once.
async function inlineAttachmentPaths(args) {
  if (!args || !Array.isArray(args.attachments)) return;

  if (args.attachments.length > MAX_ATTACHMENTS_PER_MESSAGE) {
    throw new Error(
      `Attachment count ${args.attachments.length} exceeds the ` +
      `${MAX_ATTACHMENTS_PER_MESSAGE} attachment limit`
    );
  }

  const fileInfoByIndex = new Map();
  let totalAttachmentBytes = 0;
  for (let index = 0; index < args.attachments.length; index++) {
    const entry = args.attachments[index];
    if (typeof entry !== 'string') continue;

    const fileInfo = await inspectAttachmentPath(entry);
    if (fileInfo.stat.size > MAX_TOTAL_ATTACHMENT_BYTES - totalAttachmentBytes) {
      throw new Error(
        `Attachment aggregate too large at ${entry}: exceeds the ` +
        `${MAX_TOTAL_ATTACHMENT_BYTES / 1024 / 1024} MB aggregate attachment limit`
      );
    }
    totalAttachmentBytes += fileInfo.stat.size;
    fileInfoByIndex.set(index, fileInfo);
  }

  const resolved = [];
  for (let index = 0; index < args.attachments.length; index++) {
    const entry = args.attachments[index];
    resolved.push(
      typeof entry === 'string'
        ? await readAttachmentFromPath(fileInfoByIndex.get(index))
        : entry
    );
  }
  args.attachments = resolved;
}

// Bridge-local tool definition for saveMessage. The extension can only write
// attachments into the OS temp dir and returns rawSource as text -- neither can
// save to a caller-chosen path. The bridge already reads/writes host files, so
// it owns this tool and injects it into tools/list; it is never forwarded to
// the extension.
const SAVE_MESSAGE_TOOL = {
  name: 'saveMessage',
  description: "Save a message to disk on the local machine: writes the full .eml and/or its attachments into a destination directory. Unlike saveAttachments (temp dir only), this writes to any path you choose. Runs on the bridge, so destDir is a path on the machine running the MCP.",
  inputSchema: {
    type: 'object',
    properties: {
      messageId: { type: 'string', description: 'Message ID (from searchMessages/getRecentMessages).' },
      folderPath: { type: 'string', description: 'Folder URI path containing the message.' },
      destDir: { type: 'string', description: 'Directory to write into (created recursively if missing).' },
      saveEml: { type: 'boolean', description: 'Write the full raw .eml (default true).' },
      saveAttachments: { type: 'boolean', description: 'Write attachments as loose files (default false).' },
      emlFilename: { type: 'string', description: 'Filename for the .eml (default: the sanitized message subject, else messageId). ".eml" is appended if absent.' },
      overwrite: { type: 'boolean', description: 'Overwrite existing files instead of erroring (default false).' }
    },
    required: ['messageId', 'folderPath', 'destDir']
  }
};

// saveMessage hands whatever drives this MCP an arbitrary filesystem-write
// primitive: attacker-supplied attachment bytes at an attacker-supplied path.
// A prompt-injected tool call could target ~/.bashrc, ~/.ssh/, or an autostart
// entry. So the tool is opt-in -- absent from tools/list and refused on
// tools/call unless THUNDERBIRD_MCP_SAVE_MESSAGE=1 -- and can be narrowed
// further by THUNDERBIRD_MCP_SAVE_ROOT, which destDir must resolve inside.
const SAVE_MESSAGE_ENV = 'THUNDERBIRD_MCP_SAVE_MESSAGE';
const SAVE_ROOT_ENV = 'THUNDERBIRD_MCP_SAVE_ROOT';

// Directories get 0700 and files 0600: saved mail is private, and under a
// normal umask node would otherwise create them world-readable.
const SAVE_DIR_MODE = 0o700;
const SAVE_FILE_MODE = 0o600;

function saveMessageEnabled(env = process.env) {
  return env[SAVE_MESSAGE_ENV] === '1';
}

function saveMessageRoot(env = process.env) {
  const raw = env[SAVE_ROOT_ENV];
  return typeof raw === 'string' && raw.trim() ? raw.trim() : null;
}

// Resolve a path to its real location even when it does not exist yet: walk up
// to the nearest existing ancestor, realpath *that* (which is what defeats a
// symlinked intermediate directory), then re-append the missing tail. Resolving
// only the existing prefix is the point -- realpathSync on the full path throws
// ENOENT for a directory we are about to create.
function resolveIntendedPath(target, fsImpl, pathImpl) {
  const abs = pathImpl.resolve(target);
  const missing = [];
  let cursor = abs;
  for (;;) {
    try {
      const real = fsImpl.realpathSync(cursor);
      return missing.length ? pathImpl.join(real, ...missing.reverse()) : real;
    } catch (e) {
      if (e.code !== 'ENOENT') {
        throw e;
      }
      const parent = pathImpl.dirname(cursor);
      if (parent === cursor) {
        return abs;
      }
      missing.push(pathImpl.basename(cursor));
      cursor = parent;
    }
  }
}

// Resolve destDir through any symlinks and, when a save root is configured,
// refuse anything outside it. Returns the resolved directory, which is what
// every later path operation uses -- checking the caller's spelling and then
// writing to the unresolved path would leave the symlink hole open.
function resolveDestDir(destDir, { env = process.env, fsImpl = fs, pathImpl = path } = {}) {
  const resolved = resolveIntendedPath(destDir, fsImpl, pathImpl);
  const root = saveMessageRoot(env);
  if (!root) {
    return resolved;
  }
  const realRoot = resolveIntendedPath(root, fsImpl, pathImpl);
  const rel = pathImpl.relative(realRoot, resolved);
  if (rel !== '' && (rel.startsWith('..') || pathImpl.isAbsolute(rel))) {
    throw new Error(
      `destDir resolves to ${resolved}, which is outside ${SAVE_ROOT_ENV} (${realRoot})`
    );
  }
  return resolved;
}

// Unwrap a tools/call response envelope into its inner data object. Tool results
// are shaped { result: { content: [ { type:'text', text:'<JSON>' } ] } }; a
// tool-level failure comes back as { error } inside that JSON, which we surface
// as a thrown Error so the caller reports it instead of silently continuing.
function unwrapToolResponse(resp) {
  const text = resp && resp.result && Array.isArray(resp.result.content)
    ? resp.result.content[0] && resp.result.content[0].text
    : undefined;
  if (typeof text !== 'string') {
    throw new Error('Unexpected tool response shape from getMessage');
  }
  const data = JSON.parse(text);
  if (data && data.error) {
    throw new Error(data.error);
  }
  return data;
}

// Windows device names are reserved at every directory level and with any
// extension: CON, COM3.txt and PRN.eml all resolve to the device, not a file.
const WINDOWS_RESERVED_STEM = /^(CON|PRN|AUX|NUL|COM[1-9]|LPT[1-9])$/i;

// Turn an arbitrary name into a safe single-path-segment filename: drop path
// separators, control chars and the characters Windows forbids (which also
// closes NTFS alternate-data-stream names, since those need the colon),
// collapse whitespace, strip the trailing dots and spaces Windows silently
// discards, and refuse names that would be empty or traverse ('.'/'..').
// Spaces, unicode, and parentheses are kept -- they appear in real filenames.
function sanitizeSaveFilename(name) {
  let s = String(name == null ? '' : name);
  // eslint-disable-next-line no-control-regex
  s = s.replace(/[/\\\x00-\x1f\x7f<>:"|?*]/g, ' ');
  s = s.replace(/\s+/g, ' ').trim();
  // Do this after trimming: ' .. ' and 'foo...' both need it, and a name of
  // nothing but dots must collapse to empty so the guard below rejects it.
  s = s.replace(/[. ]+$/, '');
  if (!s || s === '.' || s === '..') {
    throw new Error(`Cannot derive a safe filename from ${JSON.stringify(String(name))}`);
  }
  const stem = s.replace(/\..*$/, '');
  if (WINDOWS_RESERVED_STEM.test(stem)) {
    s = `_${s}`;
  }
  return s;
}

// Generate ' (2)', ' (3)', ... variants of a filename, suffixed before the
// extension. Yields the base name first.
function* collisionCandidates(baseName, pathImpl) {
  yield baseName;
  const ext = pathImpl.extname(baseName);
  const stem = baseName.slice(0, baseName.length - ext.length);
  for (let n = 2; ; n++) {
    yield `${stem} (${n})${ext}`;
  }
}

// Remove an existing entry so a subsequent exclusive create lands on a fresh
// file. Used only in overwrite mode, and it is what makes overwrite no-follow:
// unlinking a symlink removes the link, so the create that follows cannot be
// redirected through it.
function unlinkForOverwrite(destPath, fsImpl) {
  fsImpl.rmSync(destPath, { force: true });
}

// Copy an attachment out of the extension's temp dir, never following a
// symlink planted at the destination and never racing an existence check: the
// exclusive create *is* the check. Retries the ' (n)' suffix on collision
// unless overwrite is on, in which case the existing entry is unlinked first.
// COPYFILE_EXCL is the copy-side equivalent of 'wx':
// without it copyFileSync happily follows and overwrites a symlink planted at
// the destination, which is exactly what `overwrite: false` promises not to do.
function copyFileExclusive(srcPath, destDir, baseName, { usedNames, overwrite, fsImpl, pathImpl }) {
  for (const candidate of collisionCandidates(baseName, pathImpl)) {
    if (usedNames.has(candidate)) {
      continue;
    }
    const destPath = pathImpl.join(destDir, candidate);
    try {
      fsImpl.copyFileSync(srcPath, destPath, fs.constants.COPYFILE_EXCL);
    } catch (e) {
      if (e.code !== 'EEXIST') {
        throw e;
      }
      if (!overwrite) {
        continue;
      }
      unlinkForOverwrite(destPath, fsImpl);
      fsImpl.copyFileSync(srcPath, destPath, fs.constants.COPYFILE_EXCL);
    }
    fsImpl.chmodSync(destPath, SAVE_FILE_MODE);
    usedNames.add(candidate);
    return destPath;
  }
}

// Write a message's .eml and/or attachments into a caller-chosen directory on
// the host filesystem. fetchMessage(subMessage) returns the parsed JSON-RPC
// response for a tools/call (forwardToThunderbird in production, a mock in
// tests). Returns the inner data object { emlPath, attachments, skippedAttachments,
// destDir }; the handleMessage caller wraps it into the MCP content envelope.
async function saveMessageToDisk({ args, fetchMessage, fsImpl = fs, pathImpl = path, env = process.env }) {
  args = args || {};
  if (!saveMessageEnabled(env)) {
    throw new Error(
      `saveMessage is disabled. Set ${SAVE_MESSAGE_ENV}=1 in the MCP server environment ` +
      `to enable writing files to disk (optionally with ${SAVE_ROOT_ENV} to confine them).`
    );
  }
  const { messageId, folderPath } = args;
  if (typeof messageId !== 'string' || !messageId.trim()) {
    throw new Error('saveMessage requires a non-empty messageId string');
  }
  if (typeof folderPath !== 'string' || !folderPath.trim()) {
    throw new Error('saveMessage requires a non-empty folderPath string');
  }
  if (typeof args.destDir !== 'string' || !args.destDir.trim()) {
    throw new Error('saveMessage requires a non-empty destDir string');
  }

  const overwrite = args.overwrite === true;
  const writeEml = args.saveEml !== false;
  const writeAttachments = args.saveAttachments === true;

  // Resolve before creating: a symlinked ancestor must be followed to its real
  // location and checked against the save root *first*, or the root check
  // guards a path nothing is ever written to.
  const destDir = resolveDestDir(args.destDir, { env, fsImpl, pathImpl });
  fsImpl.mkdirSync(destDir, { recursive: true, mode: SAVE_DIR_MODE });

  // Shared across the .eml and the attachments so a second file can never take
  // a name already written this run -- including the .eml's own name, which
  // under the previous overwrite behavior an attachment could clobber.
  const usedNames = new Set();
  let emlPath = null;

  if (writeEml) {
    const resp = await fetchMessage({
      jsonrpc: '2.0',
      id: 'saveMessage-eml',
      method: 'tools/call',
      params: {
        name: 'getMessage',
        arguments: { messageId, folderPath, rawSource: true }
      }
    });
    const data = unwrapToolResponse(resp);
    const rawSource = data.rawSource;
    if (typeof rawSource !== 'string' || rawSource.length === 0) {
      throw new Error(
        'getMessage returned no rawSource; saving the .eml requires a local/offline ' +
        'copy of the message (IMAP messages not cached offline cannot be read).'
      );
    }

    let filename;
    if (typeof args.emlFilename === 'string' && args.emlFilename.trim()) {
      filename = args.emlFilename;
    } else {
      // data.subject arrives already RFC 2047-decoded from Thunderbird
      // (mime2DecodedSubject), so there is nothing left for the bridge to
      // decode -- and nothing to parse back out of the raw headers.
      filename = typeof data.subject === 'string' && data.subject.trim()
        ? data.subject
        : messageId;
    }
    filename = sanitizeSaveFilename(filename);
    if (!/\.eml$/i.test(filename)) {
      filename += '.eml';
    }

    // rawSource is a Latin-1 byte string: the extension reads the message as
    // bytes and hands each one over as a code point. Writing it as UTF-8 would
    // re-encode every byte above 0x7f into two, corrupting 8-bit bodies, MIME
    // parts and signatures. Convert back to the original bytes instead.
    const emlBytes = Buffer.from(rawSource, 'latin1');

    // The .eml does not get a ' (2)' suffix the way attachments do -- a caller
    // naming the file expects that name or an error. The exclusive create is
    // the existence check, so there is no window to lose.
    const fullPath = pathImpl.join(destDir, filename);
    try {
      fsImpl.writeFileSync(fullPath, emlBytes, { flag: 'wx', mode: SAVE_FILE_MODE });
    } catch (e) {
      if (e.code !== 'EEXIST') {
        throw e;
      }
      if (!overwrite) {
        throw new Error(
          `File already exists: ${fullPath} (set overwrite:true to replace it)`,
          { cause: e }
        );
      }
      unlinkForOverwrite(fullPath, fsImpl);
      fsImpl.writeFileSync(fullPath, emlBytes, { flag: 'wx', mode: SAVE_FILE_MODE });
    }
    usedNames.add(filename);
    emlPath = fullPath;
  }

  const writtenAttachments = [];
  const skippedAttachments = [];

  if (writeAttachments) {
    const resp = await fetchMessage({
      jsonrpc: '2.0',
      id: 'saveMessage-attachments',
      method: 'tools/call',
      params: {
        name: 'getMessage',
        arguments: { messageId, folderPath, saveAttachments: true }
      }
    });
    const data = unwrapToolResponse(resp);
    const attachments = Array.isArray(data.attachments) ? data.attachments : [];
    for (const att of attachments) {
      const srcPath = att && att.filePath;
      if (typeof srcPath !== 'string' || !srcPath) {
        skippedAttachments.push((att && att.name) || null);
        continue;
      }
      const baseName = sanitizeSaveFilename(att.name || pathImpl.basename(srcPath));
      writtenAttachments.push(
        copyFileExclusive(srcPath, destDir, baseName, { usedNames, overwrite, fsImpl, pathImpl })
      );
    }
  }

  return { emlPath, attachments: writtenAttachments, skippedAttachments, destDir };
}

/**
 * Read connection info (port + auth token) written by the Thunderbird extension.
 * Returns { port, token } or null if no valid candidate exists.
 * Caches the full candidate list for a short TTL so forwardToThunderbird can
 * advance past a stale winner on connection failure without re-running discovery.
 */
function readConnectionInfo(options = {}) {
  if (cachedConnectionInfo && Date.now() < connectionCacheExpiry) {
    return cachedConnectionInfo;
  }

  const result = discoverConnectionInfo(options);
  lastDiscoveryAttempts = result.attempts;
  cachedCandidateList = result.candidates;
  cachedCandidateIndex = 0;

  if (!cachedCandidateList.length) {
    return null;
  }

  cachedConnectionInfo = cachedCandidateList[0].data;
  connectionCacheExpiry = Date.now() + CONNECTION_CACHE_TTL_MS;
  return cachedConnectionInfo;
}

/**
 * Advance to the next cached connection candidate after the current one fails
 * to reach Thunderbird. Returns the new candidate's data, or null when the
 * cached list is exhausted (caller should rediscover from scratch).
 */
function advanceToNextCandidate() {
  if (!cachedCandidateList.length) {
    return null;
  }
  cachedCandidateIndex += 1;
  if (cachedCandidateIndex >= cachedCandidateList.length) {
    return null;
  }
  cachedConnectionInfo = cachedCandidateList[cachedCandidateIndex].data;
  connectionCacheExpiry = Date.now() + CONNECTION_CACHE_TTL_MS;
  return cachedConnectionInfo;
}

function clearConnectionCache() {
  cachedConnectionInfo = null;
  connectionCacheExpiry = 0;
  cachedCandidateList = [];
  cachedCandidateIndex = 0;
}

function formatDiscoveryAttempts(attempts = lastDiscoveryAttempts) {
  if (!attempts.length) {
    return 'no candidates generated';
  }

  return attempts
    .map((attempt) => `${attempt.label} (${attempt.path}): ${attempt.reason}`)
    .join('; ');
}

function buildConnectionDiscoveryErrorMessage() {
  return (
    'Connection discovery failed. ' +
    'Tried: ' + formatDiscoveryAttempts() + '. ' +
    'Is Thunderbird running with the MCP extension? ' +
    'The extension must be started first to create the connection file.'
  );
}

function sanitizeJson(data) {
  // Remove control chars except \n, \r, \t. The character class is
  // intentional -- some clients emit stray control bytes and we
  // sanitize them out before JSON.parse() chokes on them.
  // eslint-disable-next-line no-control-regex
  let sanitized = data.replace(/[\x00-\x08\x0b\x0c\x0e-\x1f\x7f]/g, '');
  // Escape raw newlines/carriage returns/tabs that aren't already escaped.
  // Match an even number of backslashes (including zero) before the control
  // char so we don't double-escape already-escaped sequences like \n, but
  // do escape after literal backslash pairs like \\\n (escaped-backslash + raw newline).
  sanitized = sanitized.replace(/((?:^|[^\\])(?:\\\\)*)\r/gm, '$1\\r');
  sanitized = sanitized.replace(/((?:^|[^\\])(?:\\\\)*)\n/gm, '$1\\n');
  sanitized = sanitized.replace(/((?:^|[^\\])(?:\\\\)*)\t/gm, '$1\\t');
  return sanitized;
}

// deps is a seam for tests: production passes nothing and gets the real
// transport and the process environment.
async function handleMessage(line, { forward = forwardToThunderbird, env = process.env } = {}) {
  const message = JSON.parse(line);
  const hasId = Object.prototype.hasOwnProperty.call(message, 'id');
  const isNotification =
    !hasId ||
    (typeof message.method === 'string' && message.method.startsWith('notifications/'));

  if (isNotification) {
    return null;
  }

  // Handle MCP lifecycle methods locally so the bridge can complete
  // handshake even when Thunderbird isn't running yet.
  switch (message.method) {
    case 'initialize': {
      const requested = message.params?.protocolVersion;
      if (typeof requested !== 'string') {
        return {
          jsonrpc: '2.0',
          id: message.id,
          error: {
            code: -32602,
            message: 'Invalid params: protocolVersion must be a string',
          },
        };
      }
      const negotiated = SUPPORTED_PROTOCOL_VERSIONS.has(requested)
        ? requested
        : LATEST_PROTOCOL_VERSION;
      return {
        jsonrpc: '2.0',
        id: message.id,
        result: {
          protocolVersion: negotiated,
          capabilities: { tools: {} },
          serverInfo: SERVER_INFO,
        },
      };
    }
    case 'ping':
      return { jsonrpc: '2.0', id: message.id, result: {} };
    case 'resources/list':
      return { jsonrpc: '2.0', id: message.id, result: { resources: [] } };
    case 'prompts/list':
      return { jsonrpc: '2.0', id: message.id, result: { prompts: [] } };
  }

  // For mail-sending tools, inline any attachments passed as file paths.
  // The Thunderbird extension may run inside a sandboxed snap that cannot
  // see /data/..., the host /tmp, or any path outside its confined view —
  // letting paths through results in silent "failed to attach" warnings.
  // Reading on the bridge side and shipping base64 sidesteps the sandbox.
  if (message.method === 'tools/call'
      && message.params
      && ATTACHMENT_TOOLS.has(message.params.name)) {
    try {
      await inlineAttachmentPaths(message.params.arguments);
    } catch (e) {
      return {
        jsonrpc: '2.0',
        id: message.id,
        error: { code: -32602, message: e.message }
      };
    }
  }

  // tools/list: forward to the extension, then append the bridge-local tools it
  // doesn't know about (saveMessage runs here, not in the extension). Guard the
  // shape so a malformed extension response is returned untouched.
  // saveMessage is advertised only when explicitly enabled -- a disabled tool
  // the model cannot see is a tool prompt injection cannot reach for.
  if (message.method === 'tools/list') {
    const resp = await forward(message);
    if (saveMessageEnabled(env) && resp && resp.result && Array.isArray(resp.result.tools)) {
      resp.result.tools.push(SAVE_MESSAGE_TOOL);
    }
    return resp;
  }

  // saveMessage is a bridge-local tool that writes files on the machine running
  // the MCP. Handle it here so it is never forwarded to the extension.
  if (message.method === 'tools/call'
      && message.params
      && message.params.name === 'saveMessage') {
    try {
      const data = await saveMessageToDisk({
        args: message.params.arguments || {},
        fetchMessage: forward,
        env
      });
      return {
        jsonrpc: '2.0',
        id: message.id,
        result: { content: [{ type: 'text', text: JSON.stringify(data) }] }
      };
    } catch (e) {
      return {
        jsonrpc: '2.0',
        id: message.id,
        error: { code: -32602, message: e.message }
      };
    }
  }

  return forwardToThunderbird(message);
}

function tryRequest(hostname, postData, port, token) {
  return new Promise((resolve, reject) => {
    const headers = {
      'Content-Type': 'application/json',
      'Content-Length': Buffer.byteLength(postData)
    };
    if (token) {
      headers['Authorization'] = `Bearer ${token}`;
    }
    const req = http.request({
      hostname,
      port,
      path: '/',
      method: 'POST',
      headers
    }, (res) => {
      const chunks = [];
      res.on('data', (chunk) => chunks.push(chunk));
      res.on('end', () => {
        if (res.statusCode === 403) {
          const err = new Error('Authentication failed (403). Token may be stale.');
          err.statusCode = 403;
          reject(err);
          return;
        }
        const data = Buffer.concat(chunks).toString('utf8');
        try {
          resolve(JSON.parse(data));
        } catch {
          try {
            resolve(JSON.parse(sanitizeJson(data)));
          } catch (e) {
            reject(new Error(`Invalid JSON from Thunderbird: ${e.message}`));
          }
        }
      });
    });

    req.on('error', reject);

    req.setTimeout(REQUEST_TIMEOUT, () => {
      req.destroy();
      reject(new Error('Request to Thunderbird timed out'));
    });

    req.write(postData);
    req.end();
  });
}

function isRetryableConnectionError(err) {
  return err
    && (err.statusCode === 403
      || err.code === 'ECONNREFUSED'
      || err.code === 'EADDRNOTAVAIL'
      || err.code === 'EAFNOSUPPORT');
}

function tryAllHosts(hosts, postData, port, token) {
  const tryNext = ([hostname, ...rest]) => {
    return tryRequest(hostname, postData, port, token).catch((err) => {
      if (rest.length > 0 && (err.code === 'ECONNREFUSED' || err.code === 'EADDRNOTAVAIL')) {
        return tryNext(rest);
      }
      throw err;
    });
  };
  return tryNext(hosts);
}

function compactToolResultJsonText(response) {
  const content = response?.result?.content;
  if (!Array.isArray(content)) {
    return response;
  }

  let changed = false;
  const compactedContent = content.map((item) => {
    if (item?.type !== 'text' || typeof item.text !== 'string') {
      return item;
    }
    try {
      const compactedText = JSON.stringify(JSON.parse(item.text));
      if (compactedText === item.text) {
        return item;
      }
      changed = true;
      return { ...item, text: compactedText };
    } catch {
      // Non-JSON text content is already the compact representation.
      return item;
    }
  });

  if (!changed) {
    return response;
  }
  return { ...response, result: { ...response.result, content: compactedContent } };
}

async function forwardToThunderbird(message) {
  const postData = JSON.stringify(message);

  // Read connection info (port + auth token) from the file written by the extension.
  // Fail-closed: if no connection file exists, retry a few times (Thunderbird may
  // still be starting), then fail with an error. Never forward requests without
  // authentication.
  let connInfo = readConnectionInfo();
  if (!connInfo) {
    for (let attempt = 0; attempt < CONNECTION_MAX_RETRIES; attempt++) {
      await new Promise((resolve) => setTimeout(resolve, CONNECTION_RETRY_DELAY_MS));
      connInfo = readConnectionInfo();
      if (connInfo) {
        break;
      }
    }
    if (!connInfo) {
      throw new Error(buildConnectionDiscoveryErrorMessage());
    }
  }

  // Walk through the cached candidate list on retryable failures so a stale
  // connection.json can't permanently mask a live one further down the list.
  // After the cached list is exhausted, rediscover once before giving up.
  let rediscoveryAttempted = false;

  while (connInfo) {
    if (!connInfo.port || !connInfo.token) {
      throw new Error('Invalid connection file: missing port or token');
    }
    if (typeof connInfo.port !== 'number' || connInfo.port < 1 || connInfo.port > 65535 || !Number.isInteger(connInfo.port)) {
      throw new Error('Invalid connection file: port must be an integer between 1 and 65535');
    }
    if (!isValidAuthToken(connInfo.token)) {
      throw new Error('Invalid connection file: token must be 64 lowercase hex characters');
    }

    try {
      return await tryAllHosts(THUNDERBIRD_HOSTS, postData, connInfo.port, connInfo.token);
    } catch (err) {
      if (!isRetryableConnectionError(err)) {
        throw err;
      }

      const next = advanceToNextCandidate();
      if (next) {
        connInfo = next;
        continue;
      }

      if (!rediscoveryAttempted) {
        rediscoveryAttempted = true;
        clearConnectionCache();
        connInfo = readConnectionInfo();
        if (!connInfo) {
          throw new Error(`Connection failed: ${err.message}. Is Thunderbird running with the MCP extension?`, { cause: err });
        }
        continue;
      }

      throw new Error(`Connection failed: ${err.message}. Is Thunderbird running with the MCP extension?`, { cause: err });
    }
  }
}

function startBridge() {
  let pendingRequests = 0;
  let stdinClosed = false;

  debugLog(`startup version=${BRIDGE_VERSION} pid=${process.pid} platform=${process.platform}`);

  function checkExit() {
    if (stdinClosed && pendingRequests === 0) {
      debugLog('shutdown stdin-closed and no pending requests, exiting 0');
      process.exit(0);
    }
  }

  function writeOutput(data) {
    return new Promise((resolve) => {
      if (process.stdout.write(data)) {
        resolve();
      } else {
        process.stdout.once('drain', resolve);
      }
    });
  }

  function dispatch(line) {
    if (!line.trim()) {
      return;
    }

    let messageId = null;
    let messageMethod = null;
    try {
      const parsed = JSON.parse(line);
      messageId = parsed.id ?? null;
      messageMethod = parsed.method ?? null;
    } catch {
      // Leave as null when request cannot be parsed
    }

    debugLog(`recv method=${messageMethod} id=${messageId}`);

    pendingRequests++;
    handleMessage(line)
      .then(async (response) => {
        if (response !== null) {
          await writeOutput(JSON.stringify(compactToolResultJsonText(response)) + '\n');
          debugLog(`send id=${messageId} method=${messageMethod}`);
        }
      })
      .catch(async (err) => {
        debugLog(`error id=${messageId} method=${messageMethod} message=${err.message}`);
        await writeOutput(JSON.stringify({
          jsonrpc: '2.0',
          id: messageId,
          error: { code: -32700, message: `Bridge error: ${err.message}` }
        }) + '\n');
      })
      .finally(() => {
        pendingRequests--;
        checkExit();
      });
  }

  // Manual newline-delimited JSON parsing on raw stdin. The previous
  // readline-based implementation lost the initialize response under
  // Claude Desktop's Electron-spawned Node on Windows -- writes from
  // promise callbacks never made it back through the pipe. Reading raw
  // 'data' events with explicit utf8 encoding matches what the official
  // @modelcontextprotocol/sdk stdio transport does and works reliably.
  process.stdin.setEncoding('utf8');
  let buffer = '';
  process.stdin.on('data', (chunk) => {
    buffer += chunk;
    let idx;
    while ((idx = buffer.indexOf('\n')) !== -1) {
      const line = buffer.slice(0, idx).replace(/\r$/, '');
      buffer = buffer.slice(idx + 1);
      dispatch(line);
    }
  });
  process.stdin.on('end', () => {
    if (buffer.length > 0) {
      const tail = buffer.replace(/\r$/, '');
      buffer = '';
      dispatch(tail);
    }
    stdinClosed = true;
    checkExit();
  });

  process.on('SIGINT', () => process.exit(0));
  process.on('SIGTERM', () => process.exit(0));
}

if (require.main === module) {
  startBridge();
}

module.exports = {
  advanceToNextCandidate,
  buildCandidateGroups,
  buildConnectionDiscoveryErrorMessage,
  clearConnectionCache,
  createDiscoveryContext,
  discoverConnectionInfo,
  findFlatpakConnectionCandidates,
  findMacOsConnectionCandidates,
  findSnapConnectionCandidates,
  formatDiscoveryAttempts,
  compactToolResultJsonText,
  handleMessage,
  inlineAttachmentPaths,
  isSensitiveFilePath,
  isValidAuthToken,
  readConnectionInfo,
  resolveDestDir,
  sanitizeSaveFilename,
  SAVE_MESSAGE_ENV,
  SAVE_MESSAGE_TOOL,
  SAVE_ROOT_ENV,
  saveMessageEnabled,
  saveMessageToDisk,
  startBridge,
  attachmentLimits: {
    MAX_ATTACHMENT_BYTES,
    MAX_TOTAL_ATTACHMENT_BYTES,
    MAX_ATTACHMENTS_PER_MESSAGE,
  },
};
