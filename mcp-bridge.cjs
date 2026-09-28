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
const CONNECTION_RETRY_DELAY_MS = 1000;
const CONNECTION_MAX_RETRIES = 5;
const CONNECTION_CACHE_TTL_MS = 5000; // 5 seconds

const DEFAULT_PROC_ROOT = '/proc';
const DEFAULT_DARWIN_FOLDERS_ROOT = '/var/folders';
const THUNDERBIRD_MCP_SUBDIR = 'thunderbird-mcp';
const SNAP_TMP_SUBDIR = ['Downloads', 'thunderbird.tmp'];
const CONNECTION_FILE_BASENAME = 'connection.json';
const MAX_CONNECTION_FILE_BYTES = 4096;
// Known upstream (including legacy and ESR), Fedora, and Betterbird IDs.
// Keep exact case: Flatpak application IDs are case-sensitive.
const FLATPAK_APP_IDS = new Set([
  'org.mozilla.Thunderbird',
  'org.mozilla.thunderbird',
  'org.mozilla.thunderbird_esr',
  'net.thunderbird.Thunderbird',
  'eu.betterbird.Betterbird',
]);
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
  const { fsImpl, pathImpl, darwinFoldersRoot } = context;
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
        // Scan metadata only; the opened descriptor decides trust below.
        addUniqueCandidate(candidates, seenPaths, makeCandidate('macOS temp scan', candidatePath, stat.mtimeMs));
      } catch (err) {
        if (err.code !== 'ENOENT' && err.code !== 'ENOTDIR') {
          continue;
        }
      }
    }
  }

  return buildScanGroup('macOS temp scan', pattern, candidates, 'no matching files');
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

        // argv[0] and TMPDIR are process-controlled; require the actual Snap executable.
        const executable = fsImpl.realpathSync(pathImpl.join(procRoot, pid, 'exe'));
        if (!executable.startsWith('/snap/thunderbird/')) continue;

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
    ...SNAP_TMP_SUBDIR,
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

// Shared layout definitions for discovery and attachment export recognition.
function getFlatpakScanRoots(context) {
  const { pathImpl, runtimeDir, homeDir } = context;
  return [
    {
      base: runtimeDir,
      label: '$XDG_RUNTIME_DIR',
      appRoot: pathImpl.join(runtimeDir || '$XDG_RUNTIME_DIR', 'app'),
      tmpSuffix: [],
    },
    {
      base: homeDir,
      label: '$HOME',
      appRoot: pathImpl.join(homeDir || '$HOME', '.var', 'app'),
      tmpSuffix: ['cache', 'tmp'],
    },
  ];
}

function findFlatpakConnectionCandidates(context) {
  const { fsImpl, pathImpl } = context;
  const roots = getFlatpakScanRoots(context);

  const notes = [];
  const candidates = [];
  const seenPaths = new Set();

  for (const root of roots) {
    const pattern = pathImpl.join(
      root.appRoot,
      '*',
      ...root.tmpSuffix,
      THUNDERBIRD_MCP_SUBDIR,
      CONNECTION_FILE_BASENAME
    );

    if (!root.base) {
      notes.push(makeAttempt('Flatpak scan', pattern, `${root.label} unavailable`));
      continue;
    }

    let appEntries;
    try {
      appEntries = fsImpl.readdirSync(root.appRoot, { withFileTypes: true });
    } catch (err) {
      notes.push(makeAttempt('Flatpak scan', pattern, normalizeFsError(err)));
      continue;
    }

    let found = 0;
    for (const appEntry of appEntries) {
      // Skip symlinked aliases (e.g. org.mozilla.Thunderbird -> net.thunderbird.Thunderbird)
      if (!appEntry.isDirectory()) {
        continue;
      }
      if (!FLATPAK_APP_IDS.has(appEntry.name)) {
        notes.push(makeAttempt('Flatpak scan', pathImpl.join(root.appRoot, appEntry.name), 'application ID not allowed'));
        continue;
      }

      const candidatePath = pathImpl.join(
        root.appRoot,
        appEntry.name,
        ...root.tmpSuffix,
        THUNDERBIRD_MCP_SUBDIR,
        CONNECTION_FILE_BASENAME
      );

      try {
        const stat = fsImpl.statSync(candidatePath);
        // Scan metadata only; the opened descriptor decides trust below.
        addUniqueCandidate(candidates, seenPaths, makeCandidate('Flatpak scan', candidatePath, stat.mtimeMs));
        found++;
      } catch {
        // ENOENT/ENOTDIR are expected for non-Thunderbird apps; skip anything else too.
      }
    }

    if (found === 0) {
      notes.push(makeAttempt('Flatpak scan', pattern, 'no matching files'));
    }
  }

  return { notes, candidates: sortCandidatesByMtime(candidates) };
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
  const { fsImpl, platform, uid } = context;
  let fd;
  try {
    const constants = fsImpl.constants || fs.constants;
    if (platform !== 'win32' && (!constants.O_NOFOLLOW || !constants.O_NONBLOCK)) {
      throw new Error('secure connection file open flags unavailable');
    }
    // Open once: no path-based read after checking ownership/type. NONBLOCK
    // lets fstat reject a FIFO without waiting for a writer to connect.
    fd = fsImpl.openSync(candidate.path,
      constants.O_RDONLY | (constants.O_NOFOLLOW || 0) | (constants.O_NONBLOCK || 0));
    const stat = fsImpl.fstatSync(fd);
    if (!stat.isFile()) throw new Error('connection file is not a regular file');
    if (platform !== 'win32') {
      if (!Number.isInteger(uid) || stat.uid !== uid) {
        throw new Error('connection file is not owned by the current user');
      }
      if ((stat.mode & 0o077) !== 0) {
        throw new Error('connection file permissions must be owner-only (0600)');
      }
    }
    if (!Number.isSafeInteger(stat.size) || stat.size < 0 || stat.size > MAX_CONNECTION_FILE_BYTES) {
      throw new Error(`connection file exceeds ${MAX_CONNECTION_FILE_BYTES} bytes or has an invalid size`);
    }
    // Bound the read as well as fstat: the file may grow after it was checked.
    const buffer = Buffer.alloc(MAX_CONNECTION_FILE_BYTES + 1);
    let length = 0;
    while (length < buffer.length) {
      const count = fsImpl.readSync(fd, buffer, length, buffer.length - length, null);
      if (count === 0) break;
      length += count;
    }
    if (length > MAX_CONNECTION_FILE_BYTES) {
      throw new Error(`connection file exceeds ${MAX_CONNECTION_FILE_BYTES} bytes`);
    }
    let data;
    try {
      data = JSON.parse(buffer.toString('utf8', 0, length));
    } catch {
      throw new Error('malformed JSON in connection file');
    }
    if (!data || !data.port || !data.token) {
      throw new Error('Invalid connection file: missing port or token');
    }
    if (!Number.isInteger(data.port) || data.port < 1 || data.port > 65535) {
      throw new Error('Invalid connection file: port must be an integer between 1 and 65535');
    }
    if (!isValidAuthToken(data.token)) {
      throw new Error('Invalid connection file: token must be 64 lowercase hex characters');
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
  } finally {
    if (fd !== undefined) fsImpl.closeSync(fd);
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
  // Network/device namespaces must be rejected before any filesystem access.
  /^\/\//,
  // macOS user Library remains denied even when used as a temp directory.
  /^\/users\/[^/]+\/library(\/|$)/,
  /\/thunderbird-mcp\/(?:[^/]+\/)?connection\.json$/,
  // Credential names also occur outside the usual profile directories.
  /(^|\/)id_[^/]+$/,
  /(^|\/)private[-_ ]?keys?(\.[^/]+)?$/,
  /\.(keychain|keychain-db)$/,
  /(^|\/)(web data|local state|signons\.sqlite|cert[89]\.db|pkcs11\.txt|secmod\.db|prefs\.js|profiles\.ini)$/,
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

function getAttachmentExportPathInfo(attachmentPath, exportRoots = [], windows = false) {
  // Backslashes are literal filename characters on POSIX, not separators.
  const nativePath = windows ? attachmentPath.replace(/\\/g, '/') : attachmentPath;
  if (nativePath.startsWith('//') || nativePath.split('/').some(part => part === '.' || part === '..')) return null;
  // getMessage exports exactly one sanitized message-id directory and one file.
  // The sibling "attachments" directory is outbound inline staging, not exports.
  const match = /^(.*\/thunderbird-mcp)\/([a-zA-Z0-9_]+)\/([^/]+)$/.exec(nativePath);
  if (!match || match[2].toLowerCase() === 'attachments' ||
      match[3].startsWith('.') || match[3].toLowerCase() === 'connection.json') return null;
  const roots = typeof exportRoots === 'function' ? exportRoots() : exportRoots;
  const root = roots.find(candidate => {
    const nativeRoot = (windows ? candidate.replace(/\\/g, '/') : candidate).replace(/\/$/, '');
    return windows ? nativeRoot.toLowerCase() === match[1].toLowerCase() : nativeRoot === match[1];
  });
  return root ? { root, parts: [match[2], match[3]] } : null;
}

function isSensitiveFilePath(attachmentPath, { windows = false, exportRoots = [] } = {}) {
  if (typeof attachmentPath !== 'string' || !attachmentPath) return false;
  const normalized = attachmentPath.replace(/\\/g, '/').toLowerCase();
  if (windows && (normalized.replace(/^[a-z]:/, '').includes(':') ||
      normalized.split('/').some(part => /[. ]$/.test(part)))) return true;
  // Traversal must never gain the export-directory exemption.
  if (normalized.split('/').some(part => part === '.' || part === '..')) return true;
  if (SENSITIVE_ATTACHMENT_PATTERNS.some(re => re.test(normalized))) return true;
  // Only inherited dot-directory/AppData restrictions may be waived for exports.
  if (/(^|\/)(\.[^/]*|appdata)(\/|$)/.test(normalized)) {
    return !getAttachmentExportPathInfo(attachmentPath, exportRoots, windows);
  }
  return false;
}

// Tools whose `attachments` array may contain string file paths that this
// bridge resolves on the host filesystem before forwarding. Needed because the
// Thunderbird snap (and other sandboxed installs) cannot see arbitrary host
// paths like /data/... or the host's /tmp; passing those paths through to the
// extension fails when file.exists() returns false inside the sandbox. Reading
// on the bridge side and shipping inline base64 sidesteps the sandbox entirely.
const ATTACHMENT_TOOLS = new Set(['sendMail', 'saveDraft', 'replyToMessage', 'forwardMessage']);

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

// Recognize only Thunderbird temp roots, never a root supplied in tool arguments.
function getAttachmentExportRoots(context) {
  const { pathImpl, homeDir, platform } = context;
  const roots = [pathImpl.dirname(getDefaultConnectionFile(context))];
  if (platform === 'linux') {
    if (homeDir) roots.push(pathImpl.join(homeDir, ...SNAP_TMP_SUBDIR, THUNDERBIRD_MCP_SUBDIR));
    for (const root of getFlatpakScanRoots(context)) {
      if (!root.base) continue;
      for (const appId of FLATPAK_APP_IDS) {
        roots.push(pathImpl.join(root.appRoot, appId, ...root.tmpSuffix, THUNDERBIRD_MCP_SUBDIR));
      }
    }
  }
  // Resolve only the trusted temp root. A redirect below it (including the
  // thunderbird-mcp directory itself) must not become a trusted export root.
  const canonicalRoots = roots.flatMap(root => {
    try {
      return [pathImpl.join(context.fsImpl.realpathSync(pathImpl.dirname(root)), THUNDERBIRD_MCP_SUBDIR)];
    } catch {
      return []; // An unavailable temp layout cannot provide an exemption.
    }
  });
  return [...new Set([...roots, ...canonicalRoots])];
}

async function inspectAttachmentPath(filePath, context) {
  const { fsImpl, pathImpl, attachmentPolicy } = context;
  // Check both the supplied path and its lexical normalization before any
  // filesystem access. The latter catches paths such as /tmp/../etc/passwd.
  if (isSensitiveFilePath(filePath, attachmentPolicy) || isSensitiveFilePath(pathImpl.resolve(filePath), attachmentPolicy)) {
    throw new Error(`Sensitive attachment path blocked: ${filePath}`);
  }

  let stat;
  try {
    // lstat is deliberate: stat would follow the final symlink before policy
    // could reject it.
    stat = await fsImpl.promises.lstat(filePath);
  } catch (e) {
    throw attachmentError('lstat', filePath, e);
  }
  validateAttachmentStat(filePath, stat);
  let realPath;
  try {
    realPath = fsImpl.realpathSync(filePath);
  } catch (e) {
    throw attachmentError('realpath', filePath, e);
  }
  if (isSensitiveFilePath(realPath, attachmentPolicy)) {
    throw new Error(`Sensitive attachment path blocked: ${filePath}`);
  }
  const exportInfo = getAttachmentExportPathInfo(filePath, attachmentPolicy.exportRoots, attachmentPolicy.windows);
  const resolvedExportInfo = getAttachmentExportPathInfo(realPath, attachmentPolicy.exportRoots, attachmentPolicy.windows);
  if (exportInfo || resolvedExportInfo) {
    let expectedPath = pathImpl.resolve(filePath);
    if (exportInfo) {
      const canonicalTempRoot = fsImpl.realpathSync(pathImpl.dirname(exportInfo.root));
      expectedPath = pathImpl.join(canonicalTempRoot, THUNDERBIRD_MCP_SUBDIR, ...exportInfo.parts);
    }
    if (pathImpl.relative(expectedPath, realPath) !== '') {
      throw new Error(`Attachment export path is redirected: ${filePath}`);
    }
  }
  return { filePath, realPath, stat, isExport: !!(exportInfo || resolvedExportInfo) };
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
async function readAttachmentFromPath(fileInfo, context) {
  const { fsImpl, pathImpl, platform, procRoot, attachmentPolicy } = context;
  const { filePath, stat: preflightStat } = fileInfo;
  const freshInfo = await inspectAttachmentPath(filePath, context);
  if (fileInfo.realPath !== freshInfo.realPath || !sameFile(preflightStat, freshInfo.stat) || preflightStat.size !== freshInfo.stat.size) {
    throw new Error(`Attachment changed after validation: ${filePath}`);
  }

  const constants = fsImpl.constants || fs.constants;
  const noFollow = constants.O_NOFOLLOW || 0;
  let handle;
  try {
    handle = await fsImpl.promises.open(freshInfo.realPath, constants.O_RDONLY | noFollow | (constants.O_NONBLOCK || 0));
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
    if (platform === 'linux') {
      let openedPath;
      try {
        openedPath = fsImpl.readlinkSync(pathImpl.join(procRoot, 'self', 'fd', String(handle.fd)));
      } catch (e) {
        throw attachmentError('resolve opened file', filePath, e);
      }
      if (!pathImpl.isAbsolute(openedPath) || openedPath.endsWith(' (deleted)') ||
          isSensitiveFilePath(openedPath, attachmentPolicy)) {
        throw new Error(`Sensitive or unresolved opened attachment path blocked: ${filePath}`);
      }
      if ((freshInfo.isExport ||
           getAttachmentExportPathInfo(openedPath, attachmentPolicy.exportRoots, attachmentPolicy.windows)) &&
          pathImpl.relative(freshInfo.realPath, openedPath) !== '') {
        throw new Error(`Attachment export path is redirected: ${filePath}`);
      }
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
async function inlineAttachmentPaths(args, options = {}) {
  if (!args || args.attachments === undefined || args.attachments === null) return;
  if (!Array.isArray(args.attachments)) {
    throw new Error('attachments must be an array, not a JSON-encoded string or another value');
  }
  const context = createDiscoveryContext(options);
  let exportRoots;
  context.attachmentPolicy = {
    windows: context.platform === 'win32',
    // Resolve trusted roots only after lexical network/credential checks pass.
    exportRoots: () => (exportRoots ||= getAttachmentExportRoots(context)),
  };

  if (args.attachments.length > MAX_ATTACHMENTS_PER_MESSAGE) {
    throw new Error(
      `Attachment count ${args.attachments.length} exceeds the ` +
      `${MAX_ATTACHMENTS_PER_MESSAGE} attachment limit`
    );
  }

  const fileInfoByIndex = new Map();
  const refused = [];
  let totalAttachmentBytes = 0;
  for (let index = 0; index < args.attachments.length; index++) {
    const entry = args.attachments[index];
    if (typeof entry !== 'string') continue;

    try {
      const fileInfo = await inspectAttachmentPath(entry, context);
      if (fileInfo.stat.size > MAX_TOTAL_ATTACHMENT_BYTES - totalAttachmentBytes) {
        throw new Error(
          `Attachment aggregate too large at ${entry}: exceeds the ` +
          `${MAX_TOTAL_ATTACHMENT_BYTES / 1024 / 1024} MB aggregate attachment limit`
        );
      }
      totalAttachmentBytes += fileInfo.stat.size;
      fileInfoByIndex.set(index, fileInfo);
    } catch (error) {
      refused.push(error.message);
    }
  }
  if (refused.length) throw new Error(`Attachments refused: ${refused.join('; ')}`);

  const resolved = [];
  for (let index = 0; index < args.attachments.length; index++) {
    const entry = args.attachments[index];
    resolved.push(
      typeof entry === 'string'
        ? await readAttachmentFromPath(fileInfoByIndex.get(index), context)
        : entry
    );
  }
  args.attachments = resolved;
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

async function handleMessage(line) {
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

  // For mail tools (including drafts), inline attachments passed as file paths.
  // The Thunderbird extension may run inside a sandboxed snap that cannot
  // see /data/..., the host /tmp, or any path outside its confined view —
  // passing such paths through would fail attachment validation.
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

  return forwardToThunderbird(message);
}

// BEGIN BRIDGE HTTP REQUEST HELPERS
const REQUEST_TIMEOUT = 30000;
const MAIL_OPERATION_TIMEOUT = 150000;

function getThunderbirdRequestPolicy(message) {
  let checkFolders = null;
  if (message?.method === 'tools/call') {
    const name = message.params?.name;
    const args = message.params?.arguments;
    // Match coerceToolArgs in the extension: only the exact string "true"
    // becomes true before validation and dispatch.
    const saveAsDraft = args?.saveAsDraft === true || args?.saveAsDraft === 'true';
    const skipReview = args?.skipReview === true || args?.skipReview === 'true';
    if (name === 'saveDraft' || (name === 'replyToMessage' && saveAsDraft)) {
      checkFolders = 'Drafts';
    } else if (['sendMail', 'replyToMessage', 'forwardMessage'].includes(name) && skipReview) {
      checkFolders = 'the Sent folder and the Outbox';
    }
  }
  return {
    timeoutMs: checkFolders ? MAIL_OPERATION_TIMEOUT : REQUEST_TIMEOUT,
    checkFolders,
  };
}

function tryRequest(hostname, postData, port, token, policy) {
  return new Promise((resolve, reject) => {
    let connected = false;
    let requestStarted = false;
    let settled = false;
    const settle = (error, response) => {
      if (settled) return;
      settled = true;
      if (error) reject(error);
      else resolve(response);
    };
    const connectionFailed = (error) => {
      // Writes are queued before connecting. A refused connection cannot have
      // dispatched the operation, but a lost connected socket may have done so
      // even when the entire request has not finished flushing yet.
      if (policy.checkFolders && connected && requestStarted) {
        error = new Error(
          `${error.message}. The outcome is unknown. Check ${policy.checkFolders} in Thunderbird before retrying.`,
          { cause: error }
        );
      }
      settle(error);
    };
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
      connected = true;
      const chunks = [];
      res.on('data', (chunk) => chunks.push(chunk));
      res.on('error', connectionFailed);
      res.on('aborted', () => connectionFailed(new Error('Response from Thunderbird was interrupted')));
      res.on('close', () => {
        if (!res.complete) connectionFailed(new Error('Response from Thunderbird closed before completion'));
      });
      res.on('end', () => {
        if (res.statusCode === 403) {
          const err = new Error('Authentication failed (403). Token may be stale.');
          err.statusCode = 403;
          settle(err);
          return;
        }
        const data = Buffer.concat(chunks).toString('utf8');
        try {
          settle(null, JSON.parse(data));
        } catch {
          try {
            settle(null, JSON.parse(sanitizeJson(data)));
          } catch (e) {
            settle(new Error(`Invalid JSON from Thunderbird: ${e.message}`));
          }
        }
      });
    });

    req.on('socket', (socket) => {
      if (socket.connecting) socket.once('connect', () => { connected = true; });
      else connected = true;
    });
    req.on('error', connectionFailed);
    req.on('close', () => connectionFailed(new Error('Connection to Thunderbird closed before a complete response')));

    req.setTimeout(policy.timeoutMs, () => {
      connectionFailed(new Error(`Request to Thunderbird timed out after ${policy.timeoutMs / 1000} seconds`));
      req.destroy();
    });

    requestStarted = true;
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

function tryAllHosts(hosts, postData, port, token, policy) {
  const tryNext = ([hostname, ...rest]) => {
    return tryRequest(hostname, postData, port, token, policy).catch((err) => {
      if (rest.length > 0 && (err.code === 'ECONNREFUSED' || err.code === 'EADDRNOTAVAIL')) {
        return tryNext(rest);
      }
      throw err;
    });
  };
  return tryNext(hosts);
}
// END BRIDGE HTTP REQUEST HELPERS

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

// BEGIN BRIDGE FORWARDING
async function forwardToThunderbird(message) {
  const postData = JSON.stringify(message);
  const policy = getThunderbirdRequestPolicy(message);

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
    try {
      return await tryAllHosts(THUNDERBIRD_HOSTS, postData, connInfo.port, connInfo.token, policy);
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
// END BRIDGE FORWARDING

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

// BEGIN BRIDGE STARTUP
// Desktop clients may load this program with require() from a Node bootstrap,
// so require.main does not reliably identify a CLI invocation. Test/library
// consumers must opt out explicitly before requiring the module.
if (process.env.THUNDERBIRD_MCP_NO_AUTOSTART !== '1') {
  startBridge();
}
// END BRIDGE STARTUP

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
  inlineAttachmentPaths,
  isSensitiveFilePath,
  isValidAuthToken,
  readConnectionInfo,
  startBridge,
  attachmentLimits: {
    MAX_ATTACHMENT_BYTES,
    MAX_TOTAL_ATTACHMENT_BYTES,
    MAX_ATTACHMENTS_PER_MESSAGE,
  },
};
