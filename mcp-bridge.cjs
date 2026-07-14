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
const { execFile, spawn } = require('child_process');

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

// ── Lifecycle tools ──────────────────────────────────────────────────────────
// Bridge-local tools for managing the Thunderbird process. These are handled
// entirely within the bridge and never forwarded to the extension.
const LIFECYCLE_LAUNCH_POLL_MS = 2000;
const LIFECYCLE_LAUNCH_MAX_POLLS = 15; // 15 × 2s = 30s max wait

// Module-level stdout writer set by startBridge(). Used by lifecycle tools
// to emit notifications/tools/list_changed after launch/stop.
let _writeNotification = null;

const LIFECYCLE_TOOL_SCHEMAS = [
  {
    name: 'launchThunderbird',
    description:
      'Start Thunderbird if it is not already running. ' +
      'Waits for the MCP extension to become reachable (up to ~30 s). ' +
      'Returns immediately if Thunderbird is already running.',
    inputSchema: { type: 'object', properties: {}, additionalProperties: false },
  },
  {
    name: 'stopThunderbird',
    description:
      'Gracefully stop the running Thunderbird process. ' +
      'IMPORTANT: Always ask the user for explicit confirmation before calling this tool — ' +
      'the user may have unsaved work (e.g. composing an email). ' +
      'No-op if Thunderbird is not running.',
    inputSchema: { type: 'object', properties: {}, additionalProperties: false },
  },
  {
    name: 'thunderbirdStatus',
    description:
      'Check whether Thunderbird is running and whether the MCP extension ' +
      'is reachable. Returns { running, extensionReachable, port? }.',
    inputSchema: { type: 'object', properties: {}, additionalProperties: false },
  },
];

const LIFECYCLE_TOOL_NAMES = new Set(LIFECYCLE_TOOL_SCHEMAS.map((t) => t.name));

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

// Read a file path off the host filesystem and convert it to the inline
// { name, contentType, base64 } shape the extension already supports. Throws
// a clear error on missing/unreadable/oversized files so the caller can
// surface them as MCP errors instead of the extension's silent
// "failed to attach" warning.
async function readAttachmentFromPath(filePath) {
  let stat;
  try {
    stat = await fs.promises.stat(filePath);
  } catch (e) {
    if (e.code === 'ENOENT') throw new Error(`Attachment not found: ${filePath}`, { cause: e });
    if (e.code === 'EACCES') throw new Error(`Attachment unreadable (permission denied): ${filePath}`, { cause: e });
    throw new Error(`Attachment stat failed (${e.code || 'unknown'}): ${filePath}`, { cause: e });
  }
  if (!stat.isFile()) {
    throw new Error(`Attachment is not a regular file: ${filePath}`);
  }
  if (stat.size > MAX_ATTACHMENT_BYTES) {
    throw new Error(
      `Attachment too large: ${filePath} is ${stat.size} bytes ` +
      `(limit ${MAX_ATTACHMENT_BYTES} bytes / ${MAX_ATTACHMENT_BYTES / 1024 / 1024} MB raw before base64)`
    );
  }
  let buf;
  try {
    buf = await fs.promises.readFile(filePath);
  } catch (e) {
    throw new Error(`Attachment read failed (${e.code || 'unknown'}): ${filePath}`, { cause: e });
  }
  return {
    name: path.basename(filePath),
    contentType: guessContentType(filePath),
    base64: buf.toString('base64')
  };
}

// Replace every string entry in `args.attachments` (= file path) with an
// inline { name, contentType, base64 } object read off the host filesystem.
// Inline objects pass through unchanged. Reads run in parallel; the first
// failing path rejects the whole call so the bridge fails loud instead of
// silently dropping an attachment.
async function inlineAttachmentPaths(args) {
  if (!args || !Array.isArray(args.attachments)) return;
  const resolved = await Promise.all(
    args.attachments.map(entry =>
      typeof entry === 'string' ? readAttachmentFromPath(entry) : entry
    )
  );
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

// ── Lifecycle tool helpers ────────────────────────────────────────────────────

function execCommand(cmd, args) {
  return new Promise((resolve) => {
    execFile(cmd, args, { timeout: 5000 }, (err, stdout) => {
      resolve({ ok: !err, stdout: (stdout || '').trim() });
    });
  });
}

// Process names to check across platforms. On Linux the binary can be
// 'thunderbird' or 'thunderbird-bin' (Arch, Snap, Flatpak). On macOS the
// Finder name is 'Thunderbird' but the actual process is 'thunderbird' or
// 'thunderbird-bin'. pgrep -x is case-sensitive and doesn't support regex,
// so we try each name individually.
const THUNDERBIRD_PROCESS_NAMES = ['thunderbird', 'thunderbird-bin'];

async function isThunderbirdRunning() {
  if (process.platform === 'win32') {
    const { ok, stdout } = await execCommand('tasklist', [
      '/FI', 'IMAGENAME eq thunderbird.exe', '/NH',
    ]);
    return ok && stdout.toLowerCase().includes('thunderbird.exe');
  }
  for (const name of THUNDERBIRD_PROCESS_NAMES) {
    const { ok } = await execCommand('pgrep', ['-x', name]);
    if (ok) return true;
  }
  return false;
}

function spawnThunderbird() {
  let proc;
  if (process.platform === 'darwin') {
    // 'open -a' launches the macOS app bundle and returns immediately.
    // The app runs independently of this process.
    proc = spawn('open', ['-a', 'Thunderbird'], {
      detached: true, stdio: 'ignore',
    });
  } else if (process.platform === 'win32') {
    // On Windows, Thunderbird registers its install location in the registry.
    // 'start ""' searches the App Paths registry key and PATH, so it finds
    // Thunderbird regardless of install directory (default: Program Files).
    proc = spawn('cmd', ['/c', 'start', '', 'thunderbird.exe'], {
      detached: true, stdio: 'ignore',
    });
  } else {
    proc = spawn('thunderbird', [], {
      detached: true, stdio: 'ignore',
    });
  }
  proc.unref();
  return proc;
}

async function stopThunderbirdProcess() {
  if (process.platform === 'win32') {
    // Try graceful WM_CLOSE first, fall back to forceful /F if needed.
    const { ok } = await execCommand('taskkill', ['/IM', 'thunderbird.exe']);
    if (ok) return;
    await execCommand('taskkill', ['/F', '/IM', 'thunderbird.exe']);
    return;
  }
  if (process.platform === 'darwin') {
    // AppleScript quit is the idiomatic way to close a macOS app — it lets
    // Thunderbird save state and run its shutdown sequence, same as Cmd-Q.
    const { ok } = await execCommand('osascript', [
      '-e', 'tell application "Thunderbird" to quit',
    ]);
    if (ok) return;
  }
  // Linux, or macOS osascript fallback: try SIGTERM via pkill for each
  // known process name.
  for (const name of THUNDERBIRD_PROCESS_NAMES) {
    await execCommand('pkill', ['-x', name]);
  }
}

function emitToolsListChanged() {
  if (_writeNotification) {
    debugLog('lifecycle: emitting notifications/tools/list_changed');
    _writeNotification({
      jsonrpc: '2.0',
      method: 'notifications/tools/list_changed',
    });
  }
}

async function handleLifecycleTool(name) {
  switch (name) {
    case 'thunderbirdStatus': {
      const running = await isThunderbirdRunning();
      const connInfo = readConnectionInfo();
      const result = {
        running,
        extensionReachable: !!(connInfo && connInfo.port && connInfo.token),
      };
      if (connInfo && connInfo.port) {
        result.port = connInfo.port;
      }
      return JSON.stringify(result);
    }

    case 'launchThunderbird': {
      if (await isThunderbirdRunning()) {
        const connInfo = readConnectionInfo();
        if (connInfo) {
          return JSON.stringify({ launched: false, message: 'Thunderbird is already running', extensionReachable: true });
        }
        return JSON.stringify({ launched: false, message: 'Thunderbird is running but the MCP extension is not reachable yet' });
      }

      debugLog('lifecycle: launching Thunderbird');
      spawnThunderbird();

      // Poll for the connection file (extension takes a few seconds to start)
      for (let i = 0; i < LIFECYCLE_LAUNCH_MAX_POLLS; i++) {
        await new Promise((resolve) => setTimeout(resolve, LIFECYCLE_LAUNCH_POLL_MS));
        clearConnectionCache();
        const connInfo = readConnectionInfo();
        if (connInfo) {
          debugLog(`lifecycle: extension reachable after ${(i + 1) * LIFECYCLE_LAUNCH_POLL_MS}ms`);
          emitToolsListChanged();
          return JSON.stringify({
            launched: true,
            message: 'Thunderbird started and MCP extension is reachable. '
              + 'The full set of email, calendar, and contacts tools is now available. '
              + 'You should re-discover available tools to access them.',
            port: connInfo.port,
          });
        }
      }

      // Thunderbird started but extension didn't become reachable in time
      const running = await isThunderbirdRunning();
      return JSON.stringify({
        launched: true,
        message: running
          ? 'Thunderbird started but the MCP extension did not become reachable within the timeout. It may still be loading.'
          : 'Thunderbird process exited unexpectedly',
        extensionReachable: false,
      });
    }

    case 'stopThunderbird': {
      if (!(await isThunderbirdRunning())) {
        return JSON.stringify({ stopped: false, message: 'Thunderbird is not running' });
      }
      debugLog('lifecycle: stopping Thunderbird');
      await stopThunderbirdProcess();
      // Give the process a moment to exit, then verify
      await new Promise((resolve) => setTimeout(resolve, 1000));
      clearConnectionCache();
      const stillRunning = await isThunderbirdRunning();
      if (!stillRunning) {
        emitToolsListChanged();
      }
      return JSON.stringify({
        stopped: !stillRunning,
        message: stillRunning
          ? 'Failed to stop Thunderbird'
          : 'Thunderbird stopped. Email, calendar, and contacts tools are no longer available.',
      });
    }

    default:
      throw new Error(`Unknown lifecycle tool: ${name}`);
  }
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
          capabilities: { tools: { listChanged: true } },
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

  // Handle bridge-local lifecycle tool calls without forwarding to Thunderbird.
  if (message.method === 'tools/call'
      && message.params
      && LIFECYCLE_TOOL_NAMES.has(message.params.name)) {
    try {
      const text = await handleLifecycleTool(message.params.name);
      return {
        jsonrpc: '2.0',
        id: message.id,
        result: { content: [{ type: 'text', text }] },
      };
    } catch (e) {
      return {
        jsonrpc: '2.0',
        id: message.id,
        error: { code: -32603, message: e.message },
      };
    }
  }

  // For tools/list, try to get the extension's tools and merge lifecycle tools.
  // If Thunderbird is not reachable, return only the lifecycle tools so the
  // client can still discover launchThunderbird. Auth errors (403) are
  // propagated as-is — they indicate a real token mismatch, not a missing
  // Thunderbird instance.
  if (message.method === 'tools/list') {
    try {
      const response = await forwardToThunderbird(message);
      if (response.error) {
        // Forward errors from the extension (e.g. auth failures) verbatim,
        // but still append lifecycle tools so the client can recover.
        return response;
      }
      const extensionTools = response?.result?.tools || [];
      return {
        jsonrpc: '2.0',
        id: message.id,
        result: { tools: [...extensionTools, ...LIFECYCLE_TOOL_SCHEMAS] },
      };
    } catch (err) {
      // Only fall back to lifecycle tools when Thunderbird is genuinely
      // unreachable — connection discovery found no file at all, or the
      // process isn't listening (ECONNREFUSED). All other errors (corrupt
      // connection file, auth failures, invalid port/token) must propagate
      // so security invariants hold.
      //
      // forwardToThunderbird wraps low-level errors in a new Error with
      // message "Connection failed: <original>", so we check both the
      // wrapper message and the cause's error code.
      const msg = err.message || '';
      const causeCode = err.cause?.code;
      const isAuthError =
        /authentication failed/i.test(msg) ||
        causeCode === 403 ||
        err.cause?.statusCode === 403;
      const isUnreachable = !isAuthError && (
        /Connection discovery failed/i.test(msg) ||
        /Connection failed/i.test(msg) ||
        causeCode === 'ECONNREFUSED' ||
        causeCode === 'EADDRNOTAVAIL' ||
        causeCode === 'EAFNOSUPPORT' ||
        err.code === 'ECONNREFUSED' ||
        err.code === 'EADDRNOTAVAIL' ||
        err.code === 'EAFNOSUPPORT'
      );

      if (isUnreachable) {
        debugLog(`tools/list: Thunderbird not reachable (${msg}), returning lifecycle tools only`);
        return {
          jsonrpc: '2.0',
          id: message.id,
          result: { tools: [...LIFECYCLE_TOOL_SCHEMAS] },
        };
      }

      // Auth errors, corrupt connection files, etc. — propagate
      return {
        jsonrpc: '2.0',
        id: message.id,
        error: { code: -32700, message: `Bridge error: ${msg}` },
      };
    }
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

  // Expose writeOutput so lifecycle tools can emit notifications.
  _writeNotification = (notification) =>
    writeOutput(JSON.stringify(notification) + '\n');

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
  handleLifecycleTool,
  isThunderbirdRunning,
  isValidAuthToken,
  LIFECYCLE_TOOL_NAMES,
  LIFECYCLE_TOOL_SCHEMAS,
  readConnectionInfo,
  startBridge,
};
