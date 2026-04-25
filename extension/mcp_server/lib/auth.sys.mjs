/**
 * Auth + connection-info plumbing for the Thunderbird MCP server.
 *
 * Factory pattern: api.js calls makeAuth({ Services, Cc, Ci }) once per
 * start() invocation and gets back a small operational object. Services /
 * Cc / Ci are sandbox globals in api.js but not in .sys.mjs files, so they
 * MUST be passed in.
 *
 * The auth token itself is NOT held in this module -- it lives in the
 * dispatch closure created by start() so its lifetime is tied to the
 * server, not to module-import lifecycle.
 */

const SUBDIR = "thunderbird-mcp";
const CONN_FILENAME = "connection.json";

export function makeAuth({ Services, Cc, Ci }) {
  /**
   * Generate a 256-bit hex token via XPCOM RNG.
   * crypto.getRandomValues is unavailable in experiment-API scope.
   */
  function generateAuthToken() {
    const bytes = new Uint8Array(32);
    const rng = Cc["@mozilla.org/security/random-generator;1"]
      .createInstance(Ci.nsIRandomGenerator);
    const randomBytes = rng.generateRandomBytes(32);
    for (let i = 0; i < 32; i++) bytes[i] = randomBytes[i];
    return Array.from(bytes, b => b.toString(16).padStart(2, "0")).join("");
  }

  /**
   * Constant-time string comparison to prevent timing side-channel attacks.
   * Pure function; doesn't actually need the deps but lives here for cohesion.
   */
  function timingSafeEqual(a, b) {
    const aStr = String(a);
    const bStr = String(b);
    const len = Math.max(aStr.length, bStr.length);
    let result = aStr.length ^ bStr.length;
    for (let i = 0; i < len; i++) {
      result |= (aStr.charCodeAt(i) || 0) ^ (bStr.charCodeAt(i) || 0);
    }
    return result === 0;
  }

  /**
   * Write port + auth token to <TmpD>/thunderbird-mcp/connection.json
   * with 0600 permissions, after defensive symlink checks.
   */
  function writeConnectionInfo(port, token) {
    const tmpDir = Services.dirsvc.get("TmpD", Ci.nsIFile);
    tmpDir.append(SUBDIR);
    if (!tmpDir.exists()) {
      tmpDir.create(Ci.nsIFile.DIRECTORY_TYPE, 0o700);
    } else if (tmpDir.isSymlink()) {
      throw new Error("thunderbird-mcp tmp directory is a symlink — refusing to write connection info");
    }
    const connFile = tmpDir.clone();
    connFile.append(CONN_FILENAME);
    // Symlink defense: remove any existing file first, then create with
    // O_CREAT|O_EXCL (0x08|0x80) to fail if a symlink appeared between
    // remove and create.
    if (connFile.exists()) {
      connFile.remove(false);
    }
    const data = JSON.stringify({ port, token, pid: Services.appinfo.processID });
    const ostream = Cc["@mozilla.org/network/file-output-stream;1"]
      .createInstance(Ci.nsIFileOutputStream);
    // 0x02 = O_WRONLY, 0x08 = O_CREAT, 0x80 = O_EXCL
    ostream.init(connFile, 0x02 | 0x08 | 0x80, 0o600, 0);
    const converter = Cc["@mozilla.org/intl/converter-output-stream;1"]
      .createInstance(Ci.nsIConverterOutputStream);
    converter.init(ostream, "UTF-8");
    converter.writeString(data);
    converter.close();
    return connFile.path;
  }

  /**
   * Best-effort cleanup of the connection info file on shutdown.
   */
  function removeConnectionInfo() {
    try {
      const tmpDir = Services.dirsvc.get("TmpD", Ci.nsIFile);
      tmpDir.append(SUBDIR);
      const connFile = tmpDir.clone();
      connFile.append(CONN_FILENAME);
      if (connFile.exists()) {
        connFile.remove(false);
      }
    } catch {
      // Best-effort cleanup
    }
  }

  return { generateAuthToken, timingSafeEqual, writeConnectionInfo, removeConnectionInfo };
}
