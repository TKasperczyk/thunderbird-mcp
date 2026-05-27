/**
 * Compose domain module for the Thunderbird MCP server.
 *
 * Implements the outgoing-mail and template tools: sendMail (composeMail),
 * replyToMessage, forwardMessage, saveDraft, dryRunCompose, listTemplates,
 * renderTemplate -- plus their internal helpers (identity lookup, attachment
 * descriptor building, compose-window customization, headless send, recipient
 * merging, body formatting, and template loading/parsing) and the
 * compose-related constants.
 *
 * Loaded by api.js via Services.scriptloader.loadSubScript with a
 * { module: { exports: {} } } scope, exactly like security_helpers.js and the
 * other domain modules. register(ctx) destructures the shared dependencies it
 * needs from ctx -- the XPCOM bindings (Cc/Ci/Services/MailServices), the
 * audit/idempotency/access-control infra helpers, the shared body and message
 * lookup helpers (extractPlainTextBody, findMessage) that MAIL also uses, the
 * shutdown-tracked temp-attachment / timer / claimed-window state objects that
 * also have to stay reachable from api.js's shutdown cleanup, and the
 * header/path sanitizers from security_helpers.js -- then defines the
 * functions verbatim and assigns the tool functions back onto ctx for the
 * dispatch switch.
 *
 * NOTE: function bodies are copied byte-for-byte from the monolithic api.js
 * (original indentation preserved) so runtime behavior is identical.
 */
"use strict";

module.exports = function register(ctx) {
  const {
    Cc,
    Ci,
    Services,
    MailServices,
    AUDIT_LOG_SUBDIR,
    appendComposeAudit,
    findIdempotentEntry,
    isAccountAllowed,
    isSkipReviewBlocked,
    getAccessibleAccounts,
    extractPlainTextBody,
    findMessage,
    isSensitiveFilePath,
    sanitizeHeaderLine,
    countRecipients,
    _attachTimers,
    _tempAttachFiles,
    _claimedComposeWindows,
  } = ctx;

  // ── Compose-related constants (used only by the compose/template tools). ──
  const MAX_BASE64_SIZE = 25 * 1024 * 1024; // 25 MB limit for inline base64 data (encoded)
  const MAX_FILE_PATH_ATTACHMENT_BYTES = 50 * 1024 * 1024;
  const COMPOSE_WINDOW_LOAD_DELAY_MS = 1500;
  const TEMPLATES_SUBDIR = "templates";
  // Monotonic counter for temp inline-attachment filenames. Persists for the
  // life of the loaded module (register runs once at server start), matching
  // the original file-scope `let _tempFileCounter = 0` semantics.
  let _tempFileCounter = 0;

            function templatesDir() {
              const profDir = Services.dirsvc.get("ProfD", Ci.nsIFile);
              const auditDir = profDir.clone();
              auditDir.append(AUDIT_LOG_SUBDIR);
              auditDir.append(TEMPLATES_SUBDIR);
              return auditDir;
            }

            function readTextFile(file) {
              const fis = Cc["@mozilla.org/network/file-input-stream;1"].createInstance(Ci.nsIFileInputStream);
              fis.init(file, 0x01, 0, 0);
              const cis = Cc["@mozilla.org/intl/converter-input-stream;1"].createInstance(Ci.nsIConverterInputStream);
              cis.init(fis, "UTF-8", 0, 0);
              let text = "";
              const buf = {};
              while (cis.readString(65536, buf) > 0) text += buf.value;
              cis.close();
              return text;
            }

            /**
             * Parse Jekyll-style frontmatter. Accepts a minimal YAML subset:
             * `key: value` per line for strings / numbers / booleans, and
             * `key: [a, b]` for one-line arrays. Anything more elaborate
             * (multi-line arrays, nested mappings) is intentionally not
             * supported -- the format stays predictable for the LLM.
             */
            function parseFrontmatter(raw) {
              if (!raw.startsWith("---")) return { meta: {}, body: raw };
              const end = raw.indexOf("\n---", 3);
              if (end < 0) return { meta: {}, body: raw };
              const yamlBlock = raw.slice(3, end).trim();
              let body = raw.slice(end + 4);
              if (body.startsWith("\n")) body = body.slice(1);
              const meta = Object.create(null);
              for (const line of yamlBlock.split(/\r?\n/)) {
                const stripped = line.trim();
                if (!stripped || stripped.startsWith("#")) continue;
                const m = stripped.match(/^([A-Za-z_][A-Za-z0-9_-]*)\s*:\s*(.*)$/);
                if (!m) continue;
                const key = m[1];
                let val = m[2].trim();
                if (val === "true") meta[key] = true;
                else if (val === "false") meta[key] = false;
                else if (/^-?\d+(\.\d+)?$/.test(val)) meta[key] = Number(val);
                else if (val.startsWith("[") && val.endsWith("]")) {
                  meta[key] = val.slice(1, -1).split(",").map(s => s.trim().replace(/^["']|["']$/g, "")).filter(Boolean);
                } else {
                  // Strip surrounding quotes if present
                  if ((val.startsWith('"') && val.endsWith('"')) || (val.startsWith("'") && val.endsWith("'"))) {
                    val = val.slice(1, -1);
                  }
                  meta[key] = val;
                }
              }
              return { meta, body };
            }

            function loadTemplate(name) {
              if (typeof name !== "string" || !name) return { error: "name must be a non-empty string" };
              if (!/^[A-Za-z0-9._-]+$/.test(name)) {
                return { error: "name must match /^[A-Za-z0-9._-]+$/ (no path separators)" };
              }
              const dir = templatesDir();
              if (!dir.exists()) {
                return { error: `Templates directory not found: ${dir.path}. Create it and add <name>.md files.` };
              }
              // Try <name>.md and <name> with no extension.
              for (const candidate of [name + ".md", name]) {
                const f = dir.clone();
                f.append(candidate);
                if (f.exists() && f.isFile()) {
                  try {
                    const raw = readTextFile(f);
                    const parsed = parseFrontmatter(raw);
                    parsed.file = f.path;
                    return parsed;
                  } catch (e) {
                    return { error: `Failed to read template '${name}': ${e}` };
                  }
                }
              }
              return { error: `Template not found: ${name}` };
            }

            function listTemplates() {
              const dir = templatesDir();
              if (!dir.exists()) {
                return { templates: [], note: `Templates directory does not exist yet. Create ${dir.path} and drop *.md files inside.` };
              }
              const out = [];
              let entries;
              try {
                entries = dir.directoryEntries;
              } catch (e) {
                return { error: `Failed to list templates: ${e}` };
              }
              while (entries.hasMoreElements()) {
                const f = entries.getNext().QueryInterface(Ci.nsIFile);
                if (!f.isFile()) continue;
                if (!/\.md$/i.test(f.leafName)) continue;
                try {
                  const raw = readTextFile(f);
                  const { meta } = parseFrontmatter(raw);
                  out.push({
                    name: typeof meta.name === "string" && meta.name ? meta.name : f.leafName.replace(/\.md$/i, ""),
                    description: typeof meta.description === "string" ? meta.description : "",
                    subject: typeof meta.subject === "string" ? meta.subject : "",
                    isHtml: !!meta.isHtml,
                    vars: Array.isArray(meta.vars) ? meta.vars : [],
                    file: f.path,
                  });
                } catch { /* skip unreadable templates */ }
              }
              return { templates: out };
            }

            function renderTemplate(name, vars) {
              const tpl = loadTemplate(name);
              if (tpl.error) return tpl;
              const declaredVars = Array.isArray(tpl.meta.vars) ? tpl.meta.vars : [];
              const bindings = (vars && typeof vars === "object" && !Array.isArray(vars)) ? vars : {};
              // Required-var enforcement: every declared var must be supplied.
              const missing = declaredVars.filter(v => !(v in bindings));
              if (missing.length > 0) {
                return { error: `Missing required variables: ${missing.join(", ")}` };
              }
              function substitute(text) {
                return String(text).replace(/\{\{\s*([A-Za-z_][A-Za-z0-9_]*)\s*\}\}/g, (_, key) => {
                  if (key in bindings) return String(bindings[key]);
                  // Unknown placeholder -> leave as literal so the caller
                  // notices instead of silently producing an empty value.
                  return `{{${key}}}`;
                });
              }
              return {
                name: tpl.meta.name || name,
                subject: substitute(tpl.meta.subject || ""),
                body: substitute(tpl.body),
                isHtml: !!tpl.meta.isHtml,
                file: tpl.file,
              };
            }

            /**
             * Searches the given accounts for an identity matching emailOrId
             * (by key or case-insensitive email).
             * Returns the identity object, or null if not found.
             */
            function findIdentityIn(accounts, emailOrId) {
              if (!emailOrId) return null;
              const lowerInput = emailOrId.toLowerCase();
              for (const account of accounts) {
                for (const identity of account.identities) {
                  if (identity.key === emailOrId || (identity.email || "").toLowerCase() === lowerInput) {
                    return identity;
                  }
                }
              }
              return null;
            }

            /**
             * Finds an identity by email address or identity ID
             * among accessible accounts only.  Returns null if not found.
             */
            function findIdentity(emailOrId) {
              return findIdentityIn(getAccessibleAccounts(), emailOrId);
            }

            /** Creates an nsIFile instance for the given path. */
            function createLocalFile(path) {
              const file = Cc["@mozilla.org/file/local;1"].createInstance(Ci.nsIFile);
              file.initWithPath(path);
              return file;
            }

            /**
             * Converts attachment entries to attachment descriptors.
             * Each entry can be:
             *   - A string (file path) — resolved from disk
             *   - An object { name, contentType, base64 } — decoded and written
             *     to a temp file under <TmpD>/thunderbird-mcp/attachments/
             * Returns { descs: [{url, name, size, contentType?}], failed: string[] }
             */
            function filePathsToAttachDescs(filePaths) {
              const descs = [];
              const failed = [];
              if (!filePaths || !Array.isArray(filePaths)) return { descs, failed };
              for (const entry of filePaths) {
                try {
                  if (typeof entry === "string") {
                    // File path attachment.
                    //
                    // SECURITY: reject paths that point at credentials, system
                    // files, or browser/mail profile data BEFORE touching the
                    // filesystem. This is the LLM-confused-deputy defense:
                    // attacker-controlled email content can prompt-inject an
                    // assistant into calling sendMail with attachments=["/path/to/id_rsa"]
                    // and we never want that to succeed regardless of skipReview.
                    if (isSensitiveFilePath(entry)) {
                      failed.push(`${entry} (sensitive path blocked)`);
                      continue;
                    }
                    const file = createLocalFile(entry);
                    if (!file.exists()) {
                      failed.push(entry);
                      continue;
                    }
                    // Size cap mirrors the saved-attachment ceiling and avoids
                    // ballooning outgoing messages when a caller points at a huge file.
                    let fileSize = 0;
                    try { fileSize = file.fileSize; } catch { fileSize = 0; }
                    if (fileSize > MAX_FILE_PATH_ATTACHMENT_BYTES) {
                      failed.push(`${entry} (exceeds ${MAX_FILE_PATH_ATTACHMENT_BYTES / 1024 / 1024}MB size limit)`);
                      continue;
                    }
                    descs.push({ url: Services.io.newFileURI(file).spec, name: file.leafName, size: fileSize });
                  } else if (entry && typeof entry === "object" && (entry.base64 || entry.content) && entry.name) {
                    // Inline base64 attachment — decode and write to temp file
                    const b64Data = entry.base64 || entry.content;
                    if (b64Data.length > MAX_BASE64_SIZE) {
                      failed.push(`${entry.name} (exceeds ${MAX_BASE64_SIZE / 1024 / 1024}MB size limit)`);
                      continue;
                    }
                    // Decode base64 to binary bytes
                    let bytes;
                    try {
                      // Use the global atob when available, otherwise fall back
                      const raw = typeof atob === "function" ? atob(b64Data) : ChromeUtils.base64Decode(b64Data);
                      bytes = new Uint8Array(raw.length);
                      for (let i = 0; i < raw.length; i++) bytes[i] = raw.charCodeAt(i);
                    } catch {
                      // Fallback: manual base64 decode (atob may not be available in XPCOM context)
                      try {
                        const lookup = new Uint8Array(256);
                        const chars = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/";
                        for (let i = 0; i < chars.length; i++) lookup[chars.charCodeAt(i)] = i;
                        const clean = b64Data.replace(/[^A-Za-z0-9+/]/g, "");
                        const len = clean.length;
                        const outLen = (len * 3) >> 2;
                        bytes = new Uint8Array(outLen);
                        let p = 0;
                        for (let i = 0; i < len; i += 4) {
                          const a = lookup[clean.charCodeAt(i)];
                          const b = lookup[clean.charCodeAt(i + 1)];
                          const c = lookup[clean.charCodeAt(i + 2)];
                          const d = lookup[clean.charCodeAt(i + 3)];
                          bytes[p++] = (a << 2) | (b >> 4);
                          if (i + 2 < len) bytes[p++] = ((b & 15) << 4) | (c >> 2);
                          if (i + 3 < len) bytes[p++] = ((c & 3) << 6) | d;
                        }
                        bytes = bytes.subarray(0, p);
                      } catch {
                        failed.push(`${entry.name} (invalid base64 data)`);
                        continue;
                      }
                    }
                    const tmpDir = Services.dirsvc.get("TmpD", Ci.nsIFile);
                    tmpDir.append("thunderbird-mcp");
                    tmpDir.append("attachments");
                    if (!tmpDir.exists()) {
                      tmpDir.create(Ci.nsIFile.DIRECTORY_TYPE, 0o700);
                    }
                    const tmpFile = tmpDir.clone();
                    let safeName = (entry.name || entry.filename || "attachment").replace(/[^a-zA-Z0-9._-]/g, "_");
                    if (!safeName || safeName === "." || safeName === "..") safeName = "attachment";
                    tmpFile.append(`${Date.now()}_${++_tempFileCounter}_${safeName}`);
                    // Write via XPCOM binary stream
                    const ostream = Cc["@mozilla.org/network/file-output-stream;1"]
                      .createInstance(Ci.nsIFileOutputStream);
                    ostream.init(tmpFile, 0x02 | 0x08 | 0x20, 0o600, 0);
                    const bstream = Cc["@mozilla.org/binaryoutputstream;1"]
                      .createInstance(Ci.nsIBinaryOutputStream);
                    bstream.setOutputStream(ostream);
                    bstream.writeByteArray(bytes, bytes.length);
                    bstream.close();
                    ostream.close();
                    _tempAttachFiles.add(tmpFile.path);
                    const desc = { url: Services.io.newFileURI(tmpFile).spec, name: entry.name || entry.filename, size: tmpFile.fileSize };
                    if (entry.contentType) desc.contentType = entry.contentType;
                    descs.push(desc);
                  } else {
                    failed.push(typeof entry === "object" ? JSON.stringify(entry) : String(entry));
                  }
                } catch (e) {
                  failed.push(typeof entry === "object" ? (entry.name || JSON.stringify(entry)) : String(entry));
                }
              }
              return { descs, failed };
            }

            /**
             * Injects attachment descriptors into the most recently opened compose window.
             * Uses nsITimer so the window has time to finish loading before injection.
             * Each call gets its own timer stored in _attachTimers to prevent GC.
             *
             * Known limitation: uses getMostRecentWindow("msgcompose") which is a race
             * if two compose operations happen within COMPOSE_WINDOW_LOAD_DELAY_MS --
             * attachments from the first may land on the second window.
             * OpenComposeWindowWithParams doesn't return a window handle, so there's
             * no reliable way to target a specific window. Injection failures are
             * silent (callers report success based on pre-validated descriptor counts).
             */
            /**
             * Converts attachment descriptors to nsIMsgAttachment objects.
             * Shared by injectAttachmentsAsync (compose window) and
             * sendMessageDirectly (headless send).
             */
            function descsToMsgAttachments(attachDescs) {
              const result = [];
              for (const desc of attachDescs) {
                try {
                  const att = Cc["@mozilla.org/messengercompose/attachment;1"]
                    .createInstance(Ci.nsIMsgAttachment);
                  att.url = desc.url;
                  att.name = desc.name;
                  if (desc.size != null) att.size = desc.size;
                  if (desc.contentType) att.contentType = desc.contentType;
                  result.push(att);
                } catch (e) {
                  console.warn("thunderbird-mcp: failed to convert attachment descriptor:", desc?.name || desc?.url || desc, e);
                }
              }
              return result;
            }

            function addAttachmentsToComposeWindow(composeWin, attachDescs) {
              if (!composeWin) {
                console.warn("thunderbird-mcp: skipping attachment add — no compose window");
                return;
              }
              if (typeof composeWin.AddAttachments !== "function") {
                console.warn("thunderbird-mcp: skipping attachment add — composeWin.AddAttachments not a function");
                return;
              }
              const attachList = descsToMsgAttachments(attachDescs);
              if (attachList.length > 0) {
                // Caller's try/catch (or fire-and-forget caller) is responsible for
                // surfacing failures — do not swallow here.
                composeWin.AddAttachments(attachList);
              }
            }

            function injectAttachmentsAsync(attachDescs) {
              if (!attachDescs || attachDescs.length === 0) return;
              const timer = Cc["@mozilla.org/timer;1"].createInstance(Ci.nsITimer);
              _attachTimers.add(timer);
              timer.initWithCallback({
                notify() {
                  _attachTimers.delete(timer);
                  try {
                    const composeWin = Services.wm.getMostRecentWindow("msgcompose");
                    addAttachmentsToComposeWindow(composeWin, attachDescs);
                  } catch (e) {
                    // Fire-and-forget timer — no client to error back to.
                    // Log loudly so users can find the cause in the Error Console
                    // when an "email sent without attachments" report comes in.
                    console.error("thunderbird-mcp: injectAttachmentsAsync failed:", e);
                  }
                }
              }, COMPOSE_WINDOW_LOAD_DELAY_MS, Ci.nsITimer.TYPE_ONE_SHOT);
            }

            function splitAddressHeader(header) {
              return (header || "").match(/(?:[^,"]|"[^"]*")+/g) || [];
            }

            function extractAddressEmail(address) {
              return (address.match(/<([^>]+)>/)?.[1] || address.trim()).toLowerCase();
            }

            function mergeAddressHeaders(...headers) {
              const seen = new Set();
              const merged = [];
              for (const header of headers) {
                for (const raw of splitAddressHeader(header)) {
                  const address = raw.trim();
                  if (!address) continue;
                  const email = extractAddressEmail(address);
                  if (seen.has(email)) continue;
                  seen.add(email);
                  merged.push(address);
                }
              }
              return merged.join(", ");
            }

            function getReplyAllCcRecipients(msgHdr, folder) {
              const ownAccount = MailServices.accounts.findAccountForServer(folder.server);
              const ownEmails = new Set();
              if (ownAccount) {
                for (const identity of ownAccount.identities) {
                  if (identity.email) ownEmails.add(identity.email.toLowerCase());
                }
              }

              const allRecipients = [
                ...splitAddressHeader(msgHdr.recipients),
                ...splitAddressHeader(msgHdr.ccList)
              ]
                .map(r => r.trim())
                .filter(r => r && (ownEmails.size === 0 || !ownEmails.has(extractAddressEmail(r))));

              const seen = new Set();
              const uniqueRecipients = allRecipients.filter(r => {
                const email = extractAddressEmail(r);
                if (seen.has(email)) return false;
                seen.add(email);
                return true;
              });

              return uniqueRecipients.join(", ");
            }

            function getIdentityAutoRecipientHeader(identity, kind) {
              if (!identity) return "";
              try {
                if (kind === "cc") {
                  return identity.doCc ? (identity.doCcList || "") : "";
                }
                if (kind === "bcc") {
                  return identity.doBcc ? (identity.doBccList || "") : "";
                }
              } catch {}
              return "";
            }

            function applyComposeRecipientOverrides(composeWin, identity, to, cc, bcc) {
              if (!composeWin) return;
              // Build a recipients-only delta. Do NOT pass identityKey here --
              // the compose window already has the identity from
              // msgComposeParams, and including `identityKey: null` is sketchy:
              // modern Thunderbird ignores it (`if (details.identityKey)`
              // short-circuits on null), but a future TB version could
              // interpret it as "clear the identity", which would also wipe
              // the OpenPGP signing/encrypting state that depends on it.
              const overrides = {};
              if (to) overrides.to = to;
              if (cc) overrides.cc = mergeAddressHeaders(getIdentityAutoRecipientHeader(identity, "cc"), cc);
              if (bcc) overrides.bcc = mergeAddressHeaders(getIdentityAutoRecipientHeader(identity, "bcc"), bcc);
              if (Object.keys(overrides).length === 0) return;

              if (typeof composeWin.SetComposeDetails === "function") {
                composeWin.SetComposeDetails(overrides);
                return;
              }

              const fields = composeWin.gMsgCompose?.compFields;
              if (!fields) return;
              if (Object.prototype.hasOwnProperty.call(overrides, "to")) fields.to = overrides.to;
              if (Object.prototype.hasOwnProperty.call(overrides, "cc")) fields.cc = overrides.cc;
              if (Object.prototype.hasOwnProperty.call(overrides, "bcc")) fields.bcc = overrides.bcc;
              if (typeof composeWin.CompFields2Recipients === "function") {
                composeWin.CompFields2Recipients(fields);
              }
            }

            function formatBodyFragmentHtml(body, isHtml) {
              const formatted = formatBodyHtml(body, isHtml);
              if (!isHtml) return formatted;
              if (!formatted) return "";

              const needsParsing = /<(?:html|body|head)\b/i.test(formatted) || /\bmoz-signature\b/i.test(formatted);
              if (!needsParsing) return formatted;

              try {
                const doc = new DOMParser().parseFromString(formatted, "text/html");
                for (const node of doc.querySelectorAll("div.moz-signature, pre.moz-signature")) {
                  node.remove();
                }
                return doc.body ? doc.body.innerHTML : formatted;
              } catch {
                return formatted;
              }
            }

            function moveComposeSelectionToBodyStartIfRange(composeWin) {
              const browser = typeof composeWin?.getBrowser === "function" ? composeWin.getBrowser() : null;
              const editorDoc = browser?.contentDocument;
              const root = editorDoc?.body;
              const selection = typeof editorDoc?.getSelection === "function" ? editorDoc.getSelection() : null;
              let shouldMove = false;
              let moveError = null;

              if (selection) {
                if (selection.isCollapsed) return true;
                shouldMove = true;
              }

              if (shouldMove && root && typeof editorDoc.createRange === "function") {
                try {
                  const range = editorDoc.createRange();
                  range.setStart(root, 0);
                  range.collapse(true);
                  selection.removeAllRanges();
                  selection.addRange(range);
                  return true;
                } catch (e) {
                  moveError = e;
                }
              }

              const editor = typeof composeWin?.GetCurrentEditor === "function" ? composeWin.GetCurrentEditor() : null;
              let editorSelection = null;
              if (!selection && editor) {
                try {
                  editorSelection = editor.selection || null;
                } catch (e) {
                  moveError = e;
                }
              }

              if (!selection && editorSelection) {
                try {
                  if (editorSelection.isCollapsed) return true;
                  shouldMove = true;
                } catch (e) {
                  moveError = e;
                }
              }

              if (!selection && !editorSelection) {
                // Legacy editor-only path does not expose a reliable DOM
                // selection, so anchor defensively before insertHTML.
                shouldMove = true;
              }

              if (shouldMove && editor && typeof editor.beginningOfDocument === "function") {
                try {
                  editor.beginningOfDocument();
                  return true;
                } catch (e) {
                  moveError = e;
                }
              }

              if (shouldMove) {
                console.warn("thunderbird-mcp: could not anchor compose body insertion; insertHTML may replace selected quote", moveError);
              }
              return !shouldMove;
            }

            function insertReplyBodyIntoComposeWindow(composeWin, body, isHtml) {
              if (!composeWin || !body) return;
              const fragment = formatBodyFragmentHtml(body, isHtml);
              if (!fragment) return;

              const browser = typeof composeWin.getBrowser === "function" ? composeWin.getBrowser() : null;
              const editorDoc = browser?.contentDocument;
              if (editorDoc && typeof editorDoc.execCommand === "function") {
                // Body-ready can leave the original quote selected. insertHTML
                // replaces active ranges, so anchor only when TB selected text.
                moveComposeSelectionToBodyStartIfRange(composeWin);
                editorDoc.execCommand("insertHTML", false, fragment);
              } else {
                const editor = typeof composeWin.GetCurrentEditor === "function" ? composeWin.GetCurrentEditor() : null;
                if (editor && typeof editor.insertHTML === "function") {
                  moveComposeSelectionToBodyStartIfRange(composeWin);
                  editor.insertHTML(fragment);
                }
              }

              if (composeWin.gMsgCompose) {
                composeWin.gMsgCompose.bodyModified = true;
              }
              if ("gContentChanged" in composeWin) {
                composeWin.gContentChanged = true;
              }
            }

            function openComposeWindowWithCustomizations(msgComposeParams, originalMsgURI, compType, identity, body, isHtml, to, cc, bcc, attachDescs) {
              return new Promise((resolve) => {
                const OPEN_TIMEOUT_MS = 15000;
                let settled = false;
                let matchedWindow = null;
                let pendingStateListener = null;
                let pendingStateCompose = null;

                const finish = (result) => {
                  if (settled) return;
                  settled = true;
                  try { Services.ww.unregisterNotification(windowObserver); } catch {}
                  try { timeout.cancel(); } catch {}
                  // Unregister any dangling state listener so a late
                  // NotifyComposeBodyReady cannot mutate the compose window
                  // after we have already resolved (e.g. after a timeout).
                  if (pendingStateListener && pendingStateCompose) {
                    try { pendingStateCompose.UnregisterStateListener(pendingStateListener); } catch {}
                  }
                  pendingStateListener = null;
                  pendingStateCompose = null;
                  resolve(result);
                };

                const timeout = Cc["@mozilla.org/timer;1"].createInstance(Ci.nsITimer);
                timeout.initWithCallback({
                  notify() {
                    finish({ error: "Timed out waiting for compose window" });
                  }
                }, OPEN_TIMEOUT_MS, Ci.nsITimer.TYPE_ONE_SHOT);

                const maybeCustomizeWindow = (composeWin) => {
                  try {
                    if (!composeWin || composeWin === matchedWindow) return;
                    if (composeWin.document?.documentElement?.getAttribute("windowtype") !== "msgcompose") return;
                    if (!composeWin.gMsgCompose) return;
                    if (composeWin.gMsgCompose.originalMsgURI !== originalMsgURI) return;
                    if (composeWin.gComposeType !== compType) return;
                    // When two callers reply to the same message concurrently,
                    // both observers see both compose windows. Skip any window
                    // that has already been claimed by a prior observer so each
                    // call binds to exactly one compose window.
                    if (_claimedComposeWindows.has(composeWin)) return;
                    _claimedComposeWindows.add(composeWin);

                    matchedWindow = composeWin;
                    try { Services.ww.unregisterNotification(windowObserver); } catch {}

                    const stateListener = {
                      QueryInterface: ChromeUtils.generateQI(["nsIMsgComposeStateListener"]),
                      NotifyComposeFieldsReady() {},
                      ComposeProcessDone() {},
                      SaveInFolderDone() {},
                      NotifyComposeBodyReady() {
                        // Guard against a late body-ready firing after the
                        // caller already timed out -- don't mutate the compose
                        // window once the promise is settled.
                        if (settled) {
                          try { composeWin.gMsgCompose.UnregisterStateListener(stateListener); } catch {}
                          return;
                        }

                        try {
                          composeWin.gMsgCompose.UnregisterStateListener(stateListener);
                        } catch {}
                        pendingStateListener = null;
                        pendingStateCompose = null;

                        try {
                          applyComposeRecipientOverrides(composeWin, identity, to, cc, bcc);
                          insertReplyBodyIntoComposeWindow(composeWin, body, isHtml);
                          addAttachmentsToComposeWindow(composeWin, attachDescs);
                          finish({ success: true });
                        } catch (e) {
                          finish({ error: e.toString() });
                        }
                      },
                    };

                    pendingStateListener = stateListener;
                    pendingStateCompose = composeWin.gMsgCompose;
                    composeWin.gMsgCompose.RegisterStateListener(stateListener);
                  } catch (e) {
                    finish({ error: e.toString() });
                  }
                };

                const windowObserver = {
                  observe(subject, topic) {
                    if (topic !== "domwindowopened") return;
                    const composeWin = subject;
                    if (!composeWin || typeof composeWin.addEventListener !== "function") return;

                    // Thunderbird dispatches a non-bubbling compose-window-init event
                    // from MsgComposeCommands.js after gMsgCompose is initialized and
                    // the built-in state listener is registered, but before editor
                    // creation begins. Capturing it on the window lets us register our
                    // own ComposeBodyReady listener for the specific reply window
                    // without relying on getMostRecentWindow("msgcompose").
                    composeWin.addEventListener("compose-window-init", () => {
                      maybeCustomizeWindow(composeWin);
                    }, { once: true, capture: true });
                  },
                };

                try {
                  Services.ww.registerNotification(windowObserver);
                  const msgComposeService = Cc["@mozilla.org/messengercompose;1"]
                    .getService(Ci.nsIMsgComposeService);
                  msgComposeService.OpenComposeWindowWithParams(null, msgComposeParams);
                } catch (e) {
                  finish({ error: e.toString() });
                }
              });
            }

            function markMessageDispositionState(msgHdr, dispositionState) {
              try {
                const folder = msgHdr?.folder;
                if (!folder || dispositionState == null) return false;
                if (typeof folder.addMessageDispositionState === "function") {
                  folder.addMessageDispositionState(msgHdr, dispositionState);
                  return true;
                }
                if (typeof folder.AddMessageDispositionState === "function") {
                  folder.AddMessageDispositionState(msgHdr, dispositionState);
                  return true;
                }
              } catch {}
              return false;
            }

            /**
             * Sends a message directly via nsIMsgSend without opening a compose window.
             * Used by composeMail, replyToMessage, forwardMessage when skipReview=true.
             *
             * Handles two createAndSendMessage signatures:
             * - TB 102-127 (C++): 18 args, includes aAttachments + aPreloadedAttachments
             * - TB 128+   (JS):  16 args, attachments via composeFields only
             * Attachments are always added to composeFields (works in both).
             * We try the modern 16-arg call first; if TB throws
             * NS_ERROR_XPC_NOT_ENOUGH_ARGS, fall back to the legacy 18-arg call.
             */
            function sendMessageDirectly(composeFields, identity, attachDescs, originalMsgURI, compType, deliverMode, bodyType) {
              if (!identity) {
                return Promise.resolve({ error: "No identity available for direct send" });
              }

              const mode = deliverMode ?? Ci.nsIMsgCompDeliverMode.Now;
              const bodyMimeType = bodyType || "text/html";
              const SEND_TIMEOUT_MS = 120000; // 2 min safety timeout

              return new Promise((resolve) => {
                let settled = false;
                const settle = (result) => {
                  if (!settled) {
                    settled = true;
                    resolve(result);
                  }
                };

                // Safety timeout -- if neither listener callback nor error fires
                const timer = Cc["@mozilla.org/timer;1"].createInstance(Ci.nsITimer);
                timer.initWithCallback({
                  notify() { settle({ error: "Send timed out after " + (SEND_TIMEOUT_MS / 1000) + "s" }); }
                }, SEND_TIMEOUT_MS, Ci.nsITimer.TYPE_ONE_SHOT);

                try {
                  const msgSend = Cc["@mozilla.org/messengercompose/send;1"]
                    .createInstance(Ci.nsIMsgSend);

                  // Populate sender fields from identity (normally done by compose window)
                  if (identity.email) {
                    const name = identity.fullName || "";
                    composeFields.from = name
                      ? `"${name}" <${identity.email}>`
                      : identity.email;
                  }
                  if (identity.organization) {
                    composeFields.organization = identity.organization;
                  }

                  // Add attachments to composeFields (works in all TB versions)
                  for (const att of descsToMsgAttachments(attachDescs)) {
                    composeFields.addAttachment(att);
                  }

                  // Extract body -- createAndSendMessage takes it as a separate param
                  const body = composeFields.body || "";

                  // Resolve account key from identity
                  let accountKey = "";
                  try {
                    for (const account of MailServices.accounts.accounts) {
                      for (let i = 0; i < account.identities.length; i++) {
                        if (account.identities[i].key === identity.key) {
                          accountKey = account.key;
                          break;
                        }
                      }
                      if (accountKey) break;
                    }
                  } catch {}

                  // SaveAsDraft mode routes through _mimeDoFcc() which does not
                  // call onStopSending. Completion is signaled via the copy
                  // service's onStopCopy after the message lands in the Drafts
                  // folder, so the listener also QIs nsIMsgCopyServiceListener.
                  const listener = {
                    QueryInterface: ChromeUtils.generateQI(["nsIMsgSendListener", "nsIMsgCopyServiceListener"]),
                    // nsIMsgSendListener -- fires for SMTP send paths
                    onStartSending() {},
                    onProgress() {},
                    onSendProgress() {},
                    onStatus() {},
                    onStopSending(msgID, status) {
                      timer.cancel();
                      if (Components.isSuccessCode(status)) {
                        settle({ success: true, message: "Message sent" });
                      } else {
                        settle({ error: `Send failed (status: 0x${status.toString(16)})` });
                      }
                    },
                    onGetDraftFolderURI() {},
                    onSendNotPerformed(msgID, status) {
                      timer.cancel();
                      settle({ error: "Send was not performed" });
                    },
                    onTransportSecurityError(msgID, status, secInfo, location) {
                      timer.cancel();
                      settle({ error: `Transport security error${location ? ": " + location : ""}` });
                    },
                    // nsIMsgCopyServiceListener -- fires for SaveAsDraft / Sent-folder copy
                    onStartCopy() {},
                    setMessageKey() {},
                    onStopCopy(status) {
                      timer.cancel();
                      if (Components.isSuccessCode(status)) {
                        settle({ success: true, message: "Saved" });
                      } else {
                        settle({ error: `Save failed (status: 0x${status.toString(16)})` });
                      }
                    },
                  };

                  // Common args shared by both signatures (positions 1-10)
                  const commonArgs = [
                    null,                           // editor
                    identity,                       // identity
                    accountKey,                     // account key
                    composeFields,                  // fields
                    false,                          // isDigest
                    false,                          // dontDeliver
                    mode,                           // deliver mode
                    null,                           // msgToReplace
                    bodyMimeType,                   // body type
                    body,                           // body
                  ];

                  // Tail args shared by both (parentWindow..compType)
                  const tailArgs = [
                    null,                           // parent window
                    null,                           // progress
                    listener,                       // listener
                    "",                             // password
                    originalMsgURI || "",           // original msg URI
                    compType,                       // compose type
                  ];

                  // Try modern 16-arg signature first (TB 128+).
                  // On TB 102-127, XPCOM throws NS_ERROR_XPC_NOT_ENOUGH_ARGS
                  // (0x80570001), so we fall back to legacy 18-arg with null
                  // attachment params (attachments already on composeFields).
                  // Modern TB may return a Promise -- catch async rejections.
                  let sendResult;
                  try {
                    sendResult = msgSend.createAndSendMessage(...commonArgs, ...tailArgs);
                  } catch (e) {
                    const isArgError = (e && e.result === 0x80570001) ||
                      String(e).includes("Not enough arguments");
                    if (isArgError) {
                      sendResult = msgSend.createAndSendMessage(...commonArgs, null, null, ...tailArgs);
                    } else {
                      throw e;
                    }
                  }
                  // Modern TB (128+) returns a Promise from createAndSendMessage.
                  // Handle both fulfillment and rejection -- belt-and-suspenders
                  // with the listener (settle is idempotent). For SaveAsDraft on
                  // older TB without the copy listener, the Promise fulfillment
                  // can be the only completion signal we get.
                  if (sendResult && typeof sendResult.then === "function") {
                    sendResult.then(
                      () => {
                        timer.cancel();
                        settle({ success: true });
                      },
                      e => {
                        timer.cancel();
                        settle({ error: e.toString() });
                      }
                    );
                  }
                } catch (e) {
                  timer.cancel();
                  settle({ error: e.toString() });
                }
              });
            }

            function escapeHtml(s) {
              return s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;').replace(/'/g, '&#39;');
            }

            /**
             * Converts body text to HTML for compose fields.
             * Handles both HTML input (entity-encodes non-ASCII) and plain text.
             */
            function formatBodyHtml(body, isHtml) {
              if (isHtml) {
                let text = (body || "").replace(/\n/g, '');
                text = [...text].map(c => c.codePointAt(0) > 127 ? `&#${c.codePointAt(0)};` : c).join('');
                return text;
              }
              return escapeHtml(body || "").replace(/\n/g, '<br>');
            }

            /**
             * Decides whether a compose operation will (or should) run in HTML
             * mode, and returns the matching msgComposeParams.format value.
             *
             * The caller's explicit isHtml wins (true/false). When isHtml is
             * omitted, the identity's compose-format preference is consulted.
             *
             * ForwardInline is a special case: Thunderbird's compose service
             * only passes format through when it's Default or OppositeOfDefault
             * for that compType, so when the caller's explicit isHtml conflicts
             * with the identity pref on a forward, we have to ask for
             * OppositeOfDefault instead of HTML/PlainText.
             *
             * Identity must be resolved (setComposeIdentity) before calling this.
             */
            function resolveComposeFormat(identity, isHtml, compType) {
              const identityUsesHtml = identity?.composeHtml !== false;
              const useHtml = isHtml === true || (isHtml !== false && identityUsesHtml);
              const isForward = compType === Ci.nsIMsgCompType.ForwardInline
                             || compType === Ci.nsIMsgCompType.ForwardAsAttachment;

              let format;
              if (isHtml === undefined) {
                format = Ci.nsIMsgCompFormat.Default;
              } else if (isForward) {
                const explicitMatchesPref = (isHtml === true) === identityUsesHtml;
                format = explicitMatchesPref
                  ? Ci.nsIMsgCompFormat.Default
                  : Ci.nsIMsgCompFormat.OppositeOfDefault;
              } else {
                format = isHtml === true ? Ci.nsIMsgCompFormat.HTML : Ci.nsIMsgCompFormat.PlainText;
              }
              return { useHtml, format };
            }

            /**
             * Sets compose identity from `from` param or falls back to default.
             * Returns "" on success, or { error } if `from` was explicitly
             * provided but not found / restricted.  Fallback to default only
             * applies when `from` is omitted.
             */
            function setComposeIdentity(msgComposeParams, from, fallbackServer) {
              if (from) {
                // Explicit `from` -- must resolve or fail, never silently substitute
                const identity = findIdentity(from);
                if (identity) {
                  // findIdentity searches accessible accounts, so this is safe
                  msgComposeParams.identity = identity;
                  return "";
                }
                // Not found in accessible accounts -- check ALL accounts to
                // distinguish "restricted" from "genuinely unknown"
                if (findIdentityIn(MailServices.accounts.accounts, from)) {
                  return { error: `identity ${from} belongs to a restricted account` };
                }
                return { error: `unknown identity: ${from} -- no matching account configured in Thunderbird` };
              }
              // No explicit `from` -- fall back to contextual default
              if (fallbackServer) {
                const account = MailServices.accounts.findAccountForServer(fallbackServer);
                if (account && isAccountAllowed(account.key)) {
                  msgComposeParams.identity = account.defaultIdentity;
                }
              } else {
                const defaultAccount = MailServices.accounts.defaultAccount;
                if (defaultAccount && isAccountAllowed(defaultAccount.key)) {
                  msgComposeParams.identity = defaultAccount.defaultIdentity;
                }
              }
              // If no identity was set (all fallbacks restricted), explicitly set
              // the first accessible identity. Without this, Thunderbird's
              // OpenComposeWindowWithParams fills identity from defaultAccount
              // internally, bypassing account restrictions.
              if (!msgComposeParams.identity) {
                for (const account of getAccessibleAccounts()) {
                  if (account.defaultIdentity) {
                    msgComposeParams.identity = account.defaultIdentity;
                    break;
                  }
                }
                if (!msgComposeParams.identity) {
                  return { error: "No accessible identity found -- all accounts are restricted" };
                }
              }
              return "";
            }

            function dryRunCompose(to, subject, body, cc, bcc, isHtml, from, attachments) {
              const result = {
                wouldSucceed: true,
                blockers: [],
                resolvedIdentity: null,
                isHtml: !!isHtml,
                subjectAfterSanitization: sanitizeHeaderLine(subject || ""),
                bodyLength: typeof body === "string" ? body.length : 0,
                recipients: {
                  to: countRecipients(to),
                  cc: countRecipients(cc),
                  bcc: countRecipients(bcc),
                },
                attachments: [],
                skipReviewBlocked: isSkipReviewBlocked(),
              };

              // Resolve identity through the same accessible-account path used
              // by composeMail, but never fall back silently to the default --
              // if `from` is set and doesn't match, surface the same error
              // sendMail would return.
              try {
                if (from) {
                  const identity = findIdentity(from);
                  if (!identity) {
                    result.blockers.push(`from identity not found or not accessible: ${from}`);
                    result.wouldSucceed = false;
                  } else {
                    result.resolvedIdentity = {
                      key: identity.key,
                      email: identity.email,
                      fullName: identity.fullName || null,
                    };
                  }
                } else {
                  // Default identity preview: first accessible identity.
                  const accounts = getAccessibleAccounts();
                  const firstIdentity = accounts[0] && accounts[0].defaultIdentity;
                  if (firstIdentity) {
                    result.resolvedIdentity = {
                      key: firstIdentity.key,
                      email: firstIdentity.email,
                      fullName: firstIdentity.fullName || null,
                      isDefault: true,
                    };
                  }
                }
              } catch (e) {
                result.blockers.push(`identity resolution failed: ${e.message || e}`);
                result.wouldSucceed = false;
              }

              // Per-attachment evaluation. Mirrors filePathsToAttachDescs but
              // never copies / decodes / writes anything; only reports verdict.
              if (Array.isArray(attachments)) {
                for (const entry of attachments) {
                  const att = { kind: null, name: null, status: "ok", reason: null, size: null };
                  if (typeof entry === "string") {
                    att.kind = "path";
                    att.name = entry;
                    if (isSensitiveFilePath(entry)) {
                      att.status = "blocked";
                      att.reason = "sensitive path";
                    } else {
                      try {
                        const file = createLocalFile(entry);
                        if (!file.exists()) {
                          att.status = "missing";
                          att.reason = "file does not exist";
                        } else {
                          try { att.size = file.fileSize; } catch { att.size = null; }
                          if (att.size !== null && att.size > MAX_FILE_PATH_ATTACHMENT_BYTES) {
                            att.status = "blocked";
                            att.reason = `exceeds ${MAX_FILE_PATH_ATTACHMENT_BYTES / 1024 / 1024}MB cap`;
                          }
                        }
                      } catch (e) {
                        att.status = "error";
                        att.reason = e.message || String(e);
                      }
                    }
                  } else if (entry && typeof entry === "object" && (entry.base64 || entry.content) && entry.name) {
                    att.kind = "inline";
                    att.name = entry.name;
                    const b64 = entry.base64 || entry.content;
                    att.size = typeof b64 === "string" ? Math.floor((b64.length * 3) / 4) : null;
                    if (typeof b64 === "string" && b64.length > MAX_BASE64_SIZE) {
                      att.status = "blocked";
                      att.reason = `inline base64 exceeds ${MAX_BASE64_SIZE / 1024 / 1024}MB`;
                    }
                  } else {
                    att.kind = "invalid";
                    att.name = typeof entry === "object" ? JSON.stringify(entry).slice(0, 80) : String(entry);
                    att.status = "blocked";
                    att.reason = "neither a file path nor an inline {name, base64} object";
                  }
                  if (att.status !== "ok") {
                    result.wouldSucceed = false;
                  }
                  result.attachments.push(att);
                }
              }

              return result;
            }

            /**
             * Composes a new email. Opens a compose window for review, or sends
             * directly when skipReview is true.
             *
             * HTML body handling quirks:
             * 1. Strip newlines from HTML - Thunderbird adds <br> for each \n
             * 2. Encode non-ASCII as HTML entities - compose window has charset issues
             *    with emojis/unicode even with <meta charset="UTF-8">
             */
            function composeMail(to, subject, body, cc, bcc, isHtml, from, attachments, skipReview, idempotencyKey) {
              try {
                if (skipReview && isSkipReviewBlocked()) {
                  return { error: "User preference blocks skipReview. Retry with skipReview: false (or omitted) to open the review window instead." };
                }
                // Idempotency: if a prior successful sendMail with this key
                // ran in the last 24h, return its result instead of sending
                // again. The audit-log entry is the source of truth.
                if (typeof idempotencyKey === "string" && idempotencyKey) {
                  const prior = findIdempotentEntry("sendMail", idempotencyKey);
                  if (prior) {
                    return { ...prior, idempotent: true, idempotencyKey };
                  }
                }
                appendComposeAudit({
                  tool: "sendMail",
                  skipReview: !!skipReview,
                  isHtml: !!isHtml,
                  from: typeof from === "string" ? from : null,
                  to: countRecipients(to),
                  cc: countRecipients(cc),
                  bcc: countRecipients(bcc),
                  subject: typeof subject === "string" ? subject.slice(0, 200) : null,
                  attachmentCount: Array.isArray(attachments) ? attachments.length : 0,
                  idempotencyKey: typeof idempotencyKey === "string" ? idempotencyKey.slice(0, 256) : null,
                });
                const msgComposeParams = Cc["@mozilla.org/messengercompose/composeparams;1"]
                  .createInstance(Ci.nsIMsgComposeParams);

                const composeFields = Cc["@mozilla.org/messengercompose/composefields;1"]
                  .createInstance(Ci.nsIMsgCompFields);

                composeFields.to = to || "";
                composeFields.cc = cc || "";
                composeFields.bcc = bcc || "";
                composeFields.subject = sanitizeHeaderLine(subject || "");

                msgComposeParams.type = Ci.nsIMsgCompType.New;
                msgComposeParams.composeFields = composeFields;

                const identityResult = setComposeIdentity(msgComposeParams, from, null);
                if (identityResult && identityResult.error) return identityResult;

                // Match body shape and format to caller intent / identity pref.
                // When the resolved mode is plain, ship a plain body -- the HTML
                // envelope would otherwise render as literal text in plain-mode
                // editors and recipients.
                const { useHtml, format } = resolveComposeFormat(msgComposeParams.identity, isHtml, Ci.nsIMsgCompType.New);
                msgComposeParams.format = format;
                if (useHtml) {
                  const formatted = formatBodyHtml(body, isHtml);
                  composeFields.body = isHtml && formatted.includes('<html')
                    ? formatted
                    : `<html><head><meta charset="UTF-8"></head><body>${formatted}</body></html>`;
                } else {
                  composeFields.body = body || "";
                }

                const { descs: fileDescs, failed: failedPaths } = filePathsToAttachDescs(attachments);

                if (skipReview) {
                  return sendMessageDirectly(composeFields, msgComposeParams.identity, fileDescs, null, Ci.nsIMsgCompType.New, Ci.nsIMsgCompDeliverMode.Now, useHtml ? "text/html" : "text/plain").then(result => {
                    if (result.success) {
                      let msg = "Message sent";
                      if (failedPaths.length > 0) msg += ` (failed to attach: ${failedPaths.join(", ")})`;
                      result.message = msg;
                      // Idempotency: record the successful outcome so a retry
                      // with the same key returns this result instead of
                      // sending again. The pre-send audit entry only records
                      // intent; findIdempotentEntry filters to success:true.
                      if (typeof idempotencyKey === "string" && idempotencyKey) {
                        appendComposeAudit({
                          tool: "sendMail",
                          success: true,
                          idempotencyKey: idempotencyKey.slice(0, 256),
                          result,
                        });
                      }
                    }
                    return result;
                  });
                }

                const msgComposeService = Cc["@mozilla.org/messengercompose;1"]
                  .getService(Ci.nsIMsgComposeService);
                msgComposeService.OpenComposeWindowWithParams(null, msgComposeParams);

                injectAttachmentsAsync(fileDescs);

                let msg = "Compose window opened";
                if (failedPaths.length > 0) msg += ` (failed to attach: ${failedPaths.join(", ")})`;
                return { success: true, message: msg };
              } catch (e) {
                return { error: e.toString() };
              }
            }

            /**
             * Saves a composed message to the identity's Drafts folder without
             * sending or opening a compose window. The destination folder is
             * resolved by Thunderbird from the identity's draft-folder pref.
             */
            function saveDraft(to, subject, body, cc, bcc, isHtml, from, attachments) {
              try {
                appendComposeAudit({
                  tool: "saveDraft",
                  isHtml: !!isHtml,
                  from: typeof from === "string" ? from : null,
                  to: countRecipients(to),
                  cc: countRecipients(cc),
                  bcc: countRecipients(bcc),
                  subject: typeof subject === "string" ? subject.slice(0, 200) : null,
                  attachmentCount: Array.isArray(attachments) ? attachments.length : 0,
                });
                const msgComposeParams = Cc["@mozilla.org/messengercompose/composeparams;1"]
                  .createInstance(Ci.nsIMsgComposeParams);

                const composeFields = Cc["@mozilla.org/messengercompose/composefields;1"]
                  .createInstance(Ci.nsIMsgCompFields);

                composeFields.to = to || "";
                composeFields.cc = cc || "";
                composeFields.bcc = bcc || "";
                composeFields.subject = sanitizeHeaderLine(subject || "");

                msgComposeParams.type = Ci.nsIMsgCompType.New;
                msgComposeParams.composeFields = composeFields;

                const identityResult = setComposeIdentity(msgComposeParams, from, null);
                if (identityResult && identityResult.error) return identityResult;

                const { useHtml, format } = resolveComposeFormat(msgComposeParams.identity, isHtml, Ci.nsIMsgCompType.New);
                msgComposeParams.format = format;
                if (useHtml) {
                  const formatted = formatBodyHtml(body, isHtml);
                  composeFields.body = isHtml && formatted.includes('<html')
                    ? formatted
                    : `<html><head><meta charset="UTF-8"></head><body>${formatted}</body></html>`;
                } else {
                  composeFields.body = body || "";
                }

                const { descs: fileDescs, failed: failedPaths } = filePathsToAttachDescs(attachments);

                return sendMessageDirectly(
                  composeFields,
                  msgComposeParams.identity,
                  fileDescs,
                  null,
                  Ci.nsIMsgCompType.New,
                  Ci.nsIMsgCompDeliverMode.SaveAsDraft,
                  useHtml ? "text/html" : "text/plain"
                ).then(result => {
                  if (result.success) {
                    let msg = "Draft saved";
                    if (failedPaths.length > 0) msg += ` (failed to attach: ${failedPaths.join(", ")})`;
                    result.message = msg;
                  }
                  return result;
                });
              } catch (e) {
                return { error: e.toString() };
              }
            }

            /**
             * Replies to a message with quoted original. Opens a compose window
             * for review, or sends directly when skipReview is true.
             *
             * Review path uses Thunderbird's native reply compose flow so it can
             * build the quoted original, place the identity signature according
             * to user preferences, and set threading headers/disposition flags.
             * skipReview still uses direct send, so it keeps a manual quoted body
             * and manually marks the original as replied after a successful send.
             */
	            function replyToMessage(messageId, folderPath, body, replyAll, isHtml, to, cc, bcc, from, attachments, skipReview, idempotencyKey) {
	              return new Promise((resolve) => {
	                try {
	                  if (skipReview && isSkipReviewBlocked()) {
	                    resolve({ error: "User preference blocks skipReview. Retry with skipReview: false (or omitted) to open the review window instead." });
	                    return;
	                  }
	                  if (typeof idempotencyKey === "string" && idempotencyKey) {
	                    const prior = findIdempotentEntry("replyToMessage", idempotencyKey);
	                    if (prior) {
	                      resolve({ ...prior, idempotent: true, idempotencyKey });
	                      return;
	                    }
	                  }
	                  appendComposeAudit({
	                    tool: "replyToMessage",
	                    skipReview: !!skipReview,
	                    replyAll: !!replyAll,
	                    isHtml: !!isHtml,
	                    from: typeof from === "string" ? from : null,
	                    originalMessageId: typeof messageId === "string" ? messageId.slice(0, 256) : null,
	                    to: countRecipients(to),
	                    cc: countRecipients(cc),
	                    bcc: countRecipients(bcc),
	                    attachmentCount: Array.isArray(attachments) ? attachments.length : 0,
	                    idempotencyKey: typeof idempotencyKey === "string" ? idempotencyKey.slice(0, 256) : null,
	                  });
	                  const found = findMessage(messageId, folderPath);
	                  if (found.error) {
	                    resolve({ error: found.error });
	                    return;
	                  }
	                  const { msgHdr, folder } = found;
	                  const { descs: fileDescs, failed: failedPaths } = filePathsToAttachDescs(attachments);
	                  const msgURI = folder.getUriForMsg(msgHdr);
	                  const compType = replyAll ? Ci.nsIMsgCompType.ReplyAll : Ci.nsIMsgCompType.Reply;

	                  const msgComposeParams = Cc["@mozilla.org/messengercompose/composeparams;1"]
	                    .createInstance(Ci.nsIMsgComposeParams);

	                  const composeFields = Cc["@mozilla.org/messengercompose/composefields;1"]
	                    .createInstance(Ci.nsIMsgCompFields);

	                  msgComposeParams.type = compType;
	                  msgComposeParams.originalMsgURI = msgURI;
	                  msgComposeParams.composeFields = composeFields;

	                  try {
	                    msgComposeParams.origMsgHdr = msgHdr;
	                  } catch {}

	                  const identityResult = setComposeIdentity(msgComposeParams, from, folder.server);
	                  if (identityResult && identityResult.error) {
	                    resolve(identityResult);
	                    return;
	                  }

	                  // Resolve compose mode against caller intent + identity pref.
	                  // The skipReview branch reads useHtml below to shape the body.
	                  const { useHtml: replyUseHtml, format: replyFormat } =
	                    resolveComposeFormat(msgComposeParams.identity, isHtml, compType);
	                  msgComposeParams.format = replyFormat;

	                  // Pass through only the fields the caller explicitly provided.
	                  // Any field left undefined is filled in by Thunderbird's native
	                  // reply/reply-all machinery (including proper Reply-To,
	                  // Mail-Followup-To, mailing-list handling, and self-filtering
	                  // against the selected identity). Our old custom
	                  // getReplyAllCcRecipients path bypassed all of that.
	                  const reviewTo = to;
	                  const reviewCc = cc;

	                  if (skipReview) {
	                    const { MsgHdrToMimeMessage } = ChromeUtils.importESModule(
	                      "resource:///modules/gloda/MimeMessage.sys.mjs"
                      );

	                    MsgHdrToMimeMessage(msgHdr, null, (aMsgHdr, aMimeMsg) => {
	                      try {
	                        const originalBody = extractPlainTextBody(aMimeMsg);

	                        if (replyAll) {
	                          composeFields.to = to || msgHdr.author;
	                          if (cc) {
	                            composeFields.cc = cc;
	                          } else {
	                            const replyAllCc = getReplyAllCcRecipients(msgHdr, folder);
	                            if (replyAllCc) composeFields.cc = replyAllCc;
	                          }
	                        } else {
	                          composeFields.to = to || msgHdr.author;
	                          if (cc) composeFields.cc = cc;
	                        }

	                        composeFields.bcc = bcc || "";

	                        const origSubject = sanitizeHeaderLine(msgHdr.mime2DecodedSubject || msgHdr.subject || "");
	                        composeFields.subject = /^re:/i.test(origSubject) ? origSubject : `Re: ${origSubject}`;
	                        composeFields.references = `<${messageId}>`;
	                        composeFields.setHeader("In-Reply-To", `<${messageId}>`);

	                        const dateStr = msgHdr.date ? new Date(msgHdr.date / 1000).toLocaleString() : "";
	                        const author = msgHdr.mime2DecodedAuthor || msgHdr.author || "";

	                        // Direct send goes through nsIMsgSend, not nsIMsgCompose, so
	                        // it still uses a hand-built quoted body and cannot place the
	                        // identity signature according to reply preferences. The shape
	                        // matches the resolved compose mode -- shipping an HTML envelope
	                        // for a plain-format send would otherwise render as literal
	                        // markup in the recipient's mail client.
	                        if (replyUseHtml) {
	                          const quotedLines = originalBody.split('\n').map(line =>
	                            `&gt; ${escapeHtml(line)}`
	                          ).join('<br>');
	                          const quotedHtml = escapeHtml(originalBody).replace(/\n/g, '<br>');
	                          const quoteBlock = isHtml
	                            ? `<br><br>On ${dateStr}, ${escapeHtml(author)} wrote:<blockquote type="cite">${quotedHtml}</blockquote>`
	                            : `<br><br>On ${dateStr}, ${escapeHtml(author)} wrote:<br>${quotedLines}`;
	                          composeFields.body = `<html><head><meta charset="UTF-8"></head><body>${formatBodyHtml(body, isHtml)}${quoteBlock}</body></html>`;
	                        } else {
	                          const quotedLines = originalBody.split('\n').map(line => `> ${line}`).join('\n');
	                          composeFields.body = `${body || ""}\n\nOn ${dateStr}, ${author} wrote:\n${quotedLines}`;
	                        }

	                        sendMessageDirectly(composeFields, msgComposeParams.identity, fileDescs, msgURI, compType, Ci.nsIMsgCompDeliverMode.Now, replyUseHtml ? "text/html" : "text/plain").then(result => {
	                          if (result.success) {
	                            let repliedDisposition = null;
	                            try {
	                              repliedDisposition = Ci.nsIMsgFolder.nsMsgDispositionState_Replied;
	                            } catch {}
	                            markMessageDispositionState(msgHdr, repliedDisposition);

	                            let msg = "Reply sent";
	                            if (failedPaths.length > 0) msg += ` (failed to attach: ${failedPaths.join(", ")})`;
	                            result.message = msg;
	                            if (typeof idempotencyKey === "string" && idempotencyKey) {
	                              appendComposeAudit({
	                                tool: "replyToMessage",
	                                success: true,
	                                idempotencyKey: idempotencyKey.slice(0, 256),
	                                result,
	                              });
	                            }
	                          }
	                          resolve(result);
	                        });
	                      } catch (e) {
	                        resolve({ error: e.toString() });
	                      }
	                    }, true, { examineEncryptedParts: true });
	                    return;
	                  }

	                  openComposeWindowWithCustomizations(
	                    msgComposeParams,
	                    msgURI,
	                    compType,
	                    msgComposeParams.identity,
	                    body,
	                    isHtml,
	                    reviewTo,
	                    reviewCc,
	                    bcc,
	                    fileDescs
	                  ).then(result => {
	                    if (result.success) {
	                      let msg = "Reply window opened";
	                      if (failedPaths.length > 0) msg += ` (failed to attach: ${failedPaths.join(", ")})`;
	                      result.message = msg;
	                    }
	                    resolve(result);
	                  });

	                } catch (e) {
	                  resolve({ error: e.toString() });
	                }
	              });
            }

            /**
             * Forwards a message with original content and attachments.
             *
             * Review path uses Thunderbird's native ForwardInline compose flow so
             * TB builds the forward body with a proper <blockquote type="cite">
             * quote, auto-attaches the original message's attachments, places the
             * identity signature per user preferences, and sets the $Forwarded
             * disposition on the original after a successful send. The caller's
             * intro body is injected via NotifyComposeBodyReady, mirroring how
             * replyToMessage handles intro injection.
             *
             * skipReview still uses direct send, so it keeps a manual forward
             * block + auto-attaches originals from MsgHdrToMimeMessage + manually
             * marks the original as forwarded after a successful send.
             */
            function forwardMessage(messageId, folderPath, to, body, isHtml, cc, bcc, from, attachments, skipReview, idempotencyKey) {
              return new Promise((resolve) => {
                try {
                  if (skipReview && isSkipReviewBlocked()) {
                    resolve({ error: "User preference blocks skipReview. Retry with skipReview: false (or omitted) to open the review window instead." });
                    return;
                  }
                  if (typeof idempotencyKey === "string" && idempotencyKey) {
                    const prior = findIdempotentEntry("forwardMessage", idempotencyKey);
                    if (prior) {
                      resolve({ ...prior, idempotent: true, idempotencyKey });
                      return;
                    }
                  }
                  appendComposeAudit({
                    tool: "forwardMessage",
                    skipReview: !!skipReview,
                    isHtml: !!isHtml,
                    from: typeof from === "string" ? from : null,
                    originalMessageId: typeof messageId === "string" ? messageId.slice(0, 256) : null,
                    to: countRecipients(to),
                    cc: countRecipients(cc),
                    bcc: countRecipients(bcc),
                    attachmentCount: Array.isArray(attachments) ? attachments.length : 0,
                    idempotencyKey: typeof idempotencyKey === "string" ? idempotencyKey.slice(0, 256) : null,
                  });
                  const found = findMessage(messageId, folderPath);
                  if (found.error) {
                    resolve({ error: found.error });
                    return;
                  }
                  const { msgHdr, folder } = found;
                  const { descs: fileDescs, failed: failedPaths } = filePathsToAttachDescs(attachments);
                  const msgURI = folder.getUriForMsg(msgHdr);
                  const compType = Ci.nsIMsgCompType.ForwardInline;

                  const msgComposeParams = Cc["@mozilla.org/messengercompose/composeparams;1"]
                    .createInstance(Ci.nsIMsgComposeParams);

                  const composeFields = Cc["@mozilla.org/messengercompose/composefields;1"]
                    .createInstance(Ci.nsIMsgCompFields);

                  msgComposeParams.type = compType;
                  msgComposeParams.originalMsgURI = msgURI;
                  msgComposeParams.composeFields = composeFields;

                  try {
                    msgComposeParams.origMsgHdr = msgHdr;
                  } catch {}

                  const identityResult = setComposeIdentity(msgComposeParams, from, folder.server);
                  if (identityResult && identityResult.error) {
                    resolve(identityResult);
                    return;
                  }

                  // ForwardInline only passes the format flag through when it is
                  // Default or OppositeOfDefault -- HTML/PlainText are ignored and
                  // the identity's compose pref always wins. resolveComposeFormat
                  // returns OppositeOfDefault when the caller's explicit isHtml
                  // conflicts with the identity pref so we can still force the
                  // intended editor mode.
                  const { useHtml: fwdUseHtml, format: fwdFormat } =
                    resolveComposeFormat(msgComposeParams.identity, isHtml, compType);
                  msgComposeParams.format = fwdFormat;

                  if (skipReview) {
                    const { MsgHdrToMimeMessage } = ChromeUtils.importESModule(
                      "resource:///modules/gloda/MimeMessage.sys.mjs"
                    );

                    MsgHdrToMimeMessage(msgHdr, null, (aMsgHdr, aMimeMsg) => {
                      try {
                        const originalBody = extractPlainTextBody(aMimeMsg);

                        composeFields.to = to;
                        composeFields.cc = cc || "";
                        composeFields.bcc = bcc || "";

                        const origSubject = sanitizeHeaderLine(msgHdr.mime2DecodedSubject || msgHdr.subject || "");
                        composeFields.subject = /^fwd:/i.test(origSubject) ? origSubject : `Fwd: ${origSubject}`;

                        const dateStr = msgHdr.date ? new Date(msgHdr.date / 1000).toLocaleString() : "";
                        const fwdAuthor = msgHdr.mime2DecodedAuthor || msgHdr.author || "";
                        const fwdRecipients = msgHdr.mime2DecodedRecipients || msgHdr.recipients || "";

                        // Direct send goes through nsIMsgSend, not nsIMsgCompose,
                        // so we hand-build the forward block. The shape matches the
                        // resolved compose mode -- shipping an HTML envelope for a
                        // plain-format send would render as literal markup in the
                        // recipient's mail client.
                        if (fwdUseHtml) {
                          const fwdHeaderHtml =
                            `-------- Forwarded Message --------<br>` +
                            `Subject: ${escapeHtml(origSubject)}<br>` +
                            `Date: ${dateStr}<br>` +
                            `From: ${escapeHtml(fwdAuthor)}<br>` +
                            `To: ${escapeHtml(fwdRecipients)}<br><br>`;
                          const quotedHtml = escapeHtml(originalBody).replace(/\n/g, '<br>');
                          const quotedLinesHtml = originalBody.split('\n').map(line =>
                            `&gt; ${escapeHtml(line)}`
                          ).join('<br>');
                          const forwardBlock = isHtml
                            ? `<blockquote type="cite">${fwdHeaderHtml}${quotedHtml}</blockquote>`
                            : `${fwdHeaderHtml}${quotedLinesHtml}`;
                          const introHtml = body ? formatBodyHtml(body, isHtml) + '<br><br>' : "";
                          composeFields.body = `<html><head><meta charset="UTF-8"></head><body>${introHtml}${forwardBlock}</body></html>`;
                        } else {
                          const fwdHeader =
                            `-------- Forwarded Message --------\n` +
                            `Subject: ${origSubject}\n` +
                            `Date: ${dateStr}\n` +
                            `From: ${fwdAuthor}\n` +
                            `To: ${fwdRecipients}\n\n`;
                          composeFields.body = `${body ? body + '\n\n' : ''}${fwdHeader}${originalBody}`;
                        }

                        const origDescs = [];
                        if (aMimeMsg && aMimeMsg.allUserAttachments) {
                          for (const att of aMimeMsg.allUserAttachments) {
                            try {
                              origDescs.push({ url: att.url, name: att.name, contentType: att.contentType });
                            } catch {
                              // Skip unreadable original attachments
                            }
                          }
                        }
                        const allDescs = [...origDescs, ...fileDescs];

                        sendMessageDirectly(composeFields, msgComposeParams.identity, allDescs, msgURI, compType, Ci.nsIMsgCompDeliverMode.Now, fwdUseHtml ? "text/html" : "text/plain").then(result => {
                          if (result.success) {
                            let forwardedDisposition = null;
                            try {
                              forwardedDisposition = Ci.nsIMsgFolder.nsMsgDispositionState_Forwarded;
                            } catch {}
                            markMessageDispositionState(msgHdr, forwardedDisposition);

                            let msg = `Forward sent with ${allDescs.length} attachment(s)`;
                            if (failedPaths.length > 0) msg += ` (failed to attach: ${failedPaths.join(", ")})`;
                            result.message = msg;
                            if (typeof idempotencyKey === "string" && idempotencyKey) {
                              appendComposeAudit({
                                tool: "forwardMessage",
                                success: true,
                                idempotencyKey: idempotencyKey.slice(0, 256),
                                result,
                              });
                            }
                          }
                          resolve(result);
                        });
                      } catch (e) {
                        resolve({ error: e.toString() });
                      }
                    }, true, { examineEncryptedParts: true });
                    return;
                  }

                  // Review path: TB builds the forward body and auto-attaches the
                  // original's attachments via ForwardInline. The intro body and
                  // user-specified extra attachments are injected once the compose
                  // window's editor signals NotifyComposeBodyReady. Subject is set
                  // by TB from origMsgHdr.
                  openComposeWindowWithCustomizations(
                    msgComposeParams,
                    msgURI,
                    compType,
                    msgComposeParams.identity,
                    body,
                    isHtml,
                    to,
                    cc,
                    bcc,
                    fileDescs
                  ).then(result => {
                    if (result.success) {
                      let msg = "Forward window opened";
                      if (failedPaths.length > 0) msg += ` (failed to attach: ${failedPaths.join(", ")})`;
                      result.message = msg;
                    }
                    resolve(result);
                  });
                } catch (e) {
                  resolve({ error: e.toString() });
                }
              });
            }

  Object.assign(ctx, {
    templatesDir,
    readTextFile,
    parseFrontmatter,
    loadTemplate,
    listTemplates,
    renderTemplate,
    findIdentityIn,
    findIdentity,
    createLocalFile,
    filePathsToAttachDescs,
    descsToMsgAttachments,
    addAttachmentsToComposeWindow,
    injectAttachmentsAsync,
    splitAddressHeader,
    extractAddressEmail,
    mergeAddressHeaders,
    getReplyAllCcRecipients,
    getIdentityAutoRecipientHeader,
    applyComposeRecipientOverrides,
    formatBodyFragmentHtml,
    moveComposeSelectionToBodyStartIfRange,
    insertReplyBodyIntoComposeWindow,
    openComposeWindowWithCustomizations,
    markMessageDispositionState,
    sendMessageDirectly,
    escapeHtml,
    formatBodyHtml,
    resolveComposeFormat,
    setComposeIdentity,
    dryRunCompose,
    composeMail,
    saveDraft,
    replyToMessage,
    forwardMessage,
  });
};
