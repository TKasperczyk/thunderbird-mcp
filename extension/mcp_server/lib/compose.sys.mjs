/**
 * Mail composition: sendMail / replyToMessage / forwardMessage and the
 * tangle of attachment, address, and Mozilla-window plumbing they share.
 *
 * The deepest closure of the original api.js monolith. Three things you
 * need to know before touching this module:
 *
 *  1. Attachment-injection nsITimer instances are held in _attachTimers
 *     (state.sys.mjs) so they aren't GC'd before they fire. The fire-and-
 *     forget timer in injectAttachmentsAsync MUST stay registered until
 *     the notify callback runs.
 *  2. _claimedReplyComposeWindows (state.sys.mjs) is a WeakSet on purpose.
 *     Concurrent replyToMessage calls would otherwise both bind observers
 *     to the same compose window. Do not change to a regular Set.
 *  3. sendMessageDirectly carries a 16-arg / 18-arg fallback for the TB
 *     128+ vs 102-127 nsIMsgSend.createAndSendMessage signature change.
 *     The fallback path is essential for old Thunderbird builds.
 *
 * Helpers escapeHtml / splitAddressHeader / extractAddressEmail /
 * mergeAddressHeaders are intentionally duplicated with messages.sys.mjs --
 * a future encoding.sys.mjs commit will collapse them.
 */

export function makeCompose({
  Services, Cc, Ci, MailServices,
  // From access-control.sys.mjs
  isFolderAccessible, isSkipReviewBlocked,
  // From accounts.sys.mjs
  findIdentity, findIdentityIn, getIdentityAutoRecipientHeader,
  // From folders.sys.mjs
  openFolder,
  // From messages.sys.mjs
  findMessage, markMessageDispositionState,
  // From state.sys.mjs (mutable singletons -- must be the SAME instances
  // that api.js and onShutdown both hold).
  _attachTimers, _claimedReplyComposeWindows, _tempAttachFiles,
  // From constants.sys.mjs
  MAX_BASE64_SIZE, COMPOSE_WINDOW_LOAD_DELAY_MS,
}) {
  // Per-instance counter so each makeCompose call (in practice one per
  // start()) gets a fresh sequence for temp-attachment filenames.
  let _tempFileCounter = 0;

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
          // File path attachment
          const file = createLocalFile(entry);
          if (file.exists()) {
            descs.push({ url: Services.io.newFileURI(file).spec, name: file.leafName, size: file.fileSize });
          } else {
            failed.push(entry);
          }
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
      } catch {}
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

  function applyComposeRecipientOverrides(composeWin, identity, to, cc, bcc) {
    if (!composeWin) return;
    const overrides = { identityKey: null };
    if (to) overrides.to = to;
    if (cc) overrides.cc = mergeAddressHeaders(getIdentityAutoRecipientHeader(identity, "cc"), cc);
    if (bcc) overrides.bcc = mergeAddressHeaders(getIdentityAutoRecipientHeader(identity, "bcc"), bcc);
    if (Object.keys(overrides).length === 1) return;

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

  function insertReplyBodyIntoComposeWindow(composeWin, body, isHtml) {
    if (!composeWin || !body) return;
    const fragment = formatBodyFragmentHtml(body, isHtml);
    if (!fragment) return;

    const browser = typeof composeWin.getBrowser === "function" ? composeWin.getBrowser() : null;
    const editorDoc = browser?.contentDocument;
    if (editorDoc && typeof editorDoc.execCommand === "function") {
      editorDoc.execCommand("insertHTML", false, fragment);
    } else {
      const editor = typeof composeWin.GetCurrentEditor === "function" ? composeWin.GetCurrentEditor() : null;
      if (editor && typeof editor.insertHTML === "function") {
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

  function openReplyComposeWindowWithCustomizations(msgComposeParams, originalMsgURI, compType, identity, body, isHtml, to, cc, bcc, attachDescs) {
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
          finish({ error: "Timed out waiting for reply compose window" });
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
          if (_claimedReplyComposeWindows.has(composeWin)) return;
          _claimedReplyComposeWindows.add(composeWin);

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
  function sendMessageDirectly(composeFields, identity, attachDescs, originalMsgURI, compType) {
    if (!identity) {
      return Promise.resolve({ error: "No identity available for direct send" });
    }

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

        const listener = {
          QueryInterface: ChromeUtils.generateQI(["nsIMsgSendListener"]),
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
        };

        // Common args shared by both signatures (positions 1-10)
        const commonArgs = [
          null,                           // editor
          identity,                       // identity
          accountKey,                     // account key
          composeFields,                  // fields
          false,                          // isDigest
          false,                          // dontDeliver
          Ci.nsIMsgCompDeliverMode.Now,   // deliver mode
          null,                           // msgToReplace
          "text/html",                    // body type
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
        // Handle async Promise from modern TB (128+)
        if (sendResult && typeof sendResult.catch === "function") {
          sendResult.catch(e => {
            timer.cancel();
            settle({ error: e.toString() });
          });
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

  /**
   * Composes a new email. Opens a compose window for review, or sends
   * directly when skipReview is true.
   *
   * HTML body handling quirks:
   * 1. Strip newlines from HTML - Thunderbird adds <br> for each \n
   * 2. Encode non-ASCII as HTML entities - compose window has charset issues
   *    with emojis/unicode even with <meta charset="UTF-8">
   */
  function composeMail(to, subject, body, cc, bcc, isHtml, from, attachments, skipReview) {
    try {
      if (skipReview && isSkipReviewBlocked()) {
        return { error: "User preference blocks skipReview. Retry with skipReview: false (or omitted) to open the review window instead." };
      }
      const msgComposeParams = Cc["@mozilla.org/messengercompose/composeparams;1"]
        .createInstance(Ci.nsIMsgComposeParams);

      const composeFields = Cc["@mozilla.org/messengercompose/composefields;1"]
        .createInstance(Ci.nsIMsgCompFields);

      composeFields.to = to || "";
      composeFields.cc = cc || "";
      composeFields.bcc = bcc || "";
      composeFields.subject = subject || "";

      const formatted = formatBodyHtml(body, isHtml);
      if (isHtml && formatted.includes('<html')) {
        composeFields.body = formatted;
      } else {
        composeFields.body = `<html><head><meta charset="UTF-8"></head><body>${formatted}</body></html>`;
      }

      msgComposeParams.type = Ci.nsIMsgCompType.New;
      msgComposeParams.format = Ci.nsIMsgCompFormat.HTML;
      msgComposeParams.composeFields = composeFields;

      const identityResult = setComposeIdentity(msgComposeParams, from, null);
      if (identityResult && identityResult.error) return identityResult;

      const { descs: fileDescs, failed: failedPaths } = filePathsToAttachDescs(attachments);

      if (skipReview) {
        return sendMessageDirectly(composeFields, msgComposeParams.identity, fileDescs, null, Ci.nsIMsgCompType.New).then(result => {
          if (result.success) {
            let msg = "Message sent";
            if (failedPaths.length > 0) msg += ` (failed to attach: ${failedPaths.join(", ")})`;
            result.message = msg;
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
   * Replies to a message with quoted original. Opens a compose window
   * for review, or sends directly when skipReview is true.
   *
   * Review path uses Thunderbird's native reply compose flow so it can
   * build the quoted original, place the identity signature according
   * to user preferences, and set threading headers/disposition flags.
   * skipReview still uses direct send, so it keeps a manual quoted body
   * and manually marks the original as replied after a successful send.
   */
  function replyToMessage(messageId, folderPath, body, replyAll, isHtml, to, cc, bcc, from, attachments, skipReview) {
    return new Promise((resolve) => {
      try {
        if (skipReview && isSkipReviewBlocked()) {
          resolve({ error: "User preference blocks skipReview. Retry with skipReview: false (or omitted) to open the review window instead." });
          return;
        }
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
        msgComposeParams.format = Ci.nsIMsgCompFormat.HTML;
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

              const origSubject = msgHdr.mime2DecodedSubject || msgHdr.subject || "";
              composeFields.subject = /^re:/i.test(origSubject) ? origSubject : `Re: ${origSubject}`;
              composeFields.references = `<${messageId}>`;
              composeFields.setHeader("In-Reply-To", `<${messageId}>`);

              const dateStr = msgHdr.date ? new Date(msgHdr.date / 1000).toLocaleString() : "";
              const author = msgHdr.mime2DecodedAuthor || msgHdr.author || "";
              const quotedLines = originalBody.split('\n').map(line =>
                `&gt; ${escapeHtml(line)}`
              ).join('<br>');
              const quotedHtml = escapeHtml(originalBody).replace(/\n/g, '<br>');
              const quoteBlock = isHtml
                ? `<br><br>On ${dateStr}, ${escapeHtml(author)} wrote:<blockquote type="cite">${quotedHtml}</blockquote>`
                : `<br><br>On ${dateStr}, ${escapeHtml(author)} wrote:<br>${quotedLines}`;

              // Direct send goes through nsIMsgSend, not nsIMsgCompose, so
              // it still uses a hand-built quoted body and cannot place the
              // identity signature according to reply preferences.
              composeFields.body = `<html><head><meta charset="UTF-8"></head><body>${formatBodyHtml(body, isHtml)}${quoteBlock}</body></html>`;

              sendMessageDirectly(composeFields, msgComposeParams.identity, fileDescs, msgURI, compType).then(result => {
                if (result.success) {
                  let repliedDisposition = null;
                  try {
                    repliedDisposition = Ci.nsIMsgFolder.nsMsgDispositionState_Replied;
                  } catch {}
                  markMessageDispositionState(msgHdr, repliedDisposition);

                  let msg = "Reply sent";
                  if (failedPaths.length > 0) msg += ` (failed to attach: ${failedPaths.join(", ")})`;
                  result.message = msg;
                }
                resolve(result);
              });
            } catch (e) {
              resolve({ error: e.toString() });
            }
          }, true, { examineEncryptedParts: true });
          return;
        }

        openReplyComposeWindowWithCustomizations(
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
   * Forwards a message with original content and attachments. Opens a
   * compose window for review, or sends directly when skipReview is true.
   * Uses New type with manual forward quote to preserve both intro body and forwarded content.
   */
  function forwardMessage(messageId, folderPath, to, body, isHtml, cc, bcc, from, attachments, skipReview) {
    return new Promise((resolve) => {
      try {
        if (skipReview && isSkipReviewBlocked()) {
          resolve({ error: "User preference blocks skipReview. Retry with skipReview: false (or omitted) to open the review window instead." });
          return;
        }
        const found = findMessage(messageId, folderPath);
        if (found.error) {
          resolve({ error: found.error });
          return;
        }
        const { msgHdr, folder } = found;

        // Get attachments and body from original message
        const { MsgHdrToMimeMessage } = ChromeUtils.importESModule(
          "resource:///modules/gloda/MimeMessage.sys.mjs"
        );

        MsgHdrToMimeMessage(msgHdr, null, (aMsgHdr, aMimeMsg) => {
          try {
            const msgComposeParams = Cc["@mozilla.org/messengercompose/composeparams;1"]
              .createInstance(Ci.nsIMsgComposeParams);

            const composeFields = Cc["@mozilla.org/messengercompose/composefields;1"]
              .createInstance(Ci.nsIMsgCompFields);

            composeFields.to = to;
            composeFields.cc = cc || "";
            composeFields.bcc = bcc || "";

            const origSubject = msgHdr.mime2DecodedSubject || msgHdr.subject || "";
            composeFields.subject = /^fwd:/i.test(origSubject) ? origSubject : `Fwd: ${origSubject}`;

            // Get original body
            const originalBody = extractPlainTextBody(aMimeMsg);

            // Build forward header block
            const dateStr = msgHdr.date ? new Date(msgHdr.date / 1000).toLocaleString() : "";
            const fwdAuthor = msgHdr.mime2DecodedAuthor || msgHdr.author || "";
            const fwdSubject = msgHdr.mime2DecodedSubject || msgHdr.subject || "";
            const fwdRecipients = msgHdr.mime2DecodedRecipients || msgHdr.recipients || "";
            const escapedBody = escapeHtml(originalBody).replace(/\n/g, '<br>');

            const forwardBlock = `-------- Forwarded Message --------<br>` +
              `Subject: ${escapeHtml(fwdSubject)}<br>` +
              `Date: ${dateStr}<br>` +
              `From: ${escapeHtml(fwdAuthor)}<br>` +
              `To: ${escapeHtml(fwdRecipients)}<br><br>` +
              escapedBody;

            // Combine intro body + forward block
            const introHtml = body ? formatBodyHtml(body, isHtml) + '<br><br>' : "";

            composeFields.body = `<html><head><meta charset="UTF-8"></head><body>${introHtml}${forwardBlock}</body></html>`;

            // Collect original message attachments as descriptors
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

            // Validate user-specified file attachments
            const { descs: fileDescs, failed: failedPaths } = filePathsToAttachDescs(attachments);

            // Use New type - we build forward quote manually
            msgComposeParams.type = Ci.nsIMsgCompType.New;
            msgComposeParams.format = Ci.nsIMsgCompFormat.HTML;
            msgComposeParams.composeFields = composeFields;

            const identityResult = setComposeIdentity(msgComposeParams, from, folder.server);
            if (identityResult && identityResult.error) { resolve(identityResult); return; }

            const allDescs = [...origDescs, ...fileDescs];

            if (skipReview) {
              const msgURI = folder.getUriForMsg(msgHdr);
              sendMessageDirectly(composeFields, msgComposeParams.identity, allDescs, msgURI, Ci.nsIMsgCompType.ForwardInline).then(result => {
                if (result.success) {
                  let msg = `Forward sent with ${allDescs.length} attachment(s)`;
                  if (failedPaths.length > 0) msg += ` (failed to attach: ${failedPaths.join(", ")})`;
                  result.message = msg;
                }
                resolve(result);
              });
              return;
            }

            const msgComposeService = Cc["@mozilla.org/messengercompose;1"]
              .getService(Ci.nsIMsgComposeService);
            msgComposeService.OpenComposeWindowWithParams(null, msgComposeParams);

            injectAttachmentsAsync(allDescs);

            let msg = `Forward window opened with ${allDescs.length} attachment(s)`;
            if (failedPaths.length > 0) msg += ` (failed to attach: ${failedPaths.join(", ")})`;
            resolve({ success: true, message: msg });
          } catch (e) {
            resolve({ error: e.toString() });
          }
        }, true, { examineEncryptedParts: true });

      } catch (e) {
        resolve({ error: e.toString() });
      }
    });
  }

  /**
   * Save a single composed message as a draft via nsIMsgSend, without
   * opening a compose window. Shares most of the machinery with
   * sendMessageDirectly but passes SaveAsDraft as the delivery mode.
   *
   * Returns a promise resolving to { success, draftFolderURI?, messageId? }
   * or { error }.
   */
  function saveAsDraftDirectly(composeFields, identity, attachDescs) {
    if (!identity) {
      return Promise.resolve({ error: "No identity available for draft" });
    }
    const SEND_TIMEOUT_MS = 60000;
    return new Promise((resolve) => {
      let settled = false;
      let capturedDraftUri = null;
      const settle = (result) => {
        if (!settled) { settled = true; resolve(result); }
      };

      const timer = Cc["@mozilla.org/timer;1"].createInstance(Ci.nsITimer);
      timer.initWithCallback({
        notify() { settle({ error: "Draft save timed out after " + (SEND_TIMEOUT_MS / 1000) + "s" }); }
      }, SEND_TIMEOUT_MS, Ci.nsITimer.TYPE_ONE_SHOT);

      try {
        const msgSend = Cc["@mozilla.org/messengercompose/send;1"]
          .createInstance(Ci.nsIMsgSend);

        if (identity.email) {
          const name = identity.fullName || "";
          composeFields.from = name ? `"${name}" <${identity.email}>` : identity.email;
        }
        if (identity.organization) {
          composeFields.organization = identity.organization;
        }
        for (const att of descsToMsgAttachments(attachDescs)) {
          composeFields.addAttachment(att);
        }
        const body = composeFields.body || "";

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
        } catch { /* ignore */ }

        const listener = {
          QueryInterface: ChromeUtils.generateQI(["nsIMsgSendListener"]),
          onStartSending() {},
          onProgress() {},
          onSendProgress() {},
          onStatus() {},
          onStopSending(msgID, status) {
            timer.cancel();
            if (Components.isSuccessCode(status)) {
              settle({
                success: true,
                draftFolderURI: capturedDraftUri,
                messageId: (msgID || "").replace(/^<|>$/g, "") || null,
              });
            } else {
              settle({ error: `Draft save failed (status: 0x${(status >>> 0).toString(16)})` });
            }
          },
          onGetDraftFolderURI(msgID, folderURI) {
            capturedDraftUri = folderURI;
          },
          onSendNotPerformed() { timer.cancel(); settle({ error: "Draft save was not performed" }); },
          onTransportSecurityError() { /* not applicable for drafts */ },
        };

        const commonArgs = [
          null,
          identity,
          accountKey,
          composeFields,
          false,
          false,
          Ci.nsIMsgCompDeliverMode.SaveAsDraft,
          null,
          "text/html",
          body,
        ];
        const tailArgs = [
          null,
          null,
          listener,
          "",
          "",
          Ci.nsIMsgCompType.New,
        ];

        let sendResult;
        try {
          sendResult = msgSend.createAndSendMessage(...commonArgs, ...tailArgs);
        } catch (e) {
          const isArgError = (e && e.result === 0x80570001) || String(e).includes("Not enough arguments");
          if (!isArgError) { timer.cancel(); settle({ error: e.toString() }); return; }
          // Legacy 18-arg signature (TB 102-127)
          sendResult = msgSend.createAndSendMessage(...commonArgs, null, null, ...tailArgs);
        }
        if (sendResult && typeof sendResult.then === "function") {
          sendResult.catch(e => { timer.cancel(); settle({ error: e.toString() }); });
        }
      } catch (e) {
        timer.cancel();
        settle({ error: e.toString() });
      }
    });
  }

  /**
   * Batch draft creation. Builds one composeFields per spec and calls
   * saveAsDraftDirectly for each, in sequence (serialized to avoid
   * clobbering nsIMsgSend's per-process state). Returns an array with one
   * result per input draft; failure of one doesn't abort the rest.
   */
  async function createDrafts(drafts) {
    if (!Array.isArray(drafts) || drafts.length === 0) {
      return { error: "drafts must be a non-empty array" };
    }
    if (drafts.length > 50) {
      return { error: "batch limit is 50 drafts per call" };
    }
    const tracer = globalThis.__tbMcpTracer;  // optional
    const batchTraceId = tracer ? tracer.newTraceId() : null;
    if (tracer) {
      tracer.info("create_drafts", "batch.start", {
        trace_id: batchTraceId, count: drafts.length,
      });
    }
    const results = [];
    for (const spec of drafts) {
      try {
        if (!spec || typeof spec !== "object") {
          results.push({ error: "draft spec must be an object" });
          continue;
        }
        const { to, subject, body, cc, bcc, isHtml, from, attachments } = spec;
        if (!to || !subject || typeof body !== "string") {
          results.push({ error: "draft requires to, subject, body" });
          continue;
        }

        const composeFields = Cc["@mozilla.org/messengercompose/composefields;1"]
          .createInstance(Ci.nsIMsgCompFields);
        composeFields.to = to || "";
        composeFields.cc = cc || "";
        composeFields.bcc = bcc || "";
        composeFields.subject = subject || "";

        const formatted = formatBodyHtml(body, isHtml);
        if (isHtml && formatted.includes("<html")) {
          composeFields.body = formatted;
        } else {
          composeFields.body = `<html><head><meta charset="UTF-8"></head><body>${formatted}</body></html>`;
        }

        // Resolve identity via the same path composeMail uses.
        const msgComposeParams = Cc["@mozilla.org/messengercompose/composeparams;1"]
          .createInstance(Ci.nsIMsgComposeParams);
        msgComposeParams.type = Ci.nsIMsgCompType.New;
        msgComposeParams.format = Ci.nsIMsgCompFormat.HTML;
        msgComposeParams.composeFields = composeFields;

        const identityResult = setComposeIdentity(msgComposeParams, from, null);
        if (identityResult && identityResult.error) {
          results.push({ error: identityResult.error });
          continue;
        }

        const { descs: fileDescs, failed: failedPaths } = filePathsToAttachDescs(attachments);

        if (tracer) {
          tracer.debug("create_drafts", "save.dispatch", {
            trace_id: batchTraceId, to, subject,
            identity_email: msgComposeParams.identity && msgComposeParams.identity.email,
            attachments: fileDescs.length,
          });
        }
        const saveResult = await saveAsDraftDirectly(composeFields, msgComposeParams.identity, fileDescs);
        if (saveResult && saveResult.success) {
          const entry = { success: true, draftFolderURI: saveResult.draftFolderURI, messageId: saveResult.messageId, to, subject };
          if (failedPaths.length > 0) entry.failedAttachments = failedPaths;
          results.push(entry);
          if (tracer) {
            tracer.info("create_drafts", "save.success", {
              trace_id: batchTraceId, to, subject,
              draftFolderURI: saveResult.draftFolderURI,
              messageId: saveResult.messageId,
            });
          }
        } else {
          const errResult = saveResult || { error: "unknown draft save failure" };
          results.push(errResult);
          if (tracer) {
            tracer.error("create_drafts", "save.failed", {
              trace_id: batchTraceId, to, subject,
              err: errResult.error,
            });
          }
        }
      } catch (e) {
        results.push({ error: e.toString() });
        if (tracer) {
          tracer.error("create_drafts", "save.threw", {
            trace_id: batchTraceId, err: e && e.message,
          });
        }
      }
    }
    const successCount = results.filter(r => r && r.success).length;
    if (tracer) {
      tracer.info("create_drafts", "batch.end", {
        trace_id: batchTraceId,
        total: drafts.length, succeeded: successCount, failed: drafts.length - successCount,
      });
    }
    return { total: drafts.length, succeeded: successCount, failed: drafts.length - successCount, drafts: results };
  }

  return { composeMail, replyToMessage, forwardMessage, createDrafts };
}
