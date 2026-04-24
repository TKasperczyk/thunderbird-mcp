/* global ExtensionCommon, ChromeUtils, Services, Cc, Ci */
"use strict";

/**
 * Thunderbird MCP Server Extension
 * Exposes email, calendar, and contacts via MCP protocol over HTTP.
 *
 * Architecture: MCP Client <-> mcp-bridge.cjs (stdio<->HTTP) <-> This extension (port 8765)
 *
 * Key quirks documented inline:
 * - MIME header decoding (mime2Decoded* properties)
 * - HTML body charset handling (emojis require HTML entity encoding)
 * - Compose window body preservation (must use New type, not Reply)
 * - IMAP folder sync (msgDatabase may be stale)
 */

const resProto = Cc[
  "@mozilla.org/network/protocol;1?name=resource"
].getService(Ci.nsISubstitutingProtocolHandler);

// Constants, state, and tool schemas live in extension/mcp_server/lib/*.sys.mjs.
// They cannot be imported here at module scope because the resource://
// substitution that makes them resolvable is set up inside getAPI(context).
// Instead, getAPI() and onShutdown() each lazy-import them via
// ChromeUtils.importESModule (synchronous, cached, idempotent).
//
// Mutable counter that the bridge stays in api.js for now -- it's compose-only.
let _tempFileCounter = 0;

var mcpServer = class extends ExtensionCommon.ExtensionAPI {
  getAPI(context) {
    const extensionRoot = context.extension.rootURI;
    const resourceName = "thunderbird-mcp";

    resProto.setSubstitutionWithFlags(
      resourceName,
      extensionRoot,
      resProto.ALLOW_CONTENT_ACCESS
    );

    // Now that the resource:// substitution is in place, pull in the
    // extracted modules. ChromeUtils.importESModule is synchronous and
    // cached -- safe to call here and again in onShutdown.
    const {
      MCP_DEFAULT_PORT, MCP_MAX_PORT_ATTEMPTS,
      MCP_SUPPORTED_PROTOCOL_VERSIONS, MCP_LATEST_PROTOCOL_VERSION,
      MAX_BASE64_SIZE, MAX_REQUEST_BODY, COMPOSE_WINDOW_LOAD_DELAY_MS,
      DEFAULT_MAX_RESULTS,
      PREF_ALLOWED_ACCOUNTS, PREF_DISABLED_TOOLS, PREF_BLOCK_SKIPREVIEW, PREF_PERMISSIONS,
      VALID_GROUPS, VALID_CRUD, CRUD_ORDER,
      UNDISABLEABLE_TOOLS, DEFAULT_DISABLED_TOOLS,
      MAX_SEARCH_RESULTS_CAP, SEARCH_COLLECTION_CAP, INTERNAL_KEYWORDS,
    } = ChromeUtils.importESModule("resource://thunderbird-mcp/mcp_server/lib/constants.sys.mjs");

    const {
      _attachTimers, _tempAttachFiles, _claimedReplyComposeWindows,
    } = ChromeUtils.importESModule("resource://thunderbird-mcp/mcp_server/lib/state.sys.mjs");

    const { tools } = ChromeUtils.importESModule(
      "resource://thunderbird-mcp/mcp_server/lib/tools-schema.sys.mjs"
    );

    // Read serverInfo.version lazily from the extension manifest so a single
    // bump in extension/manifest.json propagates here.
    let _cachedExtVersion = null;
    function getExtVersion() {
      if (_cachedExtVersion) return _cachedExtVersion;
      try {
        const uri = Services.io.newURI("resource://thunderbird-mcp/manifest.json");
        const channel = Services.io.newChannelFromURI(uri, null,
          Services.scriptSecurityManager.getSystemPrincipal(), null,
          Ci.nsILoadInfo.SEC_ALLOW_CROSS_ORIGIN_SEC_CONTEXT_IS_NULL,
          Ci.nsIContentPolicy.TYPE_OTHER);
        const sis = Cc["@mozilla.org/scriptableinputstream;1"]
          .createInstance(Ci.nsIScriptableInputStream);
        sis.init(channel.open());
        const text = sis.read(sis.available());
        sis.close();
        _cachedExtVersion = JSON.parse(text).version || "0.0.0";
      } catch (e) {
        console.warn("thunderbird-mcp: could not read extension manifest version:", e);
        _cachedExtVersion = "0.0.0";
      }
      return _cachedExtVersion;
    }


    // Validate tool metadata: every tool must have valid group and crud fields.
    // This prevents tools from being silently hidden in the settings UI.
    const toolErrors = [];
    for (const tool of tools) {
      if (!tool.group || !VALID_GROUPS.includes(tool.group)) {
        toolErrors.push(`Tool "${tool.name}" has invalid or missing group: "${tool.group}" (valid: ${VALID_GROUPS.join(", ")})`);
      }
      if (!tool.crud || !VALID_CRUD.includes(tool.crud)) {
        toolErrors.push(`Tool "${tool.name}" has invalid or missing crud: "${tool.crud}" (valid: ${VALID_CRUD.join(", ")})`);
      }
    }
    if (toolErrors.length > 0) {
      console.error("thunderbird-mcp: Tool metadata validation failed:\n  " + toolErrors.join("\n  "));
    }

    // Derive ALL_TOOL_NAMES from the tools array (single source of truth)
    const ALL_TOOL_NAMES = tools.map(t => t.name);

    // Group display order for settings UI
    const GROUP_ORDER = { system: 0, messages: 1, folders: 2, contacts: 3, calendar: 4, filters: 5 };
    // Group display labels
    const GROUP_LABELS = { system: "System", messages: "Messages", folders: "Folders", contacts: "Contacts", calendar: "Calendar", filters: "Filters" };

    return {
      mcpServer: {
        start: async function() {
          // Guard against double-start on extension reload (port conflict)
          if (globalThis.__tbMcpStartPromise) {
            return await globalThis.__tbMcpStartPromise;
          }
          const startPromise = (async () => {
          try {
            // Stop any previously running server (e.g. extension reload)
            if (globalThis.__tbMcpServer) {
              try { globalThis.__tbMcpServer.stop(() => {}); } catch { /* ignore */ }
              globalThis.__tbMcpServer = null;
            }
            const { HttpServer } = ChromeUtils.importESModule(
              "resource://thunderbird-mcp/httpd.sys.mjs?" + Date.now()
            );
            const { NetUtil } = ChromeUtils.importESModule(
              "resource://gre/modules/NetUtil.sys.mjs"
            );
            const { MailServices } = ChromeUtils.importESModule(
              "resource:///modules/MailServices.sys.mjs"
            );

            // ── CI-only credential injection ────────────────────────
            //
            // Headless Thunderbird in our integration image can't
            // NSS-encrypt logins.json from outside the app, and plaintext
            // encType=0 entries on disk are silently rejected by TB 128's
            // LoginStore. So we populate credentials from inside the
            // extension via two cooperating paths:
            //
            //   1. nsILoginManager.addLogin()    -- persistent; TB does
            //      the NSS encryption itself. Match on subsequent
            //      password lookups.
            //   2. nsIMsgIncomingServer.password setter -- immediate
            //      in-memory store. Bypasses LoginManager's cache of
            //      "no password for this server" which TB otherwise
            //      keeps for the whole session (classic stale-cache
            //      pattern, mirrors upstream issue #76). Critical to
            //      win the race against the first discoverallboxes URL
            //      which fires before start:'s addLogin() finishes.
            //
            // Gated on extensions.thunderbird-mcp.ciInjectLogins. Empty
            // in production -> entire block is a no-op.
            //
            // ── Structured tracing (Sqlite.sys.mjs-backed) ─────────
            // All observability in this code path flows through one
            // module, lib/trace.sys.mjs, so we end up with a queryable
            // SQLite of events (scenario, event, level, ctx, trace_id)
            // instead of free-form log lines. The file fallback kicks
            // in automatically if Sqlite.sys.mjs isn't available.
            //
            // Sink path comes from a pref so CI and prod can route
            // differently; default is /tmp/thunderbird-mcp-trace.sqlite
            // which matches the CI artifact upload. JSONL fallback is
            // always co-located for visibility-parity on sqlite-open
            // failure (disk full / permission).
            const { makeTrace } = ChromeUtils.importESModule(
              "resource://thunderbird-mcp/mcp_server/lib/trace.sys.mjs"
            );
            const tracePath = Services.prefs.getStringPref(
              "extensions.thunderbird-mcp.tracePath",
              "/tmp/thunderbird-mcp-trace.sqlite"
            );
            const traceFallback = Services.prefs.getStringPref(
              "extensions.thunderbird-mcp.traceFallbackPath",
              "/tmp/thunderbird-mcp-trace.jsonl"
            );
            let tracer;
            try {
              tracer = await makeTrace({
                ChromeUtils,
                dbPath: tracePath,
                fallbackPath: traceFallback,
                minLevel: 1,  // debug and up
              });
            } catch (e) {
              // makeTrace is designed never to throw; if it did, fall
              // back to a no-op shape so callers are still safe.
              tracer = {
                emit: () => {}, trace: () => {}, debug: () => {},
                info: () => {}, warn: () => {}, error: () => {},
                span: async (_s, _e, _c, fn) => (typeof _c === "function") ? _c() : fn(),
                flush: async () => {}, close: async () => {},
                newTraceId: () => "noop",
              };
            }
            // Export so onShutdown can flush + close cleanly. Also
            // makes it reachable from callbacks set up outside start:'s
            // closure (permissions editor, getToolAccessConfig, etc).
            globalThis.__tbMcpTracer = tracer;
            try {
              const PREF_CI_LOGINS = "extensions.thunderbird-mcp.ciInjectLogins";
              const hasValue = Services.prefs.prefHasUserValue(PREF_CI_LOGINS);
              const span_trace = tracer.newTraceId();
              tracer.info("auth.inject", "block.entered", {
                trace_id: span_trace,
                pref_set: hasValue,
                trace_mode: tracer.mode,
                trace_path: tracePath,
              });
              if (hasValue) {
                const raw = Services.prefs.getStringPref(PREF_CI_LOGINS, "");
                if (raw) {
                  const entries = JSON.parse(raw);
                  if (Array.isArray(entries)) {
                    // (1) LoginManager.addLogin for every entry.
                    for (const e of entries) {
                      if (!e || typeof e !== "object") continue;
                      if (!e.origin || !e.username || !e.password) continue;
                      // Match what addLoginAsync() will store -- same
                      // origin + httpRealm tuple. searchLoginsAsync is
                      // the modern API; fall back to findLogins for
                      // older TB.
                      const realmForLookup = e.httpRealm
                        || (e.formActionOrigin ? null : e.origin);
                      let existingFound = false;
                      try {
                        let existing;
                        if (typeof Services.logins.searchLoginsAsync === "function") {
                          existing = await Services.logins.searchLoginsAsync({
                            origin: e.origin,
                            httpRealm: realmForLookup,
                          });
                        } else {
                          existing = Services.logins.findLogins(e.origin, "", realmForLookup);
                        }
                        if (existing && existing.some(l => l.username === e.username)) {
                          existingFound = true;
                          tracer.debug("auth.inject", "login.duplicate_skip", {
                            trace_id: span_trace, origin: e.origin, username: e.username,
                          });
                        }
                      } catch (fe) {
                        tracer.warn("auth.inject", "login.findLogins_threw", {
                          trace_id: span_trace, origin: e.origin, err: fe && fe.message,
                        });
                      }
                      if (existingFound) continue;
                      // LoginManager invariant: an entry MUST have
                      // either a non-null httpRealm OR a non-null
                      // formActionOrigin. Mail-server logins are
                      // non-form-based, so the canonical Mozilla
                      // pattern is httpRealm = the origin URI itself
                      // (matches the example in the on-disk schema
                      // docs: "hostname":"imap://imap.foo.xx",
                      // "httpRealm":"imap://imap.foo.xx"). Auto-fill
                      // when neither is set, so the pref JSON stays
                      // minimal and we don't have to repeat origin.
                      const httpRealm = e.httpRealm
                        || (e.formActionOrigin ? null : e.origin);
                      const info = Cc["@mozilla.org/login-manager/loginInfo;1"]
                        .createInstance(Ci.nsILoginInfo);
                      info.init(
                        e.origin,
                        e.formActionOrigin || null,
                        httpRealm,
                        e.username,
                        e.password,
                        e.usernameField || "",
                        e.passwordField || ""
                      );
                      // TB 128 deprecated the sync addLogin() -- it
                      // throws 'JavaScript component does not have a
                      // method named: "addLogin"'. addLoginAsync is
                      // the modern replacement and returns a promise.
                      // Try async first, fall back to sync for older
                      // TB versions in case anyone runs this CI
                      // harness against ESR 102 etc.
                      try {
                        if (typeof Services.logins.addLoginAsync === "function") {
                          await Services.logins.addLoginAsync(info);
                        } else {
                          Services.logins.addLogin(info);
                        }
                        tracer.info("auth.inject", "login.added", {
                          trace_id: span_trace, origin: e.origin, username: e.username,
                          api: typeof Services.logins.addLoginAsync === "function" ? "addLoginAsync" : "addLogin",
                        });
                      } catch (ex) {
                        tracer.error("auth.inject", "login.addLogin_failed", {
                          trace_id: span_trace, origin: e.origin, err: ex && ex.message,
                        });
                      }
                    }
                    // (2) In-memory server password: walk every account
                    // and match by hostname, set the password directly
                    // on the nsIMsgIncomingServer. This beats TB's stale
                    // "no creds" cache because the server consults its
                    // own `.password` field before going to LoginManager.
                    try {
                      for (const account of MailServices.accounts.accounts) {
                        const server = account.incomingServer;
                        if (!server) continue;
                        for (const e of entries) {
                          if (!e || !e.origin) continue;
                          // Parse `imap://host[:port]` -> compare against
                          // server.hostName (no port). server.type is
                          // `imap`, `pop3`, `smtp`, etc.
                          const m = e.origin.match(/^([a-z0-9]+):\/\/([^:/]+)(?::(\d+))?$/i);
                          if (!m) continue;
                          const [, scheme, host] = m;
                          if (server.type === scheme && server.hostName === host && server.username === e.username) {
                            try {
                              server.password = e.password;
                              tracer.info("auth.inject", "server.password_set", {
                                trace_id: span_trace, scheme, host, server_key: server.key,
                              });
                            } catch (sex) {
                              tracer.error("auth.inject", "server.password_failed", {
                                trace_id: span_trace, scheme, host, err: sex && sex.message,
                              });
                            }
                          }
                        }
                      }
                      // Same treatment for SMTP servers.
                      if (MailServices.outgoingServer && MailServices.outgoingServer.servers) {
                        for (const smtp of MailServices.outgoingServer.servers) {
                          if (!smtp) continue;
                          for (const e of entries) {
                            if (!e || !e.origin || !e.origin.startsWith("smtp://")) continue;
                            const m = e.origin.match(/^smtp:\/\/([^:/]+)/i);
                            if (!m) continue;
                            const host = m[1];
                            const smtpHost = smtp.serverURI ? smtp.serverURI.host : (smtp.hostname || "");
                            if (smtpHost === host) {
                              try {
                                if ("password" in smtp) smtp.password = e.password;
                                tracer.info("auth.inject", "smtp.password_set", {
                                  trace_id: span_trace, host, smtp_key: smtp.key || null,
                                });
                              } catch (sex) {
                                tracer.error("auth.inject", "smtp.password_failed", {
                                  trace_id: span_trace, host, err: sex && sex.message,
                                });
                              }
                            }
                          }
                        }
                      }
                    } catch (ie) {
                      tracer.error("auth.inject", "server_loop.threw", {
                        trace_id: span_trace, err: ie && ie.message,
                      });
                    }
                  }
                }
              }
              tracer.info("auth.inject", "block.complete", { trace_id: span_trace });
            } catch (e) {
              tracer.error("auth.inject", "block.threw", { err: e && e.message });
            }

            let cal = null;
            let CalEvent = null;
            let CalTodo = null;
            try {
              const calModule = ChromeUtils.importESModule(
                "resource:///modules/calendar/calUtils.sys.mjs"
              );
              cal = calModule.cal;
              const { CalEvent: CE } = ChromeUtils.importESModule(
                "resource:///modules/CalEvent.sys.mjs"
              );
              CalEvent = CE;
              const { CalTodo: CT } = ChromeUtils.importESModule(
                "resource:///modules/CalTodo.sys.mjs"
              );
              CalTodo = CT;
            } catch {
              // Calendar not available
            }

            let GlodaMsgSearcher = null;
            try {
              const glodaModule = ChromeUtils.importESModule(
                "resource:///modules/gloda/GlodaMsgSearcher.sys.mjs"
              );
              GlodaMsgSearcher = glodaModule.GlodaMsgSearcher;
            } catch {
              // Gloda not available
            }



            // Auth + connection-info plumbing. Factory takes the XPCOM
            // globals (which are sandbox-only in api.js, not available
            // in .sys.mjs) and returns the operational functions.
            const { makeAuth } = ChromeUtils.importESModule(
              "resource://thunderbird-mcp/mcp_server/lib/auth.sys.mjs"
            );
            const {
              generateAuthToken,
              timingSafeEqual,
              writeConnectionInfo,
              removeConnectionInfo,
            } = makeAuth({ Services, Cc, Ci });

            const authToken = generateAuthToken();

            // Account / folder / tool / send-safety access policy. Same
            // factory pattern -- needs MailServices for account lookups
            // and the four PREF_* / UNDISABLEABLE_TOOLS constants from
            // constants.sys.mjs.
            const { makeAccessControl } = ChromeUtils.importESModule(
              "resource://thunderbird-mcp/mcp_server/lib/access-control.sys.mjs"
            );
            const {
              getAllowedAccountIds,
              isAccountAllowed,
              isSkipReviewBlocked,
              getDisabledTools,
              isToolEnabled,
              isFolderAccessible,
              getAccessibleFolder,
              getAccessibleAccounts,
              getAccountAccess,
            } = makeAccessControl({
              Services,
              MailServices,
              PREF_ALLOWED_ACCOUNTS,
              PREF_DISABLED_TOOLS,
              PREF_BLOCK_SKIPREVIEW,
              UNDISABLEABLE_TOOLS,
              DEFAULT_DISABLED_TOOLS,
            });

            // Account / identity helpers (depends on access-control + MailServices).
            const { makeAccounts } = ChromeUtils.importESModule(
              "resource://thunderbird-mcp/mcp_server/lib/accounts.sys.mjs"
            );
            const {
              listAccounts,
              findIdentityIn,
              findIdentity,
              getIdentityAutoRecipientHeader,
            } = makeAccounts({ MailServices, getAccessibleAccounts });

            // Address-book CRUD.
            const { makeContacts } = ChromeUtils.importESModule(
              "resource://thunderbird-mcp/mcp_server/lib/contacts.sys.mjs"
            );
            const {
              searchContacts,
              createContact,
              updateContact,
              deleteContact,
            } = makeContacts({
              MailServices, Cc, Ci,
              MAX_SEARCH_RESULTS_CAP, DEFAULT_MAX_RESULTS,
            });

            // Mail folder traversal + CRUD. openFolder here is a superset of
            // getAccessibleFolder (adds IMAP refresh + msgDatabase resolution)
            // and is consumed by messages.sys.mjs via dep injection below.
            const { makeFolders } = ChromeUtils.importESModule(
              "resource://thunderbird-mcp/mcp_server/lib/folders.sys.mjs"
            );
            const {
              listFolders, openFolder,
              createFolder, renameFolder, deleteFolder, moveFolder,
              emptyTrash, emptyJunk,
            } = makeFolders({
              MailServices, Cc, Ci, Services,
              getAccessibleAccounts, getAccessibleFolder, isAccountAllowed,
            });

            // Calendar + tasks. Tolerates null cal / CalEvent / CalTodo on
            // TB builds without the calendar component (returns { error })
            // rather than crashing.
            const { makeCalendar } = ChromeUtils.importESModule(
              "resource://thunderbird-mcp/mcp_server/lib/calendar.sys.mjs"
            );
            const {
              listCalendars,
              createEvent, listEvents, updateEvent, deleteEvent,
              createTask, listTasks, updateTask,
            } = makeCalendar({ Services, Cc, Ci, cal, CalEvent, CalTodo });

            // Mail filter rules. Falls back from MailServices.filters to the
            // XPCOM contract ID inside applyFilters for legacy TB builds.
            const { makeFilters } = ChromeUtils.importESModule(
              "resource://thunderbird-mcp/mcp_server/lib/filters.sys.mjs"
            );
            const {
              listFilters, createFilter, updateFilter, deleteFilter,
              reorderFilters, applyFilters,
            } = makeFilters({
              MailServices, Cc, Ci,
              isAccountAllowed, getAccessibleFolder, getAccessibleAccounts,
            });

            // Message search / read / mutate. The biggest single module --
            // owns its own private MIME helpers (escapeHtml, stripHtml,
            // htmlToMarkdown, extractBodyContent, extractPlainTextBody,
            // extractFormattedBody) duplicated with the inline copies api.js
            // still keeps for the compose path. A future commit will collapse
            // those duplicates into encoding.sys.mjs.
            const { makeMessages } = ChromeUtils.importESModule(
              "resource://thunderbird-mcp/mcp_server/lib/messages.sys.mjs"
            );
            const {
              searchMessages, getMessage, getRecentMessages, displayMessage,
              deleteMessages, updateMessage, findMessage,
              getUserTags, markMessageDispositionState,
              inboxInventory, bulkMoveByQuery,
              _compileMessageFilter, _walkFolderMessages, _groupKeyFor,
            } = makeMessages({
              MailServices, NetUtil, Services, Cc, Ci,
              GlodaMsgSearcher,
              isFolderAccessible, getAccessibleFolder, getAccessibleAccounts,
              openFolder,
              SEARCH_COLLECTION_CAP, MAX_SEARCH_RESULTS_CAP,
              DEFAULT_MAX_RESULTS, INTERNAL_KEYWORDS,
              _tempAttachFiles,
            });


            // Compose: sendMail / replyToMessage / forwardMessage. The
            // deepest closure -- attachment timers, compose-window observers,
            // sendMessageDirectly with TB version-fallback for nsIMsgSend.
            const { makeCompose } = ChromeUtils.importESModule(
              "resource://thunderbird-mcp/mcp_server/lib/compose.sys.mjs"
            );
            const {
              composeMail, replyToMessage, forwardMessage, createDrafts,
            } = makeCompose({
              Services, Cc, Ci, MailServices,
              isFolderAccessible, isSkipReviewBlocked,
              findIdentity, findIdentityIn, getIdentityAutoRecipientHeader,
              openFolder,
              findMessage, markMessageDispositionState,
              _attachTimers, _claimedReplyComposeWindows, _tempAttachFiles,
              MAX_BASE64_SIZE, COMPOSE_WINDOW_LOAD_DELAY_MS,
            });

            // Fine-grained permission engine. Sits in front of dispatch so
            // every tool call is policy-checked against the per-tool /
            // per-account / per-folder rules in PREF_PERMISSIONS, with the
            // sandbox default (DEFAULT_DISABLED_TOOLS) as the floor. See
            // lib/permissions.sys.mjs for the full schema.
            // Shared pattern for "complex JSON document stored in a string
            // pref and exposed through an experiment-API method". Wraps the
            // Mozilla workarounds (additionalProperties not honored on
            // parameters; raw throws surface as generic error) so future
            // pref-document stores don't have to re-solve them.
            const { makeJsonPrefStore, wrapApi } = ChromeUtils.importESModule(
              "resource://thunderbird-mcp/mcp_server/lib/json-pref-store.sys.mjs"
            );
            const { makePermissions } = ChromeUtils.importESModule(
              "resource://thunderbird-mcp/mcp_server/lib/permissions.sys.mjs"
            );
            const permissions = makePermissions({
              Services, MailServices,
              PREF_PERMISSIONS,
              UNDISABLEABLE_TOOLS, DEFAULT_DISABLED_TOOLS,
              makeJsonPrefStore,
            });

            // Auto-registry: tools under lib/tools/*.sys.mjs are picked up
            // via the build-time manifest (see scripts/build-xpi.cjs). The
            // shared deps bag is the same set of XPCOM globals and access-
            // control helpers the legacy factories destructure. Any tool
            // that only needs deps from this bag can skip the four-place
            // wiring (tools-schema + domain factory + dispatch switch +
            // toolHandlers bag) and live as one file.
            const { makeToolRegistry } = ChromeUtils.importESModule(
              "resource://thunderbird-mcp/mcp_server/lib/tool-registry.sys.mjs"
            );
            const registry = makeToolRegistry({
              ChromeUtils,
              moduleBasePath: "resource://thunderbird-mcp/mcp_server/lib/tools/",
              manifestPath: "resource://thunderbird-mcp/mcp_server/lib/tools/_manifest.sys.mjs",
              deps: {
                // XPCOM globals
                Services, Cc, Ci, MailServices, NetUtil,
                ChromeUtils,
                // State singletons
                _attachTimers, _tempAttachFiles, _claimedReplyComposeWindows,
                // Access-control
                isAccountAllowed, isFolderAccessible,
                getAccessibleFolder, getAccessibleAccounts,
                isToolEnabled, isSkipReviewBlocked,
                getAccountAccess,
                // Accounts
                findIdentity, findIdentityIn, getIdentityAutoRecipientHeader,
                listAccounts,
                // Folders
                openFolder, listFolders, createFolder, renameFolder,
                deleteFolder, moveFolder, emptyTrash, emptyJunk,
                // Messages + shared scan helpers
                findMessage, getUserTags, markMessageDispositionState,
                searchMessages, getMessage, getRecentMessages, displayMessage,
                deleteMessages, updateMessage,
                inboxInventory, bulkMoveByQuery,
                _compileMessageFilter, _walkFolderMessages, _groupKeyFor,
                // Compose
                composeMail, replyToMessage, forwardMessage, createDrafts,
                // Calendar + tasks
                listCalendars, createEvent, listEvents, updateEvent, deleteEvent,
                createTask, listTasks, updateTask,
                // Contacts
                searchContacts, createContact, updateContact, deleteContact,
                // Filters
                listFilters, createFilter, updateFilter, deleteFilter,
                reorderFilters, applyFilters,
                // Permissions engine
                permissions,
                // Constants
                MAX_BASE64_SIZE, MAX_REQUEST_BODY, COMPOSE_WINDOW_LOAD_DELAY_MS,
                DEFAULT_MAX_RESULTS, MAX_SEARCH_RESULTS_CAP, SEARCH_COLLECTION_CAP,
                INTERNAL_KEYWORDS,
                PREF_ALLOWED_ACCOUNTS, PREF_DISABLED_TOOLS,
                PREF_BLOCK_SKIPREVIEW, PREF_PERMISSIONS,
                UNDISABLEABLE_TOOLS, DEFAULT_DISABLED_TOOLS,
                VALID_GROUPS, VALID_CRUD, CRUD_ORDER,
                MCP_SUPPORTED_PROTOCOL_VERSIONS, MCP_LATEST_PROTOCOL_VERSION,
              },
            });

            // Registry tools win on name conflict. Legacy entries keep
            // working until they are migrated one at a time. tools/list
            // should advertise both.
            const registeredNames = registry.names;
            const mergedTools = [
              ...registry.tools,
              ...tools.filter(t => !registeredNames.has(t.name)),
            ];
            // Expose through globalThis so parallel namespace methods
            // (getToolAccessConfig, etc. -- they live on the return object
            // not inside start()'s closure) can read the same merged list.
            // Cleared on onShutdown. Fallback consumers use legacy `tools`
            // when the server hasn't started yet.
            globalThis.__tbMcpMergedTools = mergedTools;

            // Dispatch: validateToolArgs / coerceToolArgs / callTool plus a
            // few HTTP-handler helpers (readRequestBody, paginate). The HTTP
            // request handler closure stays inline below because it captures
            // res / server which are not worth threading through a factory.
            const { makeDispatch } = ChromeUtils.importESModule(
              "resource://thunderbird-mcp/mcp_server/lib/dispatch.sys.mjs"
            );
            const toolHandlers = {
              listAccounts, findIdentity, findIdentityIn, getIdentityAutoRecipientHeader,
              searchContacts, createContact, updateContact, deleteContact,
              listFolders, openFolder, createFolder, renameFolder, deleteFolder,
              moveFolder, emptyTrash, emptyJunk,
              listCalendars, createEvent, listEvents, updateEvent, deleteEvent,
              createTask, listTasks, updateTask,
              listFilters, createFilter, updateFilter, deleteFilter,
              reorderFilters, applyFilters,
              searchMessages, getMessage, getRecentMessages, displayMessage,
              deleteMessages, updateMessage, findMessage,
              bulkMoveByQuery,
              composeMail, replyToMessage, forwardMessage, createDrafts,
              getAccountAccess,
            };
            const {
              readRequestBody, validateToolArgs, coerceToolArgs, callTool,
            } = makeDispatch({
              NetUtil, tools: mergedTools, toolHandlers, permissions, isToolEnabled,
              registryHandlers: registry.handlers,
            });
























            function stripHtml(html) {
              if (!html) return "";
              let text = String(html);

              // Remove style/script blocks
              text = text.replace(/<script\b[^>]*>[\s\S]*?<\/script>/gi, " ");
              text = text.replace(/<style\b[^>]*>[\s\S]*?<\/style>/gi, " ");

              // Convert block-level tags to newlines before stripping
              text = text.replace(/<br\s*\/?>/gi, "\n");
              text = text.replace(/<\/(p|div|li|tr|h[1-6]|blockquote|pre)>/gi, "\n");
              text = text.replace(/<(p|div|li|tr|h[1-6]|blockquote|pre)\b[^>]*>/gi, "\n");

              // Strip remaining tags
              text = text.replace(/<[^>]+>/g, " ");

              // Decode entities in a single pass
              const NAMED_ENTITIES = {
                nbsp: " ", amp: "&", lt: "<", gt: ">", quot: "\"", apos: "'",
                "#39": "'",
                mdash: "\u2014", ndash: "\u2013", hellip: "\u2026",
                lsquo: "\u2018", rsquo: "\u2019", ldquo: "\u201C", rdquo: "\u201D",
                bull: "\u2022", middot: "\u00B7", ensp: "\u2002", emsp: "\u2003",
                thinsp: "\u2009", zwnj: "\u200C", zwj: "\u200D",
                laquo: "\u00AB", raquo: "\u00BB",
                copy: "\u00A9", reg: "\u00AE", trade: "\u2122", deg: "\u00B0",
                plusmn: "\u00B1", times: "\u00D7", divide: "\u00F7",
                micro: "\u00B5", para: "\u00B6", sect: "\u00A7",
                euro: "\u20AC", pound: "\u00A3", yen: "\u00A5", cent: "\u00A2",
              };
              text = text.replace(/&(#x?[0-9a-fA-F]+|[a-zA-Z]+);/gi, (match, entity) => {
                if (entity.startsWith("#x") || entity.startsWith("#X")) {
                  const cp = parseInt(entity.slice(2), 16);
                  if (!cp || cp > 0x10FFFF) return match;
                  try { return String.fromCodePoint(cp); } catch { return match; }
                }
                if (entity.startsWith("#")) {
                  const cp = parseInt(entity.slice(1), 10);
                  if (!cp || cp > 0x10FFFF) return match;
                  try { return String.fromCodePoint(cp); } catch { return match; }
                }
                return NAMED_ENTITIES[entity.toLowerCase()] || match;
              });

              // Normalize newlines/spaces
              text = text.replace(/\r\n/g, "\n").replace(/\r/g, "\n");
              text = text.replace(/\n{3,}/g, "\n\n");
              text = text.replace(/[ \t\f\v]+/g, " ");
              text = text.replace(/ *\n */g, "\n");
              text = text.trim();
              return text;
            }

            /**
             * Converts HTML to markdown using DOMParser for structure-preserving
             * body extraction. Handles headings, links, bold/italic, lists,
             * blockquotes, code blocks, images, and horizontal rules. Email
             * tables (usually layout, not data) are flattened to text.
             * Falls back to stripHtml if DOMParser is unavailable.
             */
            function htmlToMarkdown(html) {
              if (!html) return "";
              try {
                const doc = new DOMParser().parseFromString(html, "text/html");

                function walkChildren(node) {
                  return Array.from(node.childNodes).map(walk).join("");
                }

                function walk(node) {
                  if (node.nodeType === 3) { // Text
                    return node.textContent.replace(/[ \t]+/g, " ");
                  }
                  if (node.nodeType !== 1) return "";
                  const tag = node.tagName.toLowerCase();
                  const inner = () => walkChildren(node);

                  switch (tag) {
                    case "script": case "style": case "head": return "";
                    case "br": return "\n";
                    case "hr": return "\n\n---\n\n";
                    case "p": case "div": case "section": case "article":
                      return "\n\n" + inner().trim() + "\n\n";
                    case "h1": return "\n\n# " + inner().trim() + "\n\n";
                    case "h2": return "\n\n## " + inner().trim() + "\n\n";
                    case "h3": return "\n\n### " + inner().trim() + "\n\n";
                    case "h4": return "\n\n#### " + inner().trim() + "\n\n";
                    case "h5": return "\n\n##### " + inner().trim() + "\n\n";
                    case "h6": return "\n\n###### " + inner().trim() + "\n\n";
                    case "strong": case "b": {
                      const t = inner().trim();
                      return t ? "**" + t + "**" : "";
                    }
                    case "em": case "i": {
                      const t = inner().trim();
                      return t ? "*" + t + "*" : "";
                    }
                    case "a": {
                      const href = node.getAttribute("href") || "";
                      const text = inner().trim();
                      // Skip empty/anchor-only links and mailto: without text
                      if (!text && !href) return "";
                      if (href && text && text !== href) return `[${text}](${href})`;
                      return text || href;
                    }
                    case "img": {
                      const alt = node.getAttribute("alt") || "";
                      const src = node.getAttribute("src") || "";
                      // Skip tracking pixels (1x1, tiny, or data: without alt)
                      const w = parseInt(node.getAttribute("width")) || 0;
                      const h = parseInt(node.getAttribute("height")) || 0;
                      if ((w > 0 && w <= 3) || (h > 0 && h <= 3)) return "";
                      if (src.startsWith("data:") && !alt) return "";
                      if (src) return `![${alt}](${src})`;
                      return alt;
                    }
                    case "code": return "`" + node.textContent + "`";
                    case "pre": return "\n\n```\n" + node.textContent.trim() + "\n```\n\n";
                    case "blockquote": {
                      const text = inner().trim();
                      return "\n\n" + text.split("\n").map(l => "> " + l).join("\n") + "\n\n";
                    }
                    case "ul": case "ol": return "\n" + inner() + "\n";
                    case "li": {
                      const parent = node.parentElement;
                      const isOl = parent && parent.tagName.toLowerCase() === "ol";
                      return (isOl ? "1. " : "- ") + inner().trim() + "\n";
                    }
                    // Tables: extract text with spacing (email tables are usually layout)
                    case "table": return "\n\n" + inner().trim() + "\n\n";
                    case "tr": return inner().trim() + "\n";
                    case "td": case "th": return inner().trim() + " ";
                    case "thead": case "tbody": case "tfoot": return inner();
                    default: return inner();
                  }
                }

                const body = doc.body || doc.documentElement;
                let result = walk(body);
                // Collapse excessive newlines, trim
                result = result.replace(/\n{3,}/g, "\n\n").trim();
                return result;
              } catch {
                // DOMParser unavailable or parse failure -- fall back to stripHtml
                return stripHtml(html);
              }
            }

            /**
             * Walks the MIME tree to find the raw body content.
             * Returns { text, isHtml } without any format conversion.
             * Does NOT use coerceBodyToPlaintext -- callers that want
             * the raw HTML (for markdown/html output) need this.
             */
            function extractBodyContent(aMimeMsg) {
              if (!aMimeMsg) return { text: "", isHtml: false };
              try {
                function findBody(part, isRoot = false) {
                  const ct = ((part.contentType || "").split(";")[0] || "").trim().toLowerCase();
                  if (ct === "message/rfc822" && !isRoot) return null;
                  if (ct !== "message/rfc822") {
                    if (ct === "text/plain" && part.body) return { text: part.body, isHtml: false };
                    if (ct === "text/html" && part.body) return { text: part.body, isHtml: true };
                  }
                  if (part.parts) {
                    let htmlFallback = null;
                    for (const sub of part.parts) {
                      const r = findBody(sub);
                      if (r && !r.isHtml) return r;
                      if (r && r.isHtml && !htmlFallback) htmlFallback = r;
                    }
                    if (htmlFallback) return htmlFallback;
                  }
                  return null;
                }
                const found = findBody(aMimeMsg, true);
                if (found) return found;
              } catch { /* give up */ }
              return { text: "", isHtml: false };
            }

            /**
             * Extracts plain text body from a MIME message.
             * Uses coerceBodyToPlaintext as fast path, then MIME tree fallback.
             * Used by reply/forward quoting where plain text is appropriate.
             */
            function extractPlainTextBody(aMimeMsg) {
              if (!aMimeMsg) return "";
              try {
                const text = aMimeMsg.coerceBodyToPlaintext();
                if (text) return text;
              } catch { /* fall through */ }
              const { text, isHtml } = extractBodyContent(aMimeMsg);
              return isHtml ? stripHtml(text) : text;
            }

            /**
             * Extracts body from a MIME message in the requested format.
             * For "text": uses coerceBodyToPlaintext fast path (original behavior).
             * For "markdown"/"html": walks MIME tree to find raw HTML content.
             */
            function extractFormattedBody(aMimeMsg, bodyFormat) {
              if (bodyFormat === "text") {
                return { body: extractPlainTextBody(aMimeMsg), bodyIsHtml: false };
              }
              // For markdown/html: need raw MIME content, not coerced text
              const { text, isHtml } = extractBodyContent(aMimeMsg);
              if (!text) {
                // MIME tree empty -- try coerce as last resort
                const fallback = extractPlainTextBody(aMimeMsg);
                return { body: fallback, bodyIsHtml: false };
              }
              if (!isHtml) return { body: text, bodyIsHtml: false };
              if (bodyFormat === "html") return { body: text, bodyIsHtml: true };
              // Default: markdown
              return { body: htmlToMarkdown(text), bodyIsHtml: false };
            }



























































            const server = new HttpServer();

            server.registerPathHandler("/", (req, res) => {
              res.processAsync();

              // Verify auth token on ALL requests (including non-POST) to
              // prevent unauthenticated probing of the server.
              let reqToken = "";
              try {
                reqToken = req.getHeader("Authorization") || "";
              } catch {
                // getHeader throws if header is missing in httpd.sys.mjs
              }
              if (!timingSafeEqual(reqToken, `Bearer ${authToken}`)) {
                res.setStatusLine("1.1", 403, "Forbidden");
                res.setHeader("Content-Type", "application/json; charset=utf-8", false);
                res.write(JSON.stringify({
                  jsonrpc: "2.0",
                  id: null,
                  error: { code: -32600, message: "Invalid or missing auth token" }
                }));
                res.finish();
                return;
              }

              if (req.method !== "POST") {
                res.setStatusLine("1.1", 405, "Method Not Allowed");
                res.setHeader("Allow", "POST", false);
                res.setHeader("Content-Type", "application/json; charset=utf-8", false);
                res.write(JSON.stringify({
                  jsonrpc: "2.0",
                  id: null,
                  error: { code: -32600, message: "Method not allowed" }
                }));
                res.finish();
                return;
              }

              // Reject oversized request bodies to prevent memory exhaustion
              let contentLength = 0;
              try {
                contentLength = parseInt(req.getHeader("Content-Length"), 10) || 0;
              } catch {
                // Header missing — will be 0
              }
              if (contentLength > MAX_REQUEST_BODY) {
                res.setStatusLine("1.1", 413, "Payload Too Large");
                res.setHeader("Content-Type", "application/json; charset=utf-8", false);
                res.write(JSON.stringify({
                  jsonrpc: "2.0",
                  id: null,
                  error: { code: -32600, message: "Request body too large" }
                }));
                res.finish();
                return;
              }

              let message;
              try {
                message = JSON.parse(readRequestBody(req));
              } catch {
                res.setStatusLine("1.1", 200, "OK");
                res.setHeader("Content-Type", "application/json; charset=utf-8", false);
                res.write(JSON.stringify({
                  jsonrpc: "2.0",
                  id: null,
                  error: { code: -32700, message: "Parse error" }
                }));
                res.finish();
                return;
              }

              if (!message || typeof message !== "object" || Array.isArray(message)) {
                res.setStatusLine("1.1", 200, "OK");
                res.setHeader("Content-Type", "application/json; charset=utf-8", false);
                res.write(JSON.stringify({
                  jsonrpc: "2.0",
                  id: null,
                  error: { code: -32600, message: "Invalid Request" }
                }));
                res.finish();
                return;
              }

              const { id, method, params } = message;

              // Notifications don't expect a response
              if (typeof method === "string" && method.startsWith("notifications/")) {
                res.setStatusLine("1.1", 204, "No Content");
                res.finish();
                return;
              }

              (async () => {
                try {
                  let result;
                  switch (method) {
                    case "initialize": {
                      // Per MCP lifecycle: respond with the requested version if
                      // we support it, otherwise the latest version we know.
                      // Behavior never depends on the version inside Thunderbird,
                      // so we accept any well-known protocol version.
                      const requested = params?.protocolVersion;
                      const negotiated = (typeof requested === "string" && MCP_SUPPORTED_PROTOCOL_VERSIONS.has(requested))
                        ? requested
                        : MCP_LATEST_PROTOCOL_VERSION;
                      result = {
                        protocolVersion: negotiated,
                        capabilities: { tools: {} },
                        serverInfo: { name: "thunderbird-mcp", version: getExtVersion() }
                      };
                      break;
                    }
                    case "resources/list":
                      result = { resources: [] };
                      break;
                    case "prompts/list":
                      result = { prompts: [] };
                      break;
                    case "tools/list":
                      // Strip internal metadata (group, crud, title) — only expose MCP-spec fields
                      result = { tools: mergedTools.filter(t => isToolEnabled(t.name)).map(({ name, description, inputSchema }) => ({ name, description, inputSchema })) };
                      break;
                    case "tools/call":
                      if (!params?.name) {
                        throw new Error("Missing tool name");
                      }
                      if (!isToolEnabled(params.name)) {
                        throw new Error(`Tool is disabled: ${params.name}`);
                      }
                      {
                        const toolArgs = coerceToolArgs(params.name, params.arguments || {});
                        const validationErrors = validateToolArgs(params.name, toolArgs);
                        if (validationErrors.length > 0) {
                          throw new Error(`Invalid parameters for '${params.name}': ${validationErrors.join("; ")}`);
                        }
                        result = {
                          content: [{
                            type: "text",
                            text: JSON.stringify(await callTool(params.name, toolArgs), null, 2)
                          }]
                        };
                      }
                      break;
                    default:
                      res.setStatusLine("1.1", 200, "OK");
                      res.setHeader("Content-Type", "application/json; charset=utf-8", false);
                      res.write(JSON.stringify({
                        jsonrpc: "2.0",
                        id: id ?? null,
                        error: { code: -32601, message: "Method not found" }
                      }));
                      res.finish();
                      return;
                  }
                  res.setStatusLine("1.1", 200, "OK");
                  // charset=utf-8 is critical for proper emoji handling in responses
                  res.setHeader("Content-Type", "application/json; charset=utf-8", false);
                  res.write(JSON.stringify({ jsonrpc: "2.0", id: id ?? null, result }));
                } catch (e) {
                  res.setStatusLine("1.1", 200, "OK");
                  res.setHeader("Content-Type", "application/json; charset=utf-8", false);
                  res.write(JSON.stringify({
                    jsonrpc: "2.0",
                    id: id ?? null,
                    error: { code: -32000, message: e.toString() }
                  }));
                }
                res.finish();
              })().catch((e) => {
                console.error("thunderbird-mcp: unhandled dispatch error:", e);
                try { res.finish(); } catch {}
              });
            });

            // Try the default port first, then fall back to nearby ports
            let boundPort = null;
            for (let attempt = 0; attempt < MCP_MAX_PORT_ATTEMPTS; attempt++) {
              const tryPort = MCP_DEFAULT_PORT + attempt;
              try {
                server.start(tryPort);
                boundPort = tryPort;
                break;
              } catch (portErr) {
                if (attempt === MCP_MAX_PORT_ATTEMPTS - 1) {
                  throw new Error(`Could not bind to any port in range ${MCP_DEFAULT_PORT}-${tryPort}: ${portErr}`);
                }
                console.warn(`Port ${tryPort} in use, trying ${tryPort + 1}...`);
              }
            }

            globalThis.__tbMcpServer = server;
            let connFilePath;
            try {
              connFilePath = writeConnectionInfo(boundPort, authToken);
            } catch (writeErr) {
              // Connection file write failed -- stop the orphaned server
              try { server.stop(() => {}); } catch (e) { console.error("thunderbird-mcp: server.stop failed:", e); }
              globalThis.__tbMcpServer = null;
              throw writeErr;
            }
            console.log(`Thunderbird MCP server listening on port ${boundPort}`);
            console.log(`Connection info written to ${connFilePath}`);
            return { success: true, port: boundPort };
          } catch (e) {
            console.error("Failed to start MCP server:", e);
            // Stop server if it was started but something else failed
            if (globalThis.__tbMcpServer) {
              try { globalThis.__tbMcpServer.stop(() => {}); } catch (e) { console.error("thunderbird-mcp: server.stop failed:", e); }
              globalThis.__tbMcpServer = null;
            }
            // Clear cached promise so a retry can attempt to bind again
            globalThis.__tbMcpStartPromise = null;
            removeConnectionInfo();
            return { success: false, error: e.toString() };
          }
          })();
          // Set sentinel BEFORE awaiting to prevent race with concurrent start() calls
          globalThis.__tbMcpStartPromise = startPromise;
          return await startPromise;
        },

        getServerInfo: async function() {
          let port = null;
          let connectionFile = null;
          let buildVersion = null;
          let buildDate = null;

          // Read build info from bundled file via resource: protocol
          try {
            const uri = Services.io.newURI("resource://thunderbird-mcp/buildinfo.json");
            const channel = Services.io.newChannelFromURI(uri, null,
              Services.scriptSecurityManager.getSystemPrincipal(), null,
              Ci.nsILoadInfo.SEC_ALLOW_CROSS_ORIGIN_SEC_CONTEXT_IS_NULL,
              Ci.nsIContentPolicy.TYPE_OTHER);
            const sis = Cc["@mozilla.org/scriptableinputstream;1"]
              .createInstance(Ci.nsIScriptableInputStream);
            sis.init(channel.open());
            const text = sis.read(sis.available());
            sis.close();
            const bi = JSON.parse(text);
            buildVersion = bi.version || bi.commit || null;
            buildDate = bi.builtAt || null;
          } catch (e) {
            // buildinfo.json may legitimately be absent in dev builds; surface anything else.
            if (e?.name !== "NS_ERROR_FILE_NOT_FOUND") {
              console.warn("thunderbird-mcp: read buildinfo failed:", e);
            }
          }

          // Read connection info from temp file using XPCOM file I/O
          try {
            const tmpDir = Services.dirsvc.get("TmpD", Ci.nsIFile);
            tmpDir.append("thunderbird-mcp");
            const connFile = tmpDir.clone();
            connFile.append("connection.json");
            connectionFile = connFile.path;
            if (connFile.exists()) {
              const fis = Cc["@mozilla.org/network/file-input-stream;1"]
                .createInstance(Ci.nsIFileInputStream);
              fis.init(connFile, 0x01, 0, 0);
              const sis = Cc["@mozilla.org/scriptableinputstream;1"]
                .createInstance(Ci.nsIScriptableInputStream);
              sis.init(fis);
              const text = sis.read(sis.available());
              sis.close();
              const data = JSON.parse(text);
              port = data.port || null;
            }
          } catch (e) {
            // Connection file is absent before the server first binds; log other faults.
            if (e?.name !== "NS_ERROR_FILE_NOT_FOUND") {
              console.warn("thunderbird-mcp: read connection info failed:", e);
            }
          }

          return {
            running: !!globalThis.__tbMcpStartPromise,
            port,
            connectionFile,
            buildVersion,
            buildDate,
          };
        },

        getAccountAccessConfig: async function() {
          const { MailServices } = ChromeUtils.importESModule(
            "resource:///modules/MailServices.sys.mjs"
          );
          let allowed = [];
          try {
            const pref = Services.prefs.getStringPref(PREF_ALLOWED_ACCOUNTS, "");
            if (pref) allowed = JSON.parse(pref);
          } catch (e) {
            // Falls back to "all accounts allowed"; surface the corruption.
            console.warn("thunderbird-mcp: account-access pref is not valid JSON:", e.message);
          }

          const accounts = [];
          for (const account of MailServices.accounts.accounts) {
            const server = account.incomingServer;
            accounts.push({
              id: account.key,
              name: server.prettyName,
              type: server.type,
              allowed: allowed.length === 0 || allowed.includes(account.key),
            });
          }
          return {
            mode: allowed.length === 0 ? "all" : "restricted",
            allowedAccountIds: allowed,
            accounts,
          };
        },

        getToolAccessConfig: async function() {
          try {
            // Defensive: if the constants module that ships DEFAULT_DISABLED_TOOLS
            // is missing (e.g. user updated the .mcpb bridge but didn't
            // reinstall the XPI), fall back to an empty default rather than
            // throwing TypeError "...has of undefined".
            const defaultsSet = (DEFAULT_DISABLED_TOOLS && typeof DEFAULT_DISABLED_TOOLS.has === "function")
              ? DEFAULT_DISABLED_TOOLS
              : new Set();

            // First-run state: pref never user-set means we are still serving the
            // sandbox default. The UI surfaces this so users see "X is off by
            // default for safety" rather than thinking they explicitly disabled
            // something they never touched.
            let prefIsUserSet = false;
            try {
              prefIsUserSet = !!Services.prefs.prefHasUserValue(PREF_DISABLED_TOOLS);
            } catch { /* fall through to corrupt path */ }

            let disabled = [];
            let corrupt = false;
            if (!prefIsUserSet) {
              disabled = [...defaultsSet];
            } else {
              try {
                const pref = Services.prefs.getStringPref(PREF_DISABLED_TOOLS, "");
                if (pref) {
                  const parsed = JSON.parse(pref);
                  if (!Array.isArray(parsed)) {
                    corrupt = true;
                  } else {
                    disabled = parsed;
                  }
                }
              } catch {
                corrupt = true;
              }
            }

            // Build tool list with group/crud metadata, sorted by group then CRUD order.
            // Use the merged list populated by start() (registry + legacy) so the
            // Tool Access page surfaces registry-driven tools alongside legacy ones.
            // Falls back to the legacy `tools` array when the server hasn't started
            // yet (settings page can be opened before start() has run).
            const effectiveTools = Array.isArray(globalThis.__tbMcpMergedTools)
              ? globalThis.__tbMcpMergedTools
              : tools;
            const toolList = effectiveTools
              .map(t => ({
                name: t.name,
                group: t.group,
                crud: t.crud,
                enabled: corrupt ? UNDISABLEABLE_TOOLS.has(t.name) : !disabled.includes(t.name),
                undisableable: UNDISABLEABLE_TOOLS.has(t.name),
                defaultDisabled: defaultsSet.has(t.name),
              }))
              .sort((a, b) => {
                const gA = GROUP_ORDER[a.group] ?? 99;
                const gB = GROUP_ORDER[b.group] ?? 99;
                if (gA !== gB) return gA - gB;
                return (CRUD_ORDER[a.crud] ?? 99) - (CRUD_ORDER[b.crud] ?? 99);
              });
            const result = {
              mode: corrupt ? "error" : (disabled.length === 0 ? "all" : "restricted"),
              disabledTools: disabled,
              usingDefaults: !prefIsUserSet && !corrupt,
              groups: GROUP_LABELS,
              tools: toolList,
            };
            if (corrupt) {
              result.error = "Disabled tools preference is corrupt. All non-infrastructure tools are blocked. Save to reset.";
            }
            return result;
          } catch (e) {
            // Last-resort: never let this throw an unwrapped TypeError into the
            // experiment-API boundary -- the UI gets "An unexpected error
            // occurred" with no diagnostic. Return a structured error object
            // the options page can render.
            console.error("thunderbird-mcp: getToolAccessConfig failed:", e);
            return {
              mode: "error",
              error: `getToolAccessConfig failed: ${e && e.message ? e.message : String(e)}`,
              disabledTools: [],
              usingDefaults: false,
              groups: {},
              tools: [],
            };
          }
        },

        setToolAccess: async function(disabledTools) {
          if (!Array.isArray(disabledTools)) {
            return { error: "disabledTools must be an array" };
          }
          // Validate types first, then semantic checks
          if (!disabledTools.every(t => typeof t === "string")) {
            return { error: "All tool names must be strings" };
          }
          // Reject internal sentinel values
          if (disabledTools.includes("__all__")) {
            return { error: "Invalid tool name: __all__" };
          }
          // Validate: can't disable undisableable tools
          const blocked = disabledTools.filter(t => UNDISABLEABLE_TOOLS.has(t));
          if (blocked.length > 0) {
            return { error: `Cannot disable infrastructure tools: ${blocked.join(", ")}` };
          }

          if (disabledTools.length === 0) {
            try { Services.prefs.clearUserPref(PREF_DISABLED_TOOLS); } catch { /* ignore */ }
          } else {
            Services.prefs.setStringPref(PREF_DISABLED_TOOLS, JSON.stringify(disabledTools));
          }
          return {
            success: true,
            mode: disabledTools.length === 0 ? "all" : "restricted",
            disabledTools,
          };
        },

        setAccountAccess: async function(allowedAccountIds) {
          if (!Array.isArray(allowedAccountIds)) {
            return { error: "allowedAccountIds must be an array" };
          }
          const { MailServices } = ChromeUtils.importESModule(
            "resource:///modules/MailServices.sys.mjs"
          );
          const validIds = new Set();
          for (const account of MailServices.accounts.accounts) {
            validIds.add(account.key);
          }
          const invalid = allowedAccountIds.filter(id => !validIds.has(id));
          if (invalid.length > 0) {
            return { error: `Unknown account IDs: ${invalid.join(", ")}` };
          }

          if (allowedAccountIds.length === 0) {
            try { Services.prefs.clearUserPref(PREF_ALLOWED_ACCOUNTS); } catch { /* ignore */ }
          } else {
            Services.prefs.setStringPref(PREF_ALLOWED_ACCOUNTS, JSON.stringify(allowedAccountIds));
          }
          return {
            success: true,
            mode: allowedAccountIds.length === 0 ? "all" : "restricted",
            allowedAccountIds,
          };
        },

        getBlockSkipReview: async function() {
          let blocked = false;
          try {
            blocked = Services.prefs.getBoolPref(PREF_BLOCK_SKIPREVIEW, false);
          } catch { /* ignore */ }
          return { blockSkipReview: blocked };
        },

        setBlockSkipReview: async function(blockSkipReview) {
          if (typeof blockSkipReview !== "boolean") {
            return { error: "blockSkipReview must be a boolean" };
          }
          if (blockSkipReview) {
            Services.prefs.setBoolPref(PREF_BLOCK_SKIPREVIEW, true);
          } else {
            try { Services.prefs.clearUserPref(PREF_BLOCK_SKIPREVIEW); } catch { /* ignore */ }
          }
          return { success: true, blockSkipReview };
        },

        // Fine-grained permission policy (Levels 2-4: per-tool constraints,
        // per-account / per-folder rules, argument whitelists). The schema is
        // documented in lib/permissions.sys.mjs. Stored as a single JSON
        // document in PREF_PERMISSIONS.
        getPermissionsConfig: async function() {
          return wrapApi("getPermissionsConfig", () => {
            if (!permissions || typeof permissions.loadPolicy !== "function") {
              throw new Error("permissions module not loaded -- reinstall the XPI");
            }
            return { policy: permissions.loadPolicy() };
          });
        },
        setPermissionsConfig: async function(policy) {
          return wrapApi("setPermissionsConfig", () => {
            if (!permissions || typeof permissions.savePolicy !== "function") {
              throw new Error("permissions module not loaded");
            }
            return { success: true, policy: permissions.savePolicy(policy) };
          });
        },
        clearPermissionsConfig: async function() {
          return wrapApi("clearPermissionsConfig", () => {
            if (!permissions || typeof permissions.clearPolicy !== "function") {
              throw new Error("permissions module not loaded");
            }
            permissions.clearPolicy();
            return { success: true };
          });
        },
      }
    };
  }

  onShutdown(isAppShutdown) {
    // Stop the HTTP server so the port is released
    if (globalThis.__tbMcpServer) {
      try { globalThis.__tbMcpServer.stop(() => {}); } catch { /* ignore */ }
      globalThis.__tbMcpServer = null;
    }
    // Clear the start promise so a fresh start can occur on reload
    globalThis.__tbMcpStartPromise = null;
    // Clear the merged-tools snapshot that getToolAccessConfig reads
    // through globalThis. On reload, the next start() repopulates it.
    globalThis.__tbMcpMergedTools = null;

    // Always clean up the connection info file so stale tokens don't linger
    // (Inlined here because removeConnectionInfo() is scoped inside start())
    try {
      const tmpDir = Services.dirsvc.get("TmpD", Ci.nsIFile);
      tmpDir.append("thunderbird-mcp");
      const connFile = tmpDir.clone();
      connFile.append("connection.json");
      if (connFile.exists()) {
        connFile.remove(false);
      }
    } catch {
      // Best-effort cleanup
    }

    // Always clean up temp attachment files (even on app shutdown) to avoid
    // leaving sensitive decoded attachments on disk.
    // Re-import state.sys.mjs here -- onShutdown runs in a separate scope from
    // getAPI(). The substitution still resolves because resProto.setSubstitution
    // is cleared further down, after we're done with the module.
    let _tempAttachFiles;
    try {
      ({ _tempAttachFiles } = ChromeUtils.importESModule(
        "resource://thunderbird-mcp/mcp_server/lib/state.sys.mjs"
      ));
    } catch (e) {
      console.warn("thunderbird-mcp: could not load state.sys.mjs in onShutdown:", e);
    }
    if (_tempAttachFiles) {
      for (const tmpPath of _tempAttachFiles) {
        try {
          const f = Cc["@mozilla.org/file/local;1"].createInstance(Ci.nsIFile);
          f.initWithPath(tmpPath);
          if (f.exists()) f.remove(false);
        } catch {}
      }
      _tempAttachFiles.clear();
    }
    if (isAppShutdown) return;
    resProto.setSubstitution("thunderbird-mcp", null);
    Services.obs.notifyObservers(null, "startupcache-invalidate");
  }
};
