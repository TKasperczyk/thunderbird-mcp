/* global ExtensionCommon, ChromeUtils, Services, Cc, Ci */

"use strict";

// Get the resource protocol handler to register our custom resource:// URL
const resProto = Cc[
  "@mozilla.org/network/protocol;1?name=resource"
].getService(Ci.nsISubstitutingProtocolHandler);

var mcpServer = class extends ExtensionCommon.ExtensionAPI {
  getAPI(context) {
    const extensionRoot = context.extension.rootURI;

    // Register resource://thunderbird-mcp/ to point to our extension root
    const resourceName = "thunderbird-mcp";
    resProto.setSubstitutionWithFlags(
      resourceName,
      extensionRoot,
      resProto.ALLOW_CONTENT_ACCESS
    );

    return {
      mcpServer: {
        start: async function() {
          try {
            // Now we can import from our registered resource:// URL
            const { HttpServer } = ChromeUtils.importESModule(
              "resource://thunderbird-mcp/httpd.sys.mjs?" + Date.now()
            );
            const { NetUtil } = ChromeUtils.importESModule(
              "resource://gre/modules/NetUtil.sys.mjs"
            );
            const { MailServices } = ChromeUtils.importESModule(
              "resource:///modules/MailServices.sys.mjs"
            );
            const { Gloda } = ChromeUtils.importESModule(
              "resource:///modules/gloda/GlodaPublic.sys.mjs"
            );

            // Try to import calendar utils (may not be available)
            let cal = null;
            try {
              const calModule = ChromeUtils.importESModule(
                "resource:///modules/calendar/calUtils.sys.mjs"
              );
              cal = calModule.cal;
            } catch (e) {
              console.warn("Calendar module not available:", e);
            }

            const tools = [
              {
                name: "searchMessages",
                title: "Search Mail",
                description: "Find messages using Thunderbird's search index. Returns message id, subject, sender, and folder.",
                inputSchema: {
                  type: "object",
                  properties: {
                    query: {
                      type: "string",
                      description: "Text to search for in messages"
                    }
                  },
                  required: ["query"],
                },
              },
              {
                name: "sendMail",
                title: "Send Mail",
                description: "Compose and send a new email message",
                inputSchema: {
                  type: "object",
                  properties: {
                    to: {
                      type: "string",
                      description: "Recipient email address"
                    },
                    subject: {
                      type: "string",
                      description: "Email subject line"
                    },
                    body: {
                      type: "string",
                      description: "Email body text"
                    },
                  },
                  required: ["to", "subject", "body"],
                },
              },
              {
                name: "listCalendars",
                title: "List Calendars",
                description: "Return all configured calendars in Thunderbird",
                inputSchema: { type: "object", properties: {}, required: [] },
              },
              {
                name: "searchContacts",
                title: "Search Contacts",
                description: "Find contacts by email address in Thunderbird's address books",
                inputSchema: {
                  type: "object",
                  properties: {
                    query: {
                      type: "string",
                      description: "Email address or partial email to search for"
                    }
                  },
                  required: ["query"],
                },
              },
            ];

            function readRequestBody(request) {
              const stream = request.bodyInputStream;
              const available = stream.available();
              return NetUtil.readInputStreamToString(stream, available);
            }

            async function searchMessages(query) {
              const glodaQuery = Gloda.newQuery(Gloda.NOUN_MESSAGE).text(query);
              const collection = glodaQuery.getCollection();
              await collection.wait();
              return collection.items.slice(0, 50).map(m => ({
                id: m.headerMessageID,
                subject: m.subject,
                from: m.from ? m.from.toString() : null,
                date: m.date ? new Date(m.date / 1000).toISOString() : null,
                folder: m.folder ? m.folder.name : null,
                folderURI: m.folder ? m.folder.URI : null,
              }));
            }

            async function searchContacts(query) {
              const results = [];
              const lowerQuery = query.toLowerCase();
              for (const book of MailServices.ab.directories) {
                for (const card of book.childCards) {
                  if (card.isMailList) continue;
                  const email = (card.primaryEmail || "").toLowerCase();
                  const displayName = (card.displayName || "").toLowerCase();
                  if (email.includes(lowerQuery) || displayName.includes(lowerQuery)) {
                    results.push({
                      id: card.UID,
                      displayName: card.displayName,
                      email: card.primaryEmail,
                      addressBook: book.dirName,
                    });
                  }
                }
              }
              return results.slice(0, 50);
            }

            function listCalendars() {
              if (!cal) {
                return [];
              }
              try {
                const calendars = cal.manager.getCalendars();
                return calendars.map(c => ({
                  id: c.id,
                  name: c.name,
                  type: c.type,
                  uri: c.uri ? c.uri.spec : null,
                  readOnly: c.readOnly,
                }));
              } catch (e) {
                console.error("Error listing calendars:", e);
                return [];
              }
            }

            async function callTool(name, args) {
              switch (name) {
                case "searchMessages":
                  return await searchMessages(args.query);
                case "sendMail":
                  // TODO: Implement via compose API
                  return { error: "sendMail not yet implemented - use Thunderbird UI for now" };
                case "listCalendars":
                  return listCalendars();
                case "searchContacts":
                  return await searchContacts(args.query);
                default:
                  throw new Error(`Unknown tool: ${name}`);
              }
            }

            const server = new HttpServer();

            server.registerPathHandler("/", (req, res) => {
              // Use processAsync for async request handling
              res.processAsync();

              if (req.method !== "POST") {
                res.setStatusLine("1.1", 405, "Method Not Allowed");
                res.write("POST only");
                res.finish();
                return;
              }

              let message;
              try {
                const body = readRequestBody(req);
                message = JSON.parse(body);
              } catch (e) {
                res.setStatusLine("1.1", 400, "Bad Request");
                res.write(JSON.stringify({ error: "Invalid JSON" }));
                res.finish();
                return;
              }

              const { id, method, params } = message;

              (async () => {
                try {
                  let result;
                  switch (method) {
                    case "tools/list":
                      result = { tools };
                      break;
                    case "tools/call":
                      if (!params || !params.name) {
                        throw new Error("Missing tool name in params");
                      }
                      const toolResult = await callTool(params.name, params.arguments || {});
                      result = {
                        content: [
                          {
                            type: "text",
                            text: JSON.stringify(toolResult, null, 2)
                          }
                        ]
                      };
                      break;
                    default:
                      res.setStatusLine("1.1", 404, "Not Found");
                      res.write(JSON.stringify({ error: `Unknown method: ${method}` }));
                      res.finish();
                      return;
                  }
                  res.setStatusLine("1.1", 200, "OK");
                  res.setHeader("Content-Type", "application/json", false);
                  res.write(JSON.stringify({ jsonrpc: "2.0", id, result }));
                } catch (e) {
                  console.error("MCP request error:", e);
                  res.setStatusLine("1.1", 200, "OK");
                  res.setHeader("Content-Type", "application/json", false);
                  res.write(JSON.stringify({
                    jsonrpc: "2.0",
                    id,
                    error: { code: -32000, message: e.toString() }
                  }));
                }
                res.finish();
              })();
            });

            server.start(8765);
            console.log("MCP server listening on port 8765");
            return { success: true, port: 8765 };
          } catch (e) {
            console.error("Failed to start MCP server:", e);
            return { success: false, error: e.toString() };
          }
        }
      }
    };
  }

  onShutdown(isAppShutdown) {
    if (isAppShutdown) {
      return;
    }
    // Unregister our resource:// URL
    resProto.setSubstitution("thunderbird-mcp", null);
    // Flush caches
    Services.obs.notifyObservers(null, "startupcache-invalidate");
  }
};
