/**
 * Auto-discovery registry for MCP tools.
 *
 * Eliminates the four-place wiring that adding a new tool used to
 * require (tools-schema entry + domain-factory handler + dispatch
 * switch case + api.js toolHandlers bag -> one file per tool).
 *
 * Pipeline:
 *
 *   1. At BUILD time, scripts/build-xpi.cjs scans
 *         extension/mcp_server/lib/tools/**\/*.sys.mjs
 *      (skipping files starting with `_`) and writes a manifest at
 *         extension/mcp_server/lib/tools/_manifest.sys.mjs
 *      holding the list of tool module paths.
 *
 *   2. At EXTENSION LOAD time, api.js calls makeToolRegistry(...) which
 *      imports the manifest, imports each listed tool module, and
 *      collects { schema, handler } pairs. The tool list and the
 *      handler map get merged with the legacy tools-schema.sys.mjs
 *      entries (registry wins on name conflict). Dispatch.callTool
 *      checks the registry map first; if a tool name hits there, the
 *      handler is invoked directly with the args object. Legacy tools
 *      continue to flow through the existing switch statement.
 *
 * Tool-file contract (each file under lib/tools/ exports `tool`):
 *
 *   export const tool = {
 *     name: "some_tool",               // MUST be unique across registry + legacy
 *     group: "messages",                // for settings-page grouping
 *     crud: "read",                     // read | create | update | delete
 *     title: "Some Tool",               // human-readable
 *     description: "What this tool does. Claude sees this.",
 *     inputSchema: { type: "object", properties: {...}, required: [...] },
 *     factory(deps) {                   // synchronous factory, returns handler
 *       return async function handler(args) {
 *         // ... implementation ...
 *         return { ...result };
 *       };
 *     },
 *   };
 *
 * The `deps` bag passed to factory() is the shared dependency context
 * api.js assembles -- Services / Cc / Ci, MailServices, state sets,
 * access-control functions, constants, any other cross-cutting helpers.
 * Each tool pulls the subset it needs by destructuring.
 *
 * Why this pattern beats hand-wiring for tool count N:
 *   - per-tool edit scope: 4 files -> 1 file
 *   - delete-a-tool: delete one file, done (vs. 4 places to hunt)
 *   - code review: diff = the tool itself, no scattered context
 *   - AI-assisted maintenance: tokens scale O(1) per tool instead of O(4)
 */

export function makeToolRegistry({
  ChromeUtils,
  moduleBasePath,    // e.g. "resource://thunderbird-mcp/mcp_server/lib/tools/"
  manifestPath,      // e.g. "resource://thunderbird-mcp/mcp_server/lib/tools/_manifest.sys.mjs"
  deps,              // the shared dependency bag passed to each tool's factory
}) {
  const tools = [];
  const handlers = Object.create(null);
  const errors = [];

  let TOOL_MODULES = [];
  try {
    ({ TOOL_MODULES } = ChromeUtils.importESModule(manifestPath));
    if (!Array.isArray(TOOL_MODULES)) {
      errors.push(`_manifest.sys.mjs TOOL_MODULES is not an array (got ${typeof TOOL_MODULES})`);
      TOOL_MODULES = [];
    }
  } catch (e) {
    // Manifest absent is FINE -- the registry just contributes nothing
    // and the legacy tools-schema surface keeps working. This lets us
    // ship the registry infrastructure before any tools are migrated
    // and before the build script has run.
    console.info("thunderbird-mcp: tool-registry manifest not found, registry empty");
    TOOL_MODULES = [];
  }

  for (const relPath of TOOL_MODULES) {
    if (typeof relPath !== "string" || !relPath.endsWith(".sys.mjs")) {
      errors.push(`invalid manifest entry: ${JSON.stringify(relPath)}`);
      continue;
    }
    let mod;
    try {
      mod = ChromeUtils.importESModule(moduleBasePath + relPath);
    } catch (e) {
      errors.push(`failed to import ${relPath}: ${e && e.message ? e.message : String(e)}`);
      continue;
    }
    const t = mod && mod.tool;
    if (!t || typeof t !== "object") {
      errors.push(`${relPath}: missing "tool" export`);
      continue;
    }
    if (!t.name || typeof t.name !== "string") {
      errors.push(`${relPath}: tool.name is required and must be a string`);
      continue;
    }
    if (handlers[t.name]) {
      errors.push(`${relPath}: duplicate tool name "${t.name}" (already registered)`);
      continue;
    }
    if (typeof t.factory !== "function") {
      errors.push(`${relPath}: tool.factory must be a function`);
      continue;
    }
    let handler;
    try {
      handler = t.factory(deps);
    } catch (e) {
      errors.push(`${relPath}: factory threw: ${e && e.message ? e.message : String(e)}`);
      continue;
    }
    if (typeof handler !== "function") {
      errors.push(`${relPath}: factory did not return a function`);
      continue;
    }
    tools.push({
      name: t.name,
      group: t.group || "system",
      crud: t.crud || "read",
      title: t.title || t.name,
      description: t.description || "",
      inputSchema: t.inputSchema || { type: "object", properties: {}, required: [] },
    });
    handlers[t.name] = handler;
  }

  if (errors.length > 0) {
    console.error("thunderbird-mcp: tool-registry errors:\n  " + errors.join("\n  "));
  }

  return {
    tools,
    handlers,
    names: new Set(Object.keys(handlers)),
    errors,
  };
}
