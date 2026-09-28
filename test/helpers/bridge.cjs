// Requiring the CLI starts it even under a desktop bootstrap. Keep the opt-out
// local to this import so child-process CLI tests retain their normal behavior.
const previous = process.env.THUNDERBIRD_MCP_NO_AUTOSTART;
process.env.THUNDERBIRD_MCP_NO_AUTOSTART = '1';
try {
  module.exports = require('../../mcp-bridge.cjs');
} finally {
  if (previous === undefined) delete process.env.THUNDERBIRD_MCP_NO_AUTOSTART;
  else process.env.THUNDERBIRD_MCP_NO_AUTOSTART = previous;
}
