/* global browser */

// Both onInstalled and onStartup fire on first install after restart.
// This flag ensures we only call start() once per background page lifetime.
let startCalled = false;

async function init() {
  if (startCalled) return;
  startCalled = true;
  try {
    const result = await browser.mcpServer.start();
    if (result.success) {
      console.log("MCP server started on port", result.port);
    } else {
      console.error("Failed to start MCP server:", result.error);
    }
  } catch (e) {
    console.error("Error starting MCP server:", e);
  }
}

browser.runtime.onInstalled.addListener(init);
browser.runtime.onStartup.addListener(init);
