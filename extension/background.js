/* global browser */
// Simple background script that starts the MCP server via the experimental API

async function startServer() {
  try {
    const result = await browser.mcpServer.start();
    if (result.success) {
      console.log("MCP server started successfully on port", result.port);
    } else {
      console.error("Failed to start MCP server:", result.error);
    }
  } catch (e) {
    console.error("Error starting MCP server:", e);
  }
}

browser.runtime.onInstalled.addListener(startServer);
browser.runtime.onStartup.addListener(startServer);
