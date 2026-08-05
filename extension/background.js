// `browser` is declared as a global in eslint.config.mjs for the
// extension/ file group. The per-file /* global browser */ comment
// would trigger no-redeclare.
async function init() {
  try {
    const translationResult = await browser.mcpServer.enableLocalTranslations();
    if (translationResult?.error) {
      console.error("Could not enable local translations:", translationResult.error);
    }
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
browser.messageDisplay.onMessageDisplayed.addListener((tab) => {
  browser.messageDisplayAction.setTitle({
    tabId: tab.id,
    title: "AI / Basit Çeviri",
  }).catch((e) => console.error("Could not reset translation button title:", e));
});

// Also call init() directly — the event listeners above don't fire when
// a user disables and re-enables the extension.
init();
