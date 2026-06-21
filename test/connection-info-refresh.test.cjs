const { describe, it } = require("node:test");
const assert = require("node:assert/strict");
const fs = require("fs");
const path = require("path");

describe("Connection info refresh", () => {
  const apiPath = path.join(__dirname, "..", "extension", "mcp_server", "api.js");
  const source = fs.readFileSync(apiPath, "utf8");

  it("periodically rewrites missing or stale connection.json while the server is running", () => {
    assert.match(source, /const CONNECTION_FILE_REFRESH_MS = 30 \* 1000;/);
    assert.match(source, /function ensureConnectionInfo\(port, token\)/);
    assert.match(source, /return writeConnectionInfo\(port, token\);/);
    assert.match(source, /function startConnectionInfoRefresh\(port, token\)/);
    assert.match(source, /Cc\["@mozilla\.org\/timer;1"\]\.createInstance\(Ci\.nsITimer\)/);
    assert.match(source, /timer\.initWithCallback\(\(\) => \{/);
    assert.match(source, /ensureConnectionInfo\(port, token\);/);
    assert.match(source, /Ci\.nsITimer\.TYPE_REPEATING_SLACK/);
  });

  it("starts and stops the refresh timer with the server lifecycle", () => {
    assert.match(source, /startConnectionInfoRefresh\(boundPort, authToken\);/);
    assert.match(source, /function stopConnectionInfoRefreshTimer\(\)/);
    assert.match(source, /globalThis\.__tbMcpConnectionInfoRefreshTimer\.cancel\(\);/);
    assert.match(source, /stopConnectionInfoRefreshTimer\(\);\n\s*\/\/ Clear cached promise/);
    assert.match(source, /stopConnectionInfoRefreshTimer\(\);\n\s*\/\/ Clear the start promise/);
  });
});
