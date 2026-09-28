"use strict";

const { test } = require("node:test");
const assert = require("node:assert/strict");
const fs = require("node:fs");
const os = require("node:os");
const path = require("node:path");
const net = require("node:net");
const { spawn } = require("node:child_process");
const { setTimeout: delay } = require("node:timers/promises");

const source = fs.readFileSync(path.resolve(__dirname, "../extension/mcp_server/api.js"), "utf8");
function snippet(name) {
  const start = source.indexOf(`// BEGIN ${name}`);
  const end = source.indexOf(`// END ${name}`, start);
  assert.ok(start >= 0 && end > start, `Missing production marker: ${name}`);
  return source.slice(start, end);
}

function findThunderbird() {
  const name = process.env.THUNDERBIRD_BIN || (process.platform === "win32" ? "thunderbird.exe" : "thunderbird");
  const candidates = path.isAbsolute(name) ? [name]
    : (process.env.PATH || "").split(path.delimiter).map(dir => path.resolve(dir, name));
  for (const candidate of candidates) {
    try {
      fs.accessSync(candidate, fs.constants.X_OK);
      return candidate;
    } catch { /* Try the next PATH entry. */ }
  }
  if (process.env.THUNDERBIRD_BIN) throw Error(`THUNDERBIRD_BIN is not executable: ${name}`);
  return null;
}

async function terminateProcessTree(child, graceMs = 500) {
  try {
    if (!child.pid) return;
    if (process.platform === "win32") {
      await new Promise((resolve, reject) => {
        const killer = spawn("taskkill", ["/PID", String(child.pid), "/T", "/F"], {
          stdio: "ignore", windowsHide: true,
        });
        const timer = setTimeout(() => {
          killer.kill("SIGKILL");
          child.kill("SIGKILL");
          finish(Error("taskkill timed out while stopping Thunderbird"));
        }, 1500);
        function finish(error) {
          clearTimeout(timer);
          killer.unref();
          if (error) { child.kill("SIGKILL"); reject(error); }
          else resolve();
        }
        killer.once("error", finish);
        killer.once("exit", code => finish(code && child.exitCode === null
          ? Error(`taskkill failed with exit code ${code}`) : undefined));
      });
    } else {
      // detached:true gives the launcher and its descendants an isolated
      // process group. The group can outlive the launcher, so never use the
      // launcher's exitCode or pipe-dependent 'close' event to skip killing it.
      for (const signal of ["SIGTERM", "SIGKILL"]) {
        try { process.kill(-child.pid, signal); }
        catch (error) { if (error.code === "ESRCH") break; throw error; }
        if (signal === "SIGTERM") await delay(graceMs);
      }
    }
  } finally {
    // Pipe closure is not a prerequisite for teardown. A descendant may hold
    // the descriptors open after its launcher has already exited.
    child.stdin?.destroy();
    child.stdout?.destroy();
    child.stderr?.destroy();
    child.unref();
  }
}

// Minimal Marionette client; no Selenium/geckodriver dependency. Framing counts
// UTF-8 bytes, not JS characters. The protocol is documented at
// https://firefox-source-docs.mozilla.org/remote/marionette/Protocol.html.
function connectMarionette(port, signal) {
  return new Promise((resolve, reject) => {
    const socket = net.createConnection({ host: "127.0.0.1", port });
    const pending = new Map();
    let buffer = Buffer.alloc(0);
    let id = 0;
    let failed = false;
    const helloTimer = setTimeout(() => fail(Error("Marionette handshake timed out")), 10000);
    const abort = () => fail(Error("Gecko test cancelled"));
    function fail(error) {
      if (failed) return;
      failed = true;
      clearTimeout(helloTimer);
      signal?.removeEventListener("abort", abort);
      reject(error);
      for (const request of pending.values()) {
        clearTimeout(request.timer);
        request.reject(error);
      }
      pending.clear();
      socket.destroy();
    }
    function command(name, parameters = {}) {
      return new Promise((resolveCommand, rejectCommand) => {
        if (socket.destroyed) return rejectCommand(Error("Marionette connection is closed"));
        const messageId = ++id;
        const timer = setTimeout(() => fail(Error(`Marionette command timed out: ${name}`)), 10000);
        pending.set(messageId, { resolve: resolveCommand, reject: rejectCommand, timer });
        const payload = JSON.stringify([0, messageId, name, parameters]);
        socket.write(`${Buffer.byteLength(payload)}:${payload}`);
      });
    }
    socket.on("error", fail);
    socket.on("close", () => fail(Error("Marionette connection closed")));
    signal?.addEventListener("abort", abort, { once: true });
    if (signal?.aborted) { abort(); return; }
    socket.on("data", chunk => {
      buffer = Buffer.concat([buffer, chunk]);
      try {
        while (true) {
          const colon = buffer.indexOf(58);
          if (colon < 0) return;
          const size = Number(buffer.subarray(0, colon).toString("ascii"));
          if (!Number.isInteger(size) || size < 0) throw Error("Invalid Marionette frame length");
          if (buffer.length < colon + 1 + size) return;
          const packet = JSON.parse(buffer.subarray(colon + 1, colon + 1 + size).toString("utf8"));
          buffer = buffer.subarray(colon + 1 + size);
          if (!Array.isArray(packet)) {
            assert.equal(packet.marionetteProtocol, 3);
            clearTimeout(helloTimer);
            resolve({ command, close: () => fail(Error("Marionette client closed")) });
            continue;
          }
          const request = pending.get(packet[1]);
          if (!request) continue;
          pending.delete(packet[1]);
          clearTimeout(request.timer);
          if (packet[2]) request.reject(Error(`${packet[2].error}: ${packet[2].message}\n${packet[2].stacktrace}`));
          else request.resolve(packet[3]);
        }
      } catch (error) { fail(error); }
    });
  });
}

test("Gecko teardown kills a descendant holding stdout after its launcher exits", {
  skip: process.platform === "win32" ? "POSIX process-group regression fixture" : false,
  timeout: 5000,
}, async t => {
  const holder = 'process.on("SIGTERM", () => {}); process.stdout.write(String(process.pid) + "\\n"); setInterval(() => {}, 1000);';
  const launcher = `require("node:child_process").spawn(process.execPath, ["-e", ${JSON.stringify(holder)}], { stdio: ["ignore", "inherit", "inherit"] }).unref();`;
  const child = spawn(process.execPath, ["-e", launcher], { detached: true, stdio: ["ignore", "pipe", "pipe"] });
  t.after(() => terminateProcessTree(child, 50));
  let closed = false;
  let descendantPid;
  child.once("close", () => { closed = true; });
  await new Promise((resolve, reject) => {
    let output = "";
    let exited = false;
    const timer = setTimeout(() => reject(Error("Teardown fixture failed to start")), 2000);
    function ready() {
      if (exited && descendantPid) { clearTimeout(timer); resolve(); }
    }
    child.stdout.on("data", chunk => {
      output += chunk.toString();
      const pid = output.match(/^(\d+)\n/);
      if (pid) descendantPid = Number(pid[1]);
      ready();
    });
    child.once("exit", () => { exited = true; ready(); });
    child.once("error", error => { clearTimeout(timer); reject(error); });
  });
  assert.equal(child.exitCode, 0);
  assert.equal(closed, false, "descendant must still hold the launcher's stdout open");
  const started = Date.now();
  await terminateProcessTree(child, 100);
  assert.ok(Date.now() - started < 2000, "teardown must not wait for pipe closure");
  function descendantRunning() {
    try {
      process.kill(descendantPid, 0);
      // Minimal container init processes may delay reaping an orphan. A zombie
      // is already dead; unlike a live process, it cannot retain the pipe.
      if (process.platform === "linux") {
        const stat = fs.readFileSync(`/proc/${descendantPid}/stat`, "utf8");
        if (/^[ZX]/.test(stat.slice(stat.lastIndexOf(")") + 2))) return false;
      }
      return true;
    } catch (error) {
      if (error.code === "ESRCH" || error.code === "ENOENT") return false;
      throw error;
    }
  }
  const deadline = Date.now() + 1000;
  while (descendantRunning() && Date.now() < deadline) await delay(25);
  assert.equal(descendantRunning(), false, "SIGKILL must reach the TERM-ignoring descendant");
});

const thunderbird = findThunderbird();
test("real Gecko Experiment globals and message text conversion", {
  skip: thunderbird ? false : "Thunderbird is not installed; set THUNDERBIRD_BIN to enable real Gecko coverage",
  timeout: 45000,
}, async t => {
  const profile = fs.mkdtempSync(path.join(os.tmpdir(), "tb-mcp-gecko-"));
  let child;
  let client;
  t.after(async () => {
    client?.close();
    try { if (child) await terminateProcessTree(child); }
    finally {
      // This is the isolated profile created by this test, never a user's profile.
      fs.rmSync(profile, { recursive: true, force: true, maxRetries: 3, retryDelay: 100 });
    }
  });
  fs.writeFileSync(path.join(profile, "user.js"), [
    'user_pref("marionette.port", 0);',
    'user_pref("marionette.log.level", "Info");',
    'user_pref("mail.provider.suppress_dialog_on_startup", true);',
    'user_pref("mail.shell.checkDefaultClient", false);',
    'user_pref("mailnews.start_page.enabled", false);',
    'user_pref("datareporting.policy.dataSubmissionEnabled", false);',
    'user_pref("toolkit.telemetry.reportingpolicy.firstRun", false);',
  ].join("\n"));
  child = spawn(thunderbird, [
    "--headless", "--no-remote", "--profile", profile,
    "--marionette", "--remote-allow-system-access",
  ], {
    detached: process.platform !== "win32", stdio: ["ignore", "pipe", "pipe"],
    env: { ...process.env, MOZ_HEADLESS: "1" },
  });
  let log = "";
  const port = await new Promise((resolve, reject) => {
    let settled = false;
    const timer = setTimeout(() => finish(Error(`Thunderbird did not start Marionette:\n${log}`)), 15000);
    const abort = () => finish(Error("Gecko test cancelled during startup"));
    function finish(error, value) {
      if (settled) return;
      settled = true;
      clearTimeout(timer);
      t.signal?.removeEventListener("abort", abort);
      if (error) reject(error);
      else resolve(value);
    }
    const receive = chunk => {
      log = (log + chunk.toString()).slice(-8192);
      const listening = log.match(/Marionette\s+INFO\s+Listening on port (\d+)/);
      if (listening) finish(undefined, Number(listening[1]));
    };
    child.stdout.on("data", receive);
    child.stderr.on("data", receive);
    child.once("error", finish);
    child.once("exit", code => finish(Error(`Thunderbird exited (${code}):\n${log}`)));
    t.signal?.addEventListener("abort", abort, { once: true });
    if (t.signal?.aborted) abort();
  });
  client = await connectMarionette(port, t.signal);
  const session = await client.command("WebDriver:NewSession", { capabilities: { alwaysMatch: {} } });
  t.diagnostic(`Thunderbird ${session.capabilities.browserVersion}: real DOMParser in an isolated system sandbox`);
  await client.command("Marionette:SetContext", { value: "chrome" });
  const production = [
    "EXPERIMENT GLOBAL IMPORTS", "MCP TEXT SANITIZATION", "MESSAGE TEXT CONVERSION",
    "RAW MIME PARSING HELPERS", "ENCRYPTED MESSAGE GUARD",
  ].map(snippet).join("\n");

  async function evaluate(body) {
    const result = await client.command("WebDriver:ExecuteScript", {
      sandbox: "system", newSandbox: true,
      args: [production, body],
      script: `
        const Cu = Components.utils;
        const scope = Cu.Sandbox(Services.scriptSecurityManager.getSystemPrincipal(), {
          sandboxName: "Thunderbird MCP Experiment test", wantGlobalProperties: [],
        });
        scope.Cu = Cu;
        try {
          return Cu.evalInSandbox(
            'const globalsBeforeImport = Object.fromEntries(["DOMParser", "atob", "btoa", "TextDecoder", "TextEncoder", "URL"].map(name => [name, typeof globalThis[name]]));\\n' +
            arguments[0] + '\\nJSON.stringify((() => {' + arguments[1] + '\\n})())', scope
          );
        } finally { Cu.nukeSandbox(scope); }
      `,
    });
    return JSON.parse(result.value);
  }

  await t.test("imports actually supply missing parser, base64 and decoding globals", async () => {
    const result = await evaluate(`
      return {
        before: globalsBeforeImport,
        tag: new DOMParser().parseFromString("<p>readable</p>", "text/html").body.firstChild.tagName,
        encoded: btoa(String.fromCharCode(0, 255)),
        decoded: Array.from(atob("AP8="), character => character.charCodeAt(0)),
        utf8: new TextDecoder("utf-8").decode(Uint8Array.of(0xc3, 0xa9)),
      };
    `);
    for (const [name, type] of Object.entries(result.before)) assert.equal(type, "undefined", `${name} must not be supplied by the test`);
    assert.equal(result.tag, "P");
    assert.equal(result.encoded, "AP8=");
    assert.deepEqual(result.decoded, [0, 255]);
    assert.equal(result.utf8, "é");
  });

  await t.test("real HTML/CSS parsing removes hidden content and decoded invisible characters", async () => {
    const result = await evaluate(`
      const invisible = [0x200b, 0x2060, 0xfeff, 0x202a, 0x202b, 0x202c, 0x202d, 0x202e, 0x2066, 0x2067, 0x2068, 0x2069, 0xe0000, 0xe007f];
      const html = '<!doctype html><html><head><title>head-secret</title><style>style-secret</style></head><body>' +
        '<p>Visible &amp; readable 👩‍💻' + invisible.map(code => '&#' + code + ';').join('') + '</p>' +
        '<div hidden="false">hidden-secret</div><span style="DISPLAY : NONE !important">display-secret</span>' +
        '<div style="VISIBILITY: HIDDEN">visibility-secret</div><span style="font-size: 0px">font-secret</span>' +
        '<div style="opacity: 0.0">opacity-secret</div><script>script-secret</script>' +
        '<template>template-secret<template>nested-secret</template></template><!--comment-secret-->' +
        '<pre>code-visible<!--pre-secret--><span hidden>pre-hidden-secret</span><template>pre-template-secret</template></pre>' +
        '<code>inline-visible<!--code-secret--><span style="display:none">code-hidden-secret</span></code></body></html>';
      return [stripHtml(html), htmlToMarkdown(html)];
    `);
    for (const text of result) {
      assert.match(text, /Visible & readable 👩‍💻/);
      assert.match(text, /code-visible/);
      assert.match(text, /inline-visible/);
      assert.doesNotMatch(text, /secret|withheld/);
      assert.doesNotMatch(text, /[\u200B\u2060\uFEFF\u202A-\u202E\u2066-\u2069\u{E0000}-\u{E007F}]/u);
    }
  });

  await t.test("hidden html and body roots do not expose detached subtrees", async () => {
    const result = await evaluate(`
      return ['<html hidden><body>root-secret</body></html>', '<html><body hidden>root-secret</body></html>',
        '<html style="opacity:0"><body>root-secret</body></html>', '<html><body style="font-size:0">root-secret</body></html>']
        .flatMap(html => [stripHtml(html), htmlToMarkdown(html)]);
    `);
    assert.deepEqual(result, Array(8).fill(""));
  });

  await t.test("Markdown allows normalized http, https and mailto destinations only", async () => {
    const allowed = [
      ["http://safe.test/plain", "http://safe.test/plain"],
      ["HTTPS://safe.test/upper", "HTTPS://safe.test/upper"],
      [" \u0001\tHtTpS://safe.test/leading", "HtTpS://safe.test/leading"],
      ["\tMaIl\nTo:reader@example.test", "MaIlTo:reader@example.test"],
      ["h\ttt\r\nps://safe.test/split", "https://safe.test/split"],
      ["https&#58;//safe.test/entity", "https://safe.test/entity"],
    ];
    const blocked = [
      "javascript:alert(1)", " \u0001\tJaVa\nScRiPt:alert(1)", "file:///private/secret", "data:text/html,secret",
      "vbscript:msgbox(1)", "cid:secret", "/relative", "../relative", "//remote.test/path", "#fragment", "", "ftp://remote.test/file",
    ];
    const destinations = [...allowed.map(([href]) => href), ...blocked];
    const result = await evaluate(`return ${JSON.stringify(destinations)}.map(href => htmlToMarkdown('<a href="' + href + '">label</a>'));`);
    assert.deepEqual(result.slice(0, allowed.length), allowed.map(([, href]) => `[label](${href})`));
    assert.deepEqual(result.slice(allowed.length), blocked.map(() => "label"));
  });

  await t.test("images contribute alt text without remote, data or cid URLs", async () => {
    const result = await evaluate(`return htmlToMarkdown(
      '<p><img alt="Useful illustration" src="https://image.test/photo" srcset="https://image.test/large 2x"></p>' +
      '<img alt="Pixel description" width="1" height="1" src="https://tracking.test/pixel">' +
      '<img src="https://tracking.test/empty"><img alt="Attachment" src="cid:image-secret">' +
      '<img alt="Embedded" src="data:image/png;base64,secret">');`);
    assert.match(result, /Useful illustration/);
    assert.match(result, /Attachment/);
    assert.match(result, /Embedded/);
    assert.doesNotMatch(result, /https:|data:|cid:|secret|!\[/);
  });

  await t.test("text preserves ordinary punctuation while escaping link, image and HTML syntax", async () => {
    const ordinary = "some_path C# 5*3 snake_case_name `ticks` _ * # | ~ !";
    const html = `<p>${ordinary}</p><p>![x](y) [a](javascript:b) &lt;img src=x&gt;</p>`;
    const result = await evaluate(`return htmlToMarkdown(${JSON.stringify(html)});`);
    assert.equal(result, "some_path C# 5*3 snake_case_name \\`ticks\\` _ * # | ~ !\n\n" + String.raw`\!\[x\](y) \[a\](javascript:b) \<img src=x\>`);
  });

  await t.test("a literal bang cannot turn a link across a DOM boundary into an image", async () => {
    const link = '<a href="https://tracker.test/p">x</a>';
    const fixtures = [
      "!" + link, "!<!--gap-->" + link, "!<span></span>" + link,
      "!<span>" + link + "</span>", "<span>!</span>" + link,
      "!&#x200b;" + link, '<img alt="!">' + link,
      '!<img alt="Illustration" src="https://image.test/p">',
      '!<img alt="[x](https://tracker.test/p)" src="https://image.test/p">', "Thanks!",
    ];
    const results = await evaluate(`return ${JSON.stringify(fixtures)}.map(htmlToMarkdown);`);
    assert.deepEqual(results, [
      ...Array(7).fill("\\![x](https://tracker.test/p)"),
      "!Illustration", "!\\[x\\](https://tracker.test/p)", "Thanks!",
    ]);
  });

  await t.test("literal text and alt backticks cannot open generated code delimiters", async () => {
    const pixel = "![p](https://tracker.test/p)";
    const fixtures = [
      `<p>\`<code>${pixel}</code></p>`,
      `<p><img alt="\`"><code>${pixel}</code></p>`,
      '<p>`<code>[a](javascript:b)</code></p>',
      '<p>`<code>&lt;img src=x&gt;</code></p>',
      `<code>outer \`<code>\`\`${pixel}\`\`</code> tail</code>`,
      `<pre><code>outer \`\`\`${pixel}\`\` tail</code></pre>`,
    ];
    const results = await evaluate(`return ${JSON.stringify(fixtures)}.map(htmlToMarkdown);`);
    assert.deepEqual(results, [
      "\\` ` " + pixel + " `", "\\` ` " + pixel + " `",
      "\\` ` [a](javascript:b) `", "\\` ` <img src=x> `",
      "```` outer ```" + pixel + "`` tail ````",
      "````\nouter ```" + pixel + "`` tail\n````",
    ]);
  });

  await t.test("raw HTML bodies remain unchanged, including untrusted markup", async () => {
    const html = '<script>raw-secret</script><img src="https://raw.test/image"><p hidden>hidden&#x200b;\u200b</p>';
    const result = await evaluate(`return sanitizeToolResultText(extractFormattedBody({contentType: "text/html", body: ${JSON.stringify(html)}}, "html"));`);
    assert.deepEqual(result, { body: html, bodyIsHtml: true });
  });

  await t.test("oversized HTML parses only the capped input and appends a truncation note", async () => {
    const result = await evaluate(`
      const Parser = DOMParser;
      const sizes = [];
      DOMParser = class {
        parseFromString(html, type) {
          sizes.push(html.length);
          if (html.length > 2 * 1024 * 1024) throw Error("Oversized body reached DOMParser");
          return new Parser().parseFromString(html, type);
        }
      };
      const normal = stripHtml("<p>Ordinary message</p>");
      const html = '<p>Newsletter start.</p>' +
        '<a href="https://destination.test/retained">Visible link</a><img src="https://tracking.test/pixel" alt="Illustration">' +
        '<div hidden>hidden-secret'.padEnd(2 * 1024 * 1024, 'x') + '</div><p>After the cut</p>';
      const text = stripHtml(html);
      const markdown = htmlToMarkdown(html);
      const largeText = stripHtml('<p>' + 'a-'.repeat(100000) + 'x'.repeat(2 * 1024 * 1024) + '</p>');
      return { normal, sizes, text, markdown, plain: {
        start: largeText.slice(0, 10), length: largeText.length,
        note: largeText.endsWith('[Message body truncated at 2 MiB]'),
      } };
    `);
    assert.equal(result.normal, "Ordinary message");
    assert.equal(result.sizes.length, 4);
    assert.deepEqual(result.sizes.slice(1), Array(3).fill(2 * 1024 * 1024));
    for (const text of [result.text, result.markdown]) {
      assert.match(text, /Newsletter start\./);
      assert.match(text, /Visible link/);
      assert.ok(text.endsWith("[Message body truncated at 2 MiB]"));
      assert.doesNotMatch(text, /secret|After the cut|withheld|tracking\.test/);
    }
    assert.ok(result.markdown.includes("[Visible link](https://destination.test/retained)"));
    assert.match(result.markdown, /Illustration/);
    assert.equal(result.plain.start, "a-a-a-a-a-");
    assert.ok(result.plain.length < 2 * 1024 * 1024 + 100);
    assert.equal(result.plain.note, true);
  });

  await t.test("CR entities cannot escape code blocks in lists or blockquotes", async () => {
    const pixel = "![pixel](https://tracker.test/p)";
    for (const newline of ["&#13;&#13;", "&#13;&#10;&#13;&#10;"]) {
      const results = await evaluate(`
        const pre = '<pre>safe' + ${JSON.stringify(newline)} + ${JSON.stringify(pixel)} + '</pre>';
        return [
          '<ul><li>' + pre + '</li></ul>',
          '<blockquote>' + pre + '</blockquote>',
          '<blockquote><ul><li>' + pre + '</li></ul></blockquote>',
          '<ul><li><blockquote>' + pre + '</blockquote></li></ul>',
        ].map(htmlToMarkdown);
      `);
      for (const [index, [firstLine, indent]] of [
        ["- ```", "  "], ["> ```", "> "], ["> - ```", ">   "], ["- > ```", "  > "],
      ].entries()) {
        assert.equal(results[index], `${firstLine}\n${indent}safe\n${indent}\n${indent}${pixel}\n${indent}\`\`\``);
        assert.doesNotMatch(results[index], /\r/);
      }
    }
  });

  await t.test("oversized HTML encryption classification uses the full source, never the parsed cut", async () => {
    const results = await evaluate(`
      const cap = 2 * 1024 * 1024;
      const armor = '-----BEGIN PGP MESSAGE-----';
      const encoded = '&#45;----BEGIN PGP MESSAGE-----';
      const fixtures = [
        '<p>' + ' '.repeat(cap) + armor + '</p>',
        '<p>' + ' '.repeat(cap - 3 - 12) + armor + ' in prose</p>',
        '<p>' + ' '.repeat(cap - 3 - encoded.length) + encoded + ' in prose</p>',
      ];
      return fixtures.map(html => {
        const mime = { contentType: 'text/html', body: html };
        const presentation = stripHtml(html);
        return {
          presentationArmor: hasInlinePgpArmor(presentation),
          originalArmor: hasInlinePgpBodyArmor(mime, presentation),
          encrypted: isEncryptedMimeMessage(mime),
        };
      });
    `);
    assert.deepEqual(results, [
      { presentationArmor: false, originalArmor: true, encrypted: true },
      { presentationArmor: false, originalArmor: true, encrypted: true },
      { presentationArmor: true, originalArmor: false, encrypted: false },
    ]);
  });

  await t.test("genuine parser failures still fail closed", async () => {
    const result = await evaluate(`
      DOMParser = class { parseFromString() { throw Error("Parser failed"); } };
      return [stripHtml("<p>private</p>"), htmlToMarkdown("<p>private</p>")];
    `);
    assert.deepEqual(result, Array(2).fill("[HTML content withheld: safe HTML parser unavailable.]"));
  });

  await t.test("ordinary and truncated HTML share Gecko's parsing and privacy rules", async () => {
    const fixtures = [
      '<script><!--<script></script>script-secret</script>',
      '<div hidden><template></div>template-secret</template></div>',
      '<div hidden><table></div>table-secret</table></div>',
      '<foo.bar hidden>name-secret</foo.bar>',
      '<foo_bar style="display: none ! important">name-secret</foo_bar>',
      '<script>before</script\u00a0>raw-secret</script>',
      '<script>before</script\v>raw-secret</script>',
      '<div x=a" style="display:none; >">quote-secret</div>',
      '<body hidden>before</body>body-secret',
      '<div style="color:red&semi;display:none">entity-secret</div>',
      '<div style="font:0px serif">font-secret</div>',
      '</ comment-secret>',
    ];
    const results = await evaluate(`
      const suffix = '<!--' + 'x'.repeat(2 * 1024 * 1024) + '-->';
      return ${JSON.stringify(fixtures)}.map(html =>
        [stripHtml(html), htmlToMarkdown(html), stripHtml(html + suffix), htmlToMarkdown(html + suffix)]);
    `);
    results.forEach((outputs, index) => {
      assert.deepEqual(outputs, ["", "", "[Message body truncated at 2 MiB]", "[Message body truncated at 2 MiB]"], fixtures[index]);
    });
  });
});
