#!/usr/bin/env node
/**
 * Seed a greenmail instance with a fixed set of test messages.
 *
 * Uses raw SMTP over net.Socket (zero deps) to push messages through
 * greenmail's default SMTP listener on localhost:3025. Each message
 * gets From / To / Subject / Date / Body, and the IMAP account
 * picks them up when TB syncs the folder.
 *
 * The fixture set is deliberately varied so integration tests can
 * exercise:
 *   - inbox_inventory from_domain grouping (3 distinct domains)
 *   - bulk_move_by_query query matching (linkedin.com prefix)
 *   - searchMessages token matching
 *   - applyFilters matching (subjects with "Invoice" / "Newsletter")
 *
 * Usage: node scripts/ci/seed-greenmail.cjs [count=20]
 */

const net = require("net");

const HOST = process.env.GREENMAIL_HOST || "localhost";
const PORT = parseInt(process.env.GREENMAIL_SMTP_PORT || "3025", 10);
const TO = process.env.TEST_MAILBOX || "test@ci.local";

function crlf(s) { return s.replace(/\r?\n/g, "\r\n"); }

// Fixture generator. Alternates senders, subjects and dates so the
// grouping/filtering tools have interesting aggregation data.
function buildFixtures(count) {
  const senders = [
    { domain: "linkedin.com", name: "LinkedIn Notifications", email: "notifications@linkedin.com" },
    { domain: "linkedin.com", name: "LinkedIn Jobs",          email: "jobs-noreply@linkedin.com" },
    { domain: "github.com",   name: "GitHub",                  email: "noreply@github.com" },
    { domain: "stripe.com",   name: "Stripe",                  email: "invoicing@stripe.com" },
    { domain: "newsletter.example.com", name: "Weekly Digest", email: "weekly@newsletter.example.com" },
  ];
  const subjects = [
    "New connection request from Alice",
    "5 jobs matching your profile",
    "Your pull request was merged",
    "Invoice #42 is ready",
    "Newsletter: This week at Example",
    "Security alert: new sign-in",
    "Payment received: €150",
  ];
  const messages = [];
  for (let i = 0; i < count; i++) {
    const s = senders[i % senders.length];
    const subj = subjects[i % subjects.length];
    const date = new Date(Date.now() - i * 3600 * 1000).toUTCString();
    const body =
      `This is fixture message #${i + 1} seeded by scripts/ci/seed-greenmail.cjs ` +
      `for Thunderbird MCP integration tests. Sender: ${s.email}. Seq: ${i}.`;
    messages.push({
      from: `"${s.name}" <${s.email}>`,
      to: TO,
      subject: `${subj} [#${i + 1}]`,
      date,
      body,
      domain: s.domain,
    });
  }
  return messages;
}

function sendOne(msg) {
  return new Promise((resolve, reject) => {
    const socket = net.createConnection({ host: HOST, port: PORT }, () => {});
    let buf = "";
    let step = 0;
    // Minimal SMTP dialog. Greenmail accepts any MAIL FROM / RCPT TO.
    const steps = [
      () => `EHLO ci-seeder\r\n`,
      () => `MAIL FROM:<seeder@ci.local>\r\n`,
      () => `RCPT TO:<${msg.to}>\r\n`,
      () => `DATA\r\n`,
      () => crlf(
        `From: ${msg.from}\n` +
        `To: ${msg.to}\n` +
        `Subject: ${msg.subject}\n` +
        `Date: ${msg.date}\n` +
        `Message-ID: <seed-${Date.now()}-${Math.random().toString(36).slice(2)}@ci.local>\n` +
        `MIME-Version: 1.0\n` +
        `Content-Type: text/plain; charset=utf-8\n` +
        `\n` +
        `${msg.body}\n` +
        `.\n`
      ),
      () => `QUIT\r\n`,
    ];
    const expect = [/^220 /, /^250-/, /^250 /, /^250 /, /^354 /, /^250 /, /^221 /];

    socket.on("data", (chunk) => {
      buf += chunk.toString("utf8");
      // Process line-by-line; greenmail may send multi-line responses.
      while (true) {
        const nl = buf.indexOf("\r\n");
        if (nl === -1) break;
        const line = buf.slice(0, nl);
        buf = buf.slice(nl + 2);
        const re = expect[step];
        if (!re.test(line)) {
          socket.destroy();
          return reject(new Error(`SMTP step ${step} unexpected: ${line}`));
        }
        // Skip continuation lines of a 250-EHLO response
        if (step === 1 && line.startsWith("250-")) continue;
        step++;
        if (step === expect.length) { socket.end(); return resolve(); }
        const cmd = steps[step - 1]();
        socket.write(cmd);
      }
    });
    socket.on("error", reject);
    socket.on("end", () => {
      if (step < expect.length) reject(new Error(`SMTP closed early at step ${step}`));
    });
  });
}

(async () => {
  const count = parseInt(process.argv[2] || process.env.SEED_COUNT || "20", 10);
  const fixtures = buildFixtures(count);
  console.log(`Seeding ${fixtures.length} messages to ${TO} via ${HOST}:${PORT}`);
  let ok = 0;
  for (const f of fixtures) {
    try { await sendOne(f); ok++; }
    catch (e) { console.error(`  FAIL "${f.subject}": ${e.message}`); }
  }
  console.log(`Seeded ${ok}/${fixtures.length} messages`);
  if (ok === 0) process.exit(1);
})();
