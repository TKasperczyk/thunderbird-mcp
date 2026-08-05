"use strict";

const { describe, it } = require("node:test");
const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const vm = require("node:vm");

function loadDailyDigestHelpers() {
  const apiPath = path.resolve(__dirname, "../extension/mcp_server/api.js");
  const source = fs.readFileSync(apiPath, "utf8");
  const startMarker = "// BEGIN DAILY DIGEST HELPERS";
  const endMarker = "// END DAILY DIGEST HELPERS";
  const start = source.indexOf(startMarker);
  const end = source.indexOf(endMarker, start + startMarker.length);
  assert.ok(start >= 0, "daily digest helper start marker missing");
  assert.ok(end > start, "daily digest helper end marker missing");

  const sandbox = {};
  vm.createContext(sandbox);
  vm.runInContext(`
${source.slice(start, end)}
this.helpers = {
  normalizeDailyDigestText,
  parseDailyDigestDateRange,
  truncateDailyDigestText,
  classifyDailyDigestMessage,
  buildDailyDigestCounts,
  getTurkishTranslationInstructions,
};
`, sandbox);
  return sandbox.helpers;
}

const helpers = loadDailyDigestHelpers();

describe("daily digest date range", () => {
  it("parses a valid local calendar date", () => {
    const range = helpers.parseDailyDigestDateRange("2026-08-03");
    assert.equal(range.date, "2026-08-03");
    assert.ok(range.endMs > range.startMs);
    assert.ok(range.endMs - range.startMs >= 23 * 60 * 60 * 1000);
    assert.ok(range.endMs - range.startMs <= 25 * 60 * 60 * 1000);
  });

  it("rejects malformed and impossible dates", () => {
    assert.throws(
      () => helpers.parseDailyDigestDateRange("03-08-2026"),
      /YYYY-MM-DD/
    );
    assert.throws(
      () => helpers.parseDailyDigestDateRange("2026-02-30"),
      /valid calendar day/
    );
  });
});

describe("daily digest classification", () => {
  it("detects an overdue or failed invoice as urgent action", () => {
    const analysis = helpers.classifyDailyDigestMessage({
      subject: "Credit card charge failed",
      author: "billing@example.com",
      body: "Invoice 123 is overdue. Please pay 195.00 EUR by 2026-08-03.",
    });
    assert.ok(analysis.categories.includes("invoice"));
    assert.ok(analysis.categories.includes("payment"));
    assert.ok(analysis.categories.includes("action"));
    assert.equal(analysis.priority, "urgent");
    assert.equal(analysis.amount, "195.00 EUR");
    assert.equal(analysis.deadline, "2026-08-03");
  });

  it("detects GitHub token expiry", () => {
    const analysis = helpers.classifyDailyDigestMessage({
      subject: "[GitHub] Your personal access token is about to expire",
      author: "GitHub <noreply@github.com>",
      body: "The token needs attention.",
    });
    assert.ok(analysis.categories.includes("github"));
    assert.ok(analysis.categories.includes("payment") === false);
    assert.ok(analysis.categories.includes("action"));
    assert.equal(analysis.priority, "high");
  });

  it("detects security alerts", () => {
    const analysis = helpers.classifyDailyDigestMessage({
      subject: "New login to Instagram",
      author: "security@example.com",
      body: "A new device signed in.",
    });
    assert.ok(analysis.categories.includes("security"));
    assert.ok(analysis.actionRequired);
    assert.equal(analysis.priority, "urgent");
  });

  it("separates a completed payment from an unpaid invoice", () => {
    const paid = helpers.classifyDailyDigestMessage({
      subject: "Ödemen Başarıyla Alındı!",
      author: "iyzico <billing@example.com>",
      body: "1.250,00 TL ödemeniz için teşekkür ederiz.",
    });
    assert.ok(paid.categories.includes("payment"));
    assert.equal(paid.paymentStatus, "paid");
    assert.equal(paid.actionRequired, false);
    assert.equal(paid.primaryCategory, "payment");
    assert.ok(paid.suggestedActions.some(action => action.type === "record_expense"));
    assert.ok(!paid.suggestedActions.some(action => action.type === "review_payment"));
  });

  it("classifies banking, subscriptions, orders, and appointments independently", () => {
    const bank = helpers.classifyDailyDigestMessage({
      subject: "Para Transferi Bilgilendirmesi",
      author: "Garanti BBVA <notice@example.com>",
      body: "Hesap hareketiniz hakkında bilgilendirme.",
    });
    assert.ok(bank.categories.includes("banking"));

    const subscription = helpers.classifyDailyDigestMessage({
      subject: "Your annual subscription renewal",
      author: "service@example.com",
      body: "Your membership will auto-renew next month.",
    });
    assert.ok(subscription.categories.includes("subscription"));
    assert.ok(subscription.suggestedActions.some(action => action.type === "review_subscription"));

    const order = helpers.classifyDailyDigestMessage({
      subject: "Siparişiniz kargoya verildi",
      author: "shop@example.com",
      body: "Takip numarası ABC123.",
    });
    assert.ok(order.categories.includes("order"));
    assert.ok(order.suggestedActions.some(action => action.type === "track_order"));

    const appointment = helpers.classifyDailyDigestMessage({
      subject: "Randevu onayı",
      author: "clinic@example.com",
      body: "Randevunuzu onaylamak için yanıtlayın.",
    });
    assert.ok(appointment.categories.includes("appointment"));
    assert.ok(appointment.actionRequired);
    assert.ok(appointment.suggestedActions.some(action => action.type === "add_to_calendar"));
  });

  it("flags explicit spam language without treating normal newsletters as spam", () => {
    const spam = helpers.classifyDailyDigestMessage({
      subject: "Jackpot casino bonus",
      author: "promo@example.com",
      body: "Free money and crypto bonus.",
    });
    assert.ok(spam.categories.includes("spam"));
    assert.ok(spam.spamScore >= 3);

    const newsletter = helpers.classifyDailyDigestMessage({
      subject: "Ağustos bülteni",
      author: "news@example.com",
      body: "Yeni ürün haberleri ve abonelikten çık bağlantısı.",
    });
    assert.ok(newsletter.categories.includes("newsletter"));
    assert.ok(!newsletter.categories.includes("spam"));
  });

  it("builds overlapping category and priority totals", () => {
    const messages = [
      { analysis: { categories: ["invoice", "payment", "action"], priority: "urgent", actionRequired: true } },
      { analysis: { categories: ["newsletter"], priority: "normal", actionRequired: false } },
    ];
    const counts = helpers.buildDailyDigestCounts(messages);
    assert.equal(counts.categoryCounts.invoice, 1);
    assert.equal(counts.categoryCounts.payment, 1);
    assert.equal(counts.categoryCounts.newsletter, 1);
    assert.equal(counts.priorityCounts.urgent, 1);
    assert.equal(counts.priorityCounts.normal, 1);
    assert.equal(counts.actionRequired, 1);
  });

  it("counts primary categories and stable suggested action types", () => {
    const messages = [
      {
        analysis: {
          categories: ["invoice", "payment"],
          primaryCategory: "invoice",
          priority: "high",
          actionRequired: true,
          suggestedActions: [
            { type: "review_payment" },
            { type: "record_expense" },
          ],
        },
      },
    ];
    const counts = helpers.buildDailyDigestCounts(messages);
    assert.equal(counts.primaryCategoryCounts.invoice, 1);
    assert.equal(counts.suggestedActionCounts.review_payment, 1);
    assert.equal(counts.suggestedActionCounts.record_expense, 1);
  });
});

describe("Turkish translation preparation", () => {
  it("requires faithful formatting and preservation of identifiers", () => {
    const instructions = helpers.getTurkishTranslationInstructions("markdown");
    const text = instructions.join(" ");
    assert.match(text, /Türkçeye çevir/);
    assert.match(text, /URL'leri/);
    assert.match(text, /markdown/);
    assert.match(text, /bağlantı açma/);
  });

  it("truncates normalized body text at the requested bound", () => {
    const result = helpers.truncateDailyDigestText("a   b   c   d", 6);
    assert.equal(result.truncated, true);
    assert.ok(result.text.length <= 6);
    assert.match(result.text, /…$/);
  });
});
