"use strict";

const { describe, it } = require("node:test");
const assert = require("node:assert/strict");

const {
  categoryRules,
  classify,
  managedCategoryKeys,
  primaryRuleOrder,
} = require("../scripts/categorize-mail.cjs");

function message(subject, author = "sender@example.com", preview = "") {
  return { subject, author, preview, recipients: "", ccList: "" };
}

describe("simple Thunderbird mail categories", () => {
  it("defines seven primary categories plus the optional action overlay", () => {
    assert.equal(categoryRules.length, 8);
    assert.equal(managedCategoryKeys.size, 8);
    assert.deepEqual(
      primaryRuleOrder.map(rule => rule.name),
      [
        "Bankalar",
        "Faturalar",
        "Siparişler",
        "Randevular",
        "Ödemeler",
        "Bültenler",
        "Teknik",
      ]
    );
  });

  it("uses invoice rather than stacking invoice, payment, and technical tags", () => {
    const tags = classify(
      message(
        "Hetzner Invoice 123 is overdue",
        "noreply.billing@hetzner.com",
        "Payment is required before the due date."
      ),
      "Inbox"
    );
    assert.deepEqual([...tags].sort(), ["$label1", "$label4"]);
  });

  it("prefers order and appointment categories over generic payment language", () => {
    assert.deepEqual(
      [...classify(message("Siparişiniz için ödeme alındı, kargoya verildi"), "Inbox")],
      ["$label5"]
    );
    assert.deepEqual(
      [...classify(message("Randevunuzu onaylayın"), "Inbox")].sort(),
      ["$label1", "randevular"]
    );
  });

  it("keeps newsletters non-urgent and recognizes a Newsletter folder", () => {
    assert.deepEqual(
      [...classify(message("Ağustos ürün haberleri"), "Newsletter")],
      ["bultenler"]
    );
  });

  it("never returns more than one primary category plus Action", () => {
    const tags = classify(
      message(
        "Urgent invoice payment for order and server renewal",
        "billing@github.com",
        "Confirm the appointment and track shipment."
      ),
      "Newsletter"
    );
    assert.ok(tags.size <= 2);
    assert.ok(tags.has("$label1"));
  });
});
