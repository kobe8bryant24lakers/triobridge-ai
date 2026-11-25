import { readFile } from 'node:fs/promises';
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { runInNewContext } from 'node:vm';
import { randomUUID } from 'node:crypto';

// Exercise the exact script embedded in the compiled WinForms control. The only
// C# interpolation is the validated host enum; no parallel bridge implementation.
const source = await readFile(new URL('../offline/windows/TaskPane.cs', import.meta.url), 'utf8');
const literal = source.match(/return @"(\(\(\) => \{[\s\S]+?\}\)\(\));";/)?.[1];
assert.ok(literal, 'The embedded page bridge must remain available for contract tests.');
const script = literal.replace('" + hostName + @"', 'word').replaceAll('""', '"') + ';';

function page({
  url = 'https://triobridge.local/offline.html',
  frame = false,
  sendError = false,
} = {}) {
  const sent = [];
  let receive;
  const window = {
    chrome: {
      webview: {
        postMessage(message) {
          if (sendError) throw Error('closed channel');
          sent.push(message);
        },
        addEventListener(name, listener) {
          assert.equal(name, 'message');
          receive = listener;
        },
      },
    },
  };
  window.top = frame ? {} : window;
  runInNewContext(script, { window, location: { href: url }, crypto: { randomUUID } });
  return { bridge: window.trioWorkbench, sent, reply: (data) => receive({ data }), window };
}

test('pane exposes the native workbench API only in its exact top-level page', () => {
  for (const args of [
    { frame: true },
    { url: 'https://triobridge.local/other.html' },
    { url: 'https://triobridge.local/offline.html?source=remote' },
    { url: 'https://triobridge.local.evil/offline.html' },
    { url: 'file:///offline.html' },
  ])
    assert.equal(page(args).bridge, undefined);
  const current = page();
  assert.equal(current.bridge.host, 'word');
  assert.deepEqual(Object.keys(current.bridge), ['host', 'model', 'cancel', 'office']);
  assert.ok(Object.isFrozen(current.bridge));
  assert.equal(
    Object.getOwnPropertyDescriptor(current.window, 'trioWorkbench').configurable,
    false,
  );
});

test('pane correlates model and Office replies and preserves complete large output', async () => {
  const pane = page();
  const model = pane.bridge.model('request', 'complete', '{}');
  const office = pane.bridge.office('write', { token: 'selection', text: 'new content' });
  const [first, second] = pane.sent;
  assert.notEqual(first.id, second.id);
  assert.equal(first.method, 'model');
  assert.equal(first.requestId, 'request');
  assert.equal(second.method, 'office');
  const text = '中文𠮷<完整正文>\n'.repeat(180_000);
  pane.reply({ id: second.id, ok: true, value: { ok: true, value: null } });
  pane.reply({ id: 'unknown', ok: true, value: 'ignored' });
  pane.reply({ id: first.id, ok: true, value: { status: 200, body: text } });
  assert.equal((await model).body, text);
  assert.equal((await office).ok, true);
});

test('pane forwards cancellation and rejects all outstanding calls on disconnect', async () => {
  const pane = page();
  const first = pane.bridge.model('request', 'state', '{}');
  const second = pane.bridge.office('selection');
  const expectedFirst = assert.rejects(first, /worker closed/);
  const expectedSecond = assert.rejects(second, /worker closed/);
  pane.bridge.cancel('request');
  assert.equal(pane.sent[2].method, 'cancel');
  assert.equal(pane.sent[2].requestId, 'request');
  pane.reply({ type: 'disconnect', message: 'worker closed' });
  await Promise.all([expectedFirst, expectedSecond]);
  await assert.rejects(pane.bridge.office('selection'), /worker closed/);
  pane.bridge.cancel('request');
  assert.equal(pane.sent.length, 3);
});

test('pane reports individual failures and synchronous channel failures', async () => {
  const pane = page();
  const operation = pane.bridge.office('write');
  const expected = assert.rejects(operation, /selection changed/);
  pane.reply({ id: pane.sent[0].id, ok: false, message: 'selection changed' });
  await expected;
  await assert.rejects(
    page({ sendError: true }).bridge.model('id', 'state', '{}'),
    /closed channel/,
  );
});
