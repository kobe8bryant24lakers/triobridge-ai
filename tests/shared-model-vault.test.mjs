import { test } from 'node:test';
import assert from 'node:assert/strict';
import { createNativeModelVault } from '../scripts/shared-model-vault.mjs';

test('native credential writes finish before catalog cancellation can roll them back', async () => {
  let finish;
  let begin;
  const started = new Promise((resolve) => {
    begin = resolve;
  });
  let nativeArguments;
  class AsyncEntry {
    setPassword(...args) {
      nativeArguments = args;
      begin();
      return new Promise((resolve) => {
        finish = resolve;
      });
    }
  }
  const vault = createNativeModelVault({
    platform: 'darwin',
    loadBinding: async () => ({ AsyncEntry }),
  });
  const controller = new AbortController();
  let completed = false;
  const pending = vault
    .set('synthetic-reference', 'synthetic-value', controller.signal)
    .then(() => {
      completed = true;
    });
  await started;
  controller.abort();
  await Promise.resolve();
  assert.equal(completed, false);
  assert.deepEqual(nativeArguments, ['synthetic-value']);
  finish();
  await pending;
  assert.equal(completed, true);
});

test('an already-cancelled save never starts a native credential operation', async () => {
  let loaded = false;
  const vault = createNativeModelVault({
    platform: 'win32',
    loadBinding: async () => {
      loaded = true;
      throw new Error('must not load');
    },
  });
  const controller = new AbortController();
  controller.abort();
  await assert.rejects(vault.set('synthetic-reference', 'synthetic-value', controller.signal), {
    name: 'AbortError',
  });
  assert.equal(loaded, false);
});

test('unsupported OS credential storage fails closed with no filesystem fallback', () => {
  assert.throws(() => createNativeModelVault({ platform: 'linux' }));
});
