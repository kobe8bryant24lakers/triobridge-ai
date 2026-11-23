import { test } from 'node:test';
import assert from 'node:assert/strict';
import { mkdtemp, readFile, readdir, rm, stat, symlink, utimes, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { createSharedModelStore, sharedCompletionUrl } from '../scripts/shared-model-store.mjs';

const profile = (id = 'first', extra = {}) => ({
  id,
  name: `Model ${id}`,
  baseUrl: 'https://example.test/v1',
  model: 'synthetic',
  timeoutSeconds: 5,
  ...extra,
});
const body = (content = '合成结果') => ({
  id: 'completion-1',
  model: 'actual-model',
  choices: [{ message: { content }, finish_reason: 'stop' }],
  usage: { prompt_tokens: 12, completion_tokens: 23, total_tokens: 35 },
});
const replace = (value = 'private-test-secret') => ({ action: 'replace', value });

async function fixture(t, options = {}) {
  const directory = await mkdtemp(join(tmpdir(), 'triobridge-shared-models-'));
  t.after(() => rm(directory, { recursive: true, force: true }));
  const keys = new Map();
  const operations = [];
  const vault = {
    kind: 'macos-keychain',
    async get(id) {
      operations.push(['get', id]);
      return keys.get(id) ?? null;
    },
    async set(id, value) {
      operations.push(['set', id]);
      keys.set(id, value);
    },
    async delete(id) {
      operations.push(['delete', id]);
      keys.delete(id);
    },
    ...options.vault,
  };
  const calls = [];
  const fetcher =
    options.fetcher ||
    (async (...args) => {
      calls.push(args);
      return new Response(JSON.stringify(body()), { headers: { 'x-request-id': 'request-1' } });
    });
  const store = createSharedModelStore({ directory, vault, fetcher });
  return { store, directory, keys, vault, calls, operations, fetcher };
}
const add = (store, id = 'first', expectedRevision = 0, credential = replace()) =>
  store.upsert({ expectedRevision, profile: profile(id), credential });

test('shared metadata persists without credentials and is readable by another store instance', async (t) => {
  const { store, directory, keys, vault, fetcher } = await fixture(t);
  assert.equal((await store.state()).revision, 0);
  const saved = await add(store);
  assert.equal(saved.profiles[0].hasKey, true);
  assert.equal(saved.profiles[0].revision, 1);
  assert.equal(saved.activeId, 'first');
  assert.equal(JSON.stringify(saved).includes('private-test-secret'), false);
  assert.equal(JSON.stringify(saved).includes('credentialRef'), false);
  const disk = await readFile(join(directory, 'profiles.json'), 'utf8');
  assert.equal(disk.includes('private-test-secret'), false);
  assert.equal(disk.includes('apiKey'), false);
  assert.equal(keys.size, 1);
  assert.equal((await stat(directory)).mode & 0o777, 0o700);
  assert.equal((await stat(join(directory, 'profiles.json'))).mode & 0o777, 0o600);
  const reopened = createSharedModelStore({ directory, vault, fetcher });
  assert.deepEqual(await reopened.state(), saved);
});

test('stale catalog writes fail with latest public state and cannot overwrite another Office', async (t) => {
  const { store, directory, vault, fetcher } = await fixture(t);
  await add(store);
  const second = createSharedModelStore({ directory, vault, fetcher });
  await second.upsert({
    expectedRevision: 1,
    profile: profile('first', { name: 'Changed elsewhere' }),
    credential: { action: 'keep' },
  });
  await assert.rejects(
    store.upsert({
      expectedRevision: 1,
      profile: profile('first', { name: 'stale' }),
      credential: replace('wrong-overwrite'),
    }),
    (error) => {
      assert.equal(error.code, 'conflict');
      assert.equal(error.status, 409);
      assert.equal(error.state.profiles[0].name, 'Changed elsewhere');
      assert.equal(JSON.stringify(error.state).includes('private-test-secret'), false);
      return true;
    },
  );
  assert.equal((await store.state()).profiles[0].name, 'Changed elsewhere');
});

test('switching the global default preserves each profile revision and deleting removes credentials', async (t) => {
  const { store, keys } = await fixture(t);
  await add(store);
  await add(store, 'second', 1, replace('second-secret'));
  const switched = await store.activate({ expectedRevision: 2, id: 'second' });
  assert.equal(switched.revision, 3);
  assert.equal(switched.activeId, 'second');
  assert.deepEqual(
    switched.profiles.map((p) => p.revision),
    [1, 2],
  );
  const removed = await store.delete({ expectedRevision: 3, id: 'second' });
  assert.equal(removed.activeId, 'first');
  assert.equal(removed.profiles.length, 1);
  assert.equal(keys.size, 1);
  await assert.rejects(
    store.complete({
      profileId: 'second',
      profileRevision: 2,
      messages: [{ role: 'user', content: 'hi' }],
    }),
    { code: 'missing' },
  );
});

test('changing a service cannot carry an existing key to the new endpoint', async (t) => {
  const { store, keys } = await fixture(t);
  await add(store);
  await assert.rejects(
    store.upsert({
      expectedRevision: 1,
      profile: profile('first', { baseUrl: 'https://different.test/v1' }),
      credential: { action: 'keep' },
    }),
    { code: 'credential' },
  );
  assert.equal((await store.state()).revision, 1);
  const saved = await store.upsert({
    expectedRevision: 1,
    profile: profile('first', { baseUrl: 'https://different.test/v1' }),
    credential: { action: 'remove' },
  });
  assert.equal(saved.profiles[0].hasKey, false);
  assert.equal(keys.size, 0);
});

test('completion uses only saved endpoint/model/key and filters provider evidence without secret echoes', async (t) => {
  let recorded;
  const { store } = await fixture(t, {
    fetcher: async (...args) => {
      recorded = args;
      return new Response(
        JSON.stringify({
          ...body('private-test-secret must be redacted'),
          Authorization: 'private-test-secret',
          usage: {
            total_tokens: 35,
            apiKey: 'private-test-secret',
            completion_tokens_details: { reasoning_tokens: 4, secret: 'private-test-secret' },
          },
        }),
        { headers: { 'x-request-id': 'request-id', 'x-secret': 'private-test-secret' } },
      );
    },
  });
  await add(store);
  const result = await store.complete({
    profileId: 'first',
    profileRevision: 1,
    messages: [{ role: 'user', content: 'synthetic input' }],
  });
  assert.equal(recorded[0], 'https://example.test/v1/chat/completions');
  assert.equal(recorded[1].headers.Authorization, 'Bearer private-test-secret');
  assert.equal(recorded[1].redirect, 'error');
  assert.deepEqual(JSON.parse(recorded[1].body), {
    model: 'synthetic',
    messages: [{ role: 'user', content: 'synthetic input' }],
    stream: false,
  });
  assert.equal(result.requestId, 'request-id');
  assert.equal(result.body.model, 'actual-model');
  assert.equal(result.body.choices[0].message.content, '[REDACTED] must be redacted');
  assert.deepEqual(result.body.usage, {
    total_tokens: 35,
    completion_tokens_details: { reasoning_tokens: 4 },
  });
  assert.equal(JSON.stringify(result).includes('private-test-secret'), false);
  assert.equal(JSON.stringify(result).includes('Authorization'), false);
  await assert.rejects(
    store.complete({
      profileId: 'first',
      profileRevision: 0,
      messages: [{ role: 'user', content: 'stale' }],
    }),
    { code: 'conflict', status: 409 },
  );
});

test('atomic batch import rolls back vault entries on partial vault failure', async (t) => {
  const { store, keys, vault, directory } = await fixture(t);
  let writes = 0;
  vault.set = async (id, value) => {
    if (++writes === 2) throw new Error('locked');
    keys.set(id, value);
  };
  await assert.rejects(
    store.import({
      expectedRevision: 0,
      profiles: [
        { profile: profile('first'), credential: replace('first-private') },
        { profile: profile('second'), credential: replace('second-private') },
      ],
    }),
    { code: 'credential' },
  );
  assert.equal((await store.state()).revision, 0);
  assert.deepEqual((await store.state()).profiles, []);
  assert.equal(keys.size, 0);
  const disk = await readFile(join(directory, 'profiles.json'), 'utf8');
  assert.equal(disk.includes('first-private'), false);
  assert.equal(disk.includes('second-private'), false);
});

test('atomic batch import validates all records, preserves existing models and activates the requested default', async (t) => {
  const { store, keys } = await fixture(t);
  await add(store);
  await assert.rejects(
    store.import({
      expectedRevision: 1,
      profiles: [
        { profile: profile('valid'), credential: replace('new-secret') },
        { profile: profile('invalid', { baseUrl: 'http://remote.test' }), credential: replace() },
      ],
    }),
    { code: 'config' },
  );
  assert.equal((await store.state()).profiles.length, 1);
  assert.equal(keys.size, 1);
  const result = await store.import({
    expectedRevision: 1,
    activeId: 'second',
    profiles: [
      { profile: profile('second'), credential: replace('new-secret') },
      { profile: profile('third'), credential: { action: 'remove' } },
    ],
  });
  assert.equal(result.revision, 2);
  assert.equal(result.activeId, 'second');
  assert.deepEqual(
    result.profiles.map((p) => p.hasKey),
    [true, true, false],
  );
  await assert.rejects(
    store.import({
      expectedRevision: 2,
      profiles: [{ profile: profile('first'), credential: replace('overwrite') }],
    }),
    { code: 'conflict' },
  );
});

test('connection test failure and cancellation do not commit configuration or credentials', async (t) => {
  const { store, keys } = await fixture(t, {
    fetcher: async () => new Response('provider might echo private-test-secret', { status: 401 }),
  });
  await assert.rejects(
    store.upsert({
      expectedRevision: 0,
      profile: profile(),
      credential: replace(),
      testConnection: true,
    }),
    { code: 'upstream', status: 401 },
  );
  assert.equal((await store.state()).revision, 0);
  assert.equal(keys.size, 0);
  const cancelled = new AbortController();
  cancelled.abort();
  await assert.rejects(
    store.upsert(
      { expectedRevision: 0, profile: profile(), credential: replace() },
      cancelled.signal,
    ),
    { code: 'cancelled' },
  );
  assert.equal(keys.size, 0);
});

test('connection test rechecks the revision after another Office changes the catalog', async (t) => {
  let finish;
  let started;
  const waiting = new Promise((resolve) => {
    started = resolve;
  });
  const { store, keys } = await fixture(t, {
    fetcher: async () => {
      started();
      return new Promise((resolve) => {
        finish = resolve;
      });
    },
  });
  const pending = store.upsert({
    expectedRevision: 0,
    profile: profile('tested'),
    credential: replace('tested-secret'),
    testConnection: true,
  });
  await waiting;
  await add(store, 'other', 0, replace('other-secret'));
  finish(new Response(JSON.stringify(body('OK'))));
  await assert.rejects(pending, { code: 'conflict' });
  assert.deepEqual(
    (await store.state()).profiles.map((p) => p.id),
    ['other'],
  );
  assert.equal(keys.size, 1);
});

test('disconnect aborts an upstream request and does not persist document contents', async (t) => {
  let started;
  const waiting = new Promise((resolve) => {
    started = resolve;
  });
  const { store, directory } = await fixture(t, {
    fetcher: async (_url, init) => {
      started();
      return new Promise((_resolve, reject) =>
        init.signal.addEventListener('abort', () => reject(new Error('aborted'))),
      );
    },
  });
  await add(store);
  const controller = new AbortController();
  const pending = store.complete(
    {
      profileId: 'first',
      profileRevision: 1,
      messages: [{ role: 'user', content: 'synthetic-private-document' }],
    },
    controller.signal,
  );
  await waiting;
  controller.abort();
  await assert.rejects(pending, { code: 'cancelled' });
  assert.equal(
    (await readFile(join(directory, 'profiles.json'), 'utf8')).includes(
      'synthetic-private-document',
    ),
    false,
  );
});

test('deleted keys awaiting vault recovery are reported and cleanup retries without restoring profiles', async (t) => {
  const { store, vault, keys } = await fixture(t);
  await add(store);
  vault.delete = async () => {
    throw new Error('vault locked');
  };
  const deleted = await store.delete({ expectedRevision: 1, id: 'first' });
  assert.deepEqual(deleted.profiles, []);
  assert.match(deleted.warning, /旧密钥/);
  assert.equal(keys.size, 1);
  vault.delete = async (id) => keys.delete(id);
  const cleaned = await store.cleanup();
  assert.equal(cleaned.warning, undefined);
  assert.equal(keys.size, 0);
  assert.equal(cleaned.revision, 2);
});

test('corrupt and linked metadata are not overwritten, and an abandoned lock can recover', async (t) => {
  const { store, directory } = await fixture(t);
  const file = join(directory, 'profiles.json');
  await writeFile(file, '{corrupt');
  await assert.rejects(add(store), { code: 'storage' });
  assert.equal(await readFile(file, 'utf8'), '{corrupt');
  await rm(file);
  const target = join(directory, 'outside.json');
  await writeFile(target, 'do not touch');
  await symlink(target, file);
  await assert.rejects(store.state(), { code: 'storage' });
  assert.equal(await readFile(target, 'utf8'), 'do not touch');
  await rm(file);
  const lock = join(directory, 'profiles.lock');
  await writeFile(lock, 'incomplete old lock');
  await utimes(lock, new Date(0), new Date(0));
  await add(store);
  assert.equal((await store.state()).revision, 1);
  assert.equal(
    (await readdir(directory)).some((name) => name.endsWith('.tmp') || name.endsWith('.lock')),
    false,
  );
});

test('shared endpoints accept only loopback plaintext and a locally fixed proxy target', () => {
  for (const url of [
    'http://example.test',
    'https://user:secret@example.test',
    'https://example.test?key=x',
    'https://example.test#key',
    'file:///tmp/key',
  ])
    assert.throws(() => sharedCompletionUrl(url));
  assert.equal(
    sharedCompletionUrl('https://example.test/v1/chat/completions/'),
    'https://example.test/v1/chat/completions',
  );
  assert.equal(
    sharedCompletionUrl('http://localhost:11434/v1'),
    'http://localhost:11434/v1/chat/completions',
  );
  assert.equal(
    sharedCompletionUrl('http://127.0.0.1:11434/v1'),
    'http://127.0.0.1:11434/v1/chat/completions',
  );
  assert.equal(
    sharedCompletionUrl('http://[::1]:11434/v1'),
    'http://[::1]:11434/v1/chat/completions',
  );
  assert.equal(
    sharedCompletionUrl('/model-proxy/v1', 'https://provider.test'),
    'https://provider.test/v1/chat/completions',
  );
  assert.equal(
    sharedCompletionUrl('/model-proxy/v1', 'http://127.0.0.1:11434'),
    'http://127.0.0.1:11434/v1/chat/completions',
  );
  for (const target of [
    undefined,
    'http://remote.test',
    'https://provider.test/path',
    'https://user:key@provider.test',
  ])
    assert.throws(() => sharedCompletionUrl('/model-proxy/v1', target));
});

test('saved proxy destinations remain bound when the server target changes and credentials cannot silently follow', async (t) => {
  const { directory, vault, keys } = await fixture(t);
  const calls = [];
  const fetcher = async (url) => {
    calls.push(url);
    return new Response(JSON.stringify(body()));
  };
  const original = createSharedModelStore({
    directory,
    vault,
    fetcher,
    proxyTarget: 'https://first.test',
  });
  await original.upsert({
    expectedRevision: 0,
    profile: profile('proxy', { baseUrl: '/model-proxy/v1' }),
    credential: replace(),
  });
  const restarted = createSharedModelStore({
    directory,
    vault,
    fetcher,
    proxyTarget: 'https://different.test',
  });
  await restarted.complete({
    profileId: 'proxy',
    profileRevision: 1,
    messages: [{ role: 'user', content: 'synthetic' }],
  });
  assert.deepEqual(calls, ['https://first.test/v1/chat/completions']);
  await assert.rejects(
    restarted.upsert({
      expectedRevision: 1,
      profile: profile('proxy', { baseUrl: '/model-proxy/v1' }),
      credential: { action: 'keep' },
    }),
    { code: 'credential' },
  );
  assert.equal(keys.size, 1);
});

test('concurrent Office mutations cannot both commit the same starting revision', async (t) => {
  const { store, directory, vault, keys, fetcher } = await fixture(t);
  const second = createSharedModelStore({ directory, vault, fetcher });
  const results = await Promise.allSettled([add(store, 'one'), add(second, 'two')]);
  assert.equal(results.filter((result) => result.status === 'fulfilled').length, 1);
  const rejected = results.find((result) => result.status === 'rejected');
  assert.equal(rejected.reason.status, 409);
  assert.equal((await store.state()).revision, 1);
  assert.equal((await store.state()).profiles.length, 1);
  assert.equal(keys.size, 1);
});

test('cancellation after vault staging removes the staged credential and preserves the old catalog', async (t) => {
  const { store, vault, keys } = await fixture(t);
  const controller = new AbortController();
  vault.set = async (id, value) => {
    keys.set(id, value);
    controller.abort();
  };
  await assert.rejects(
    store.upsert(
      { expectedRevision: 0, profile: profile(), credential: replace() },
      controller.signal,
    ),
    { code: 'cancelled' },
  );
  assert.equal((await store.state()).revision, 0);
  assert.deepEqual((await store.state()).profiles, []);
  assert.equal(keys.size, 0);
});

test('upstream timeout aborts the request and distinguishes cancellation from timeout', async (t) => {
  let wasAborted = false;
  const { store } = await fixture(t, {
    fetcher: async (_url, init) =>
      new Promise((_resolve, reject) => {
        init.signal.addEventListener('abort', () => {
          wasAborted = true;
          reject(new Error('aborted'));
        });
      }),
  });
  await add(store, 'timeout', 0, { action: 'remove' });
  await assert.rejects(
    store.complete({
      profileId: 'timeout',
      profileRevision: 1,
      messages: [{ role: 'user', content: 'timeout fixture' }],
    }),
    { code: 'timeout', status: 504 },
  );
  assert.equal(wasAborted, true);
});
