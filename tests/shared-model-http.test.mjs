import { test } from 'node:test';
import assert from 'node:assert/strict';
import { createServer, request as httpRequest } from 'node:http';
import { createServer as createHttp2Server, connect } from 'node:http2';
import { once } from 'node:events';
import { mkdtemp, rm } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { createSharedModelStore } from '../scripts/shared-model-store.mjs';
import { createSharedModelMiddleware } from '../scripts/shared-model-http.mjs';
const authToken = 'a'.repeat(64);

const profile = {
  id: 'test-model',
  name: '合成共享模型',
  baseUrl: 'https://example.test/v1',
  model: 'configured-model',
  timeoutSeconds: 5,
};

async function fixture(
  t,
  fetcher = async () =>
    new Response(
      JSON.stringify({
        id: 'test-id',
        model: 'provider-model',
        choices: [{ message: { content: 'synthetic answer' }, finish_reason: 'stop' }],
        usage: { total_tokens: 42 },
      }),
      { headers: { 'x-request-id': 'provider-request-id', 'x-private': 'do-not-return' } },
    ),
) {
  const directory = await mkdtemp(join(tmpdir(), 'triobridge-shared-http-'));
  const keys = new Map();
  const store = createSharedModelStore({
    directory,
    fetcher,
    vault: {
      kind: 'macos-keychain',
      async set(id, key) {
        keys.set(id, key);
      },
      async get(id) {
        return keys.get(id) ?? null;
      },
      async delete(id) {
        keys.delete(id);
      },
    },
  });
  let handler;
  // HTTP is only a test transport. The application mounts this middleware in its HTTPS server.
  const server = createServer((req, res) =>
    handler(req, res, () => {
      res.writeHead(404);
      res.end();
    }),
  );
  server.listen(0, '127.0.0.1');
  await once(server, 'listening');
  const port = server.address().port;
  const origin = `https://localhost:${port}`;
  handler = createSharedModelMiddleware({ store, allowedOrigin: origin, authToken });
  t.after(async () => {
    server.closeAllConnections();
    await new Promise((resolve) => server.close(resolve));
    await rm(directory, { recursive: true, force: true });
  });
  const request = (action = 'state', input = {}, options = {}) =>
    fetch(`http://127.0.0.1:${port}/local-models/${action}`, {
      method: 'POST',
      headers: {
        Origin: origin,
        'Content-Type': 'application/json',
        'X-TrioBridge-Local-Token': authToken,
      },
      body: JSON.stringify(input),
      ...options,
    });
  return { store, keys, request, origin, port };
}

test('same-origin shared API returns metadata and completion evidence without credentials', async (t) => {
  const { request } = await fixture(t);
  assert.equal((await request()).status, 200);
  const added = await request('upsert', {
    expectedRevision: 0,
    profile,
    credential: { action: 'replace', value: 'private-http-secret' },
    activate: true,
  });
  assert.equal(added.status, 200);
  const state = await added.json();
  assert.equal(state.profiles[0].hasKey, true);
  assert.equal(state.activeId, profile.id);
  assert.equal(JSON.stringify(state).includes('private-http-secret'), false);
  const response = await request('complete', {
    profileId: profile.id,
    profileRevision: state.profiles[0].revision,
    messages: [{ role: 'user', content: '合成输入' }],
  });
  assert.equal(response.status, 200);
  assert.equal(response.headers.get('x-request-id'), 'provider-request-id');
  assert.equal(response.headers.get('x-private'), null);
  assert.equal(response.headers.get('cache-control'), 'no-store');
  assert.equal(response.headers.get('access-control-allow-origin'), null);
  const data = await response.json();
  assert.equal(data.model, 'provider-model');
  assert.equal(data.usage.total_tokens, 42);
  assert.equal(data.localModel.endpoint, 'https://example.test/v1/chat/completions');
  assert.equal(JSON.stringify(data).includes('private-http-secret'), false);
});

test('API rejects cross-origin, missing origin, rebinding, query routes and non-JSON access', async (t) => {
  const { request, origin, port, store } = await fixture(t);
  for (const headers of [
    { Origin: 'https://external.test', 'Content-Type': 'application/json' },
    { 'Content-Type': 'application/json' },
  ])
    assert.equal((await request('state', {}, { headers })).status, 403);
  assert.equal((await request('state', {}, { method: 'GET', body: undefined })).status, 405);
  assert.equal(
    (
      await request(
        'state',
        {},
        {
          headers: {
            Origin: origin,
            'Content-Type': 'text/plain',
            'X-TrioBridge-Local-Token': authToken,
          },
        },
      )
    ).status,
    405,
  );
  assert.equal((await request('state?target=external')).status, 404);
  assert.equal((await request('state/extra')).status, 404);
  const rebinding = await new Promise((resolve, reject) => {
    const req = httpRequest(
      `http://127.0.0.1:${port}/local-models/state`,
      {
        method: 'POST',
        headers: { Host: 'evil.test', Origin: origin, 'Content-Type': 'application/json' },
      },
      (res) => {
        res.resume();
        res.on('end', () => resolve(res.statusCode));
      },
    );
    req.on('error', reject);
    req.end('{}');
  });
  assert.equal(rebinding, 403);
  assert.deepEqual((await store.state()).profiles, []);
});

test('API rejects arbitrary completion URL/key overrides, oversize and malformed requests', async (t) => {
  const { request, keys } = await fixture(t);
  for (const property of ['url', 'baseUrl', 'apiKey', 'model', 'headers']) {
    const response = await request('complete', {
      profileId: profile.id,
      profileRevision: 1,
      messages: [{ role: 'user', content: 'hi' }],
      [property]: 'attacker-controlled',
    });
    assert.equal(response.status, 400);
  }
  assert.equal((await request('state', {}, { body: 'not JSON' })).status, 400);
  assert.equal((await request('state', {}, { body: 'x'.repeat(512_001) })).status, 413);
  assert.equal(keys.size, 0);
});

test('API exposes optimistic conflicts as 409 and preserves catalog state', async (t) => {
  const { request } = await fixture(t);
  await request('upsert', { expectedRevision: 0, profile, credential: { action: 'remove' } });
  const response = await request('delete', { expectedRevision: 0, id: profile.id });
  assert.equal(response.status, 409);
  const data = await response.json();
  assert.equal(data.error.code, 'conflict');
  assert.equal(data.state.profiles[0].id, profile.id);
  assert.equal(data.state.revision, 1);
});

test('all shared endpoints reject missing or wrong installation capabilities without revealing the token', async (t) => {
  const { request, origin } = await fixture(t);
  for (const action of ['state', 'upsert', 'import', 'delete', 'activate', 'complete']) {
    for (const token of [undefined, 'b'.repeat(64), 'short']) {
      const response = await request(
        action,
        {},
        {
          headers: {
            Origin: origin,
            'Content-Type': 'application/json',
            ...(token ? { 'X-TrioBridge-Local-Token': token } : {}),
          },
        },
      );
      assert.equal(response.status, 401);
      const output = await response.text();
      assert.equal(output.includes(authToken), false);
      assert.equal(JSON.parse(output).error.code, 'local_auth');
    }
  }
  const output = await (await request()).text();
  assert.equal(output.includes(authToken), false);
});

test('client disconnection reaches the actual in-flight upstream abort signal', async (t) => {
  let entered;
  let aborted;
  const started = new Promise((resolve) => {
    entered = resolve;
  });
  const stopped = new Promise((resolve) => {
    aborted = resolve;
  });
  const { request } = await fixture(t, async (_url, init) => {
    entered();
    return new Promise((_resolve, reject) =>
      init.signal.addEventListener('abort', () => {
        aborted();
        reject(new Error('aborted'));
      }),
    );
  });
  await request('upsert', { expectedRevision: 0, profile, credential: { action: 'remove' } });
  const controller = new AbortController();
  const pending = request(
    'complete',
    {
      profileId: profile.id,
      profileRevision: 1,
      messages: [{ role: 'user', content: 'slow request' }],
    },
    { signal: controller.signal },
  );
  await started;
  controller.abort();
  await assert.rejects(pending);
  await stopped;
});

test('HTTP/2 native Office authority works while conflicting Host stays blocked', async (t) => {
  const { store } = await fixture(t);
  let handler;
  const server = createHttp2Server((req, res) => handler(req, res, () => res.end()));
  server.listen(0, '127.0.0.1');
  await once(server, 'listening');
  const port = server.address().port;
  const origin = `https://localhost:${port}`;
  handler = createSharedModelMiddleware({ store, allowedOrigin: origin, authToken });
  const client = connect(`http://127.0.0.1:${port}`);
  t.after(async () => {
    client.close();
    await new Promise((resolve) => server.close(resolve));
  });
  const request = (headers = {}) =>
    new Promise((resolve, reject) => {
      const stream = client.request({
        ':method': 'POST',
        ':path': '/local-models/state',
        ':authority': `localhost:${port}`,
        origin,
        'content-type': 'application/json',
        'x-triobridge-local-token': authToken,
        ...headers,
      });
      let status;
      stream.on('response', (value) => {
        status = value[':status'];
        stream.resume();
      });
      stream.on('error', reject);
      stream.on('end', () => resolve(status));
      stream.end('{}');
    });
  assert.equal(await request(), 200);
  assert.equal(await request({ ':authority': `127.0.0.1:${port}` }), 200);
  assert.equal(await request({ host: 'evil.test' }), 403);
  assert.equal(await request({ ':authority': 'evil.test' }), 403);
  assert.equal(await request({ origin: '' }), 403);
});

test('shared service cannot be mounted on public or plaintext origins', () => {
  for (const allowedOrigin of [
    'http://localhost:3000',
    'https://example.test',
    'https://localhost:3000/path',
    'https://user:pass@localhost:3000',
  ])
    assert.throws(() => createSharedModelMiddleware({ store: {}, allowedOrigin }));
});
