import { test } from 'node:test';
import assert from 'node:assert/strict';
import { createServer } from 'node:http';
import { createServer as createHttp2Server, connect } from 'node:http2';
import { once } from 'node:events';
import { Readable, Writable } from 'node:stream';
import { sendModelRequest } from '../scripts/model-protocol.mjs';
import { createProxy } from '../scripts/server-utils.mjs';

async function listen(server) {
  server.listen(0, '127.0.0.1');
  await once(server, 'listening');
  return server.address().port;
}
async function close(server) {
  server.closeAllConnections();
  await new Promise((resolve) => server.close(resolve));
}

test('only a deck completion directed to the local proxy carries the internal purpose header', async () => {
  const messages = [{ role: 'user', content: 'Prepare a deck' }];
  const calls = [];
  const fetcher = async (url, options) => {
    const body = options.body && JSON.parse(options.body);
    calls.push([url, options.headers['x-triobridge-purpose']]);
    assert.equal(body?.purpose, undefined);
    if (url.endsWith('/models/local')) return Response.json({ max_tokens: 8192 });
    if (url.endsWith('/messages') && body.max_tokens === undefined)
      return Response.json({ error: { message: 'max_tokens is required' } }, { status: 400 });
    return Response.json({ done: true });
  };
  const claude = '/model-proxy/v1/messages';
  const openai = '/model-proxy/v1/chat/completions';
  const direct = 'https://example.test/v1/chat/completions';
  const send = (url, format, purpose) =>
    sendModelRequest(fetcher, url, format, 'local', messages, '', undefined, purpose);
  await send(claude, 'anthropic', 'presentation-deck');
  await send(openai, 'openai', 'presentation-deck');
  await send(direct, 'openai', 'presentation-deck');
  await send(openai, 'openai');
  assert.deepEqual(calls, [
    [claude, 'presentation-deck'],
    ['/model-proxy/v1/models/local', undefined],
    [claude, 'presentation-deck'],
    [openai, 'presentation-deck'],
    [direct, undefined],
    [openai, undefined],
  ]);
});

test('proxy deadline excludes deck generation, preserves cancellation and bounds ordinary requests and lookups', async (t) => {
  t.mock.timers.enable({ apis: ['setTimeout'] });
  const origin = 'http://localhost:3000';
  const proxy = createProxy('http://127.0.0.1:8001', origin);
  const chat = '/v1/chat/completions';
  const body = JSON.stringify({ stream: false, messages: [{ role: 'user', content: 'Example' }] });
  for (const [path, purpose, disconnect, expected] of [
    [chat, 'presentation-deck', false, 200],
    [chat, 'presentation-deck', true, 504],
    [chat, undefined, false, 504],
    [chat, 'unknown', false, 504],
    ['/v1/models/local', 'presentation-deck', false, 504],
  ]) {
    const request = Readable.from(path === chat ? [Buffer.from(body)] : []);
    request.url = '/model-proxy' + path;
    request.method = 'POST';
    request.headers = {
      origin,
      host: 'localhost:3000',
      'content-type': 'application/json',
      'x-triobridge-purpose': purpose,
    };
    const response = new Writable({
      write(_chunk, _encoding, callback) {
        callback();
      },
    });
    let status;
    response.writeHead = (value) => {
      status = value;
    };
    const upstream = t.mock.method(globalThis, 'fetch', async (_url, options) => {
      assert.equal(new Headers(options.headers).has('x-triobridge-purpose'), false);
      t.mock.timers.tick(300_001);
      if (disconnect) response.emit('close');
      options.signal.throwIfAborted();
      return new Response(null, { status: 200 });
    });
    await proxy(request, response, assert.fail);
    upstream.mock.restore();
    assert.equal(status, expected);
  }
});

test('local model proxy enforces fixed target, source, path and body boundaries', async (t) => {
  let lastRequest;
  const upstream = createServer(async (req, res) => {
    const parts = [];
    for await (const chunk of req) parts.push(chunk);
    lastRequest = {
      url: req.url,
      method: req.method,
      authorization: req.headers.authorization,
      key: req.headers['x-api-key'],
      version: req.headers['anthropic-version'],
      body: parts.length ? JSON.parse(Buffer.concat(parts)) : null,
    };
    if (req.url === '/v1/models/org%2Flocal%3Frevision%3D1') {
      res.writeHead(200, { 'Content-Type': 'application/json' });
      res.end(JSON.stringify({ max_tokens: 65536 }));
      return;
    }
    if (req.url === '/redirect/chat/completions') {
      res.writeHead(302, { Location: 'https://example.invalid' });
      res.end();
      return;
    }
    res.writeHead(200, { 'Content-Type': 'application/json' });
    const content = req.url === '/long/messages' ? '全'.repeat(350000) : 'Hello';
    res.end(JSON.stringify({ choices: [{ message: { content } }] }));
  });
  const upstreamPort = await listen(upstream);
  let handler;
  const server = createServer((req, res) => {
    handler(req, res, () => {
      res.writeHead(404);
      res.end();
    });
  });
  const port = await listen(server);
  const origin = `http://localhost:${port}`;
  const base = `http://127.0.0.1:${port}`;
  handler = createProxy(`http://127.0.0.1:${upstreamPort}`, origin);
  const request = (path = '/model-proxy/v1/chat/completions', overrides = {}) =>
    fetch(base + path, {
      method: 'POST',
      headers: {
        Origin: origin,
        'Content-Type': 'application/json',
        Authorization: 'Bearer test-only',
      },
      body: JSON.stringify({
        model: 'test',
        messages: [{ role: 'user', content: 'Hello' }],
        stream: false,
      }),
      ...overrides,
    });
  try {
    await t.test('forwards a valid request and the authorization header', async () => {
      const response = await request();
      assert.equal(response.status, 200);
      assert.equal((await response.json()).choices[0].message.content, 'Hello');
      assert.equal(lastRequest.url, '/v1/chat/completions');
      assert.equal(lastRequest.authorization, 'Bearer test-only');
    });
    await t.test('rejects cross-origin and missing-origin requests', async () => {
      for (const headers of [
        {
          Origin: 'https://malicious.example',
          'Content-Type': 'application/json',
        },
        { 'Content-Type': 'application/json' },
      ])
        assert.equal((await request(undefined, { headers })).status, 403);
    });
    await t.test(
      'forwards the Messages credential only to the configured Messages endpoint',
      async () => {
        const res = await request('/model-proxy/v1/messages', {
          headers: {
            Origin: origin,
            'Content-Type': 'application/json',
            Authorization: 'Bearer unrelated',
            'x-api-key': 'synthetic-messages-key',
          },
        });
        assert.equal(res.status, 200);
        assert.equal(lastRequest.url, '/v1/messages');
        assert.equal(lastRequest.key, 'synthetic-messages-key');
        assert.equal(lastRequest.version, '2023-06-01');
        assert.equal(lastRequest.authorization, undefined);
      },
    );
    await t.test('rejects unsupported paths, body and content types', async () => {
      assert.equal((await request('/model-proxy/admin')).status, 400);
      assert.equal(
        (await request('/model-proxy/v1/chat/completions?target=example.com')).status,
        400,
      );
      assert.equal((await request(undefined, { body: 'invalid' })).status, 400);
      assert.equal((await request(undefined, { body: 'x'.repeat(513000) })).status, 413);
      assert.equal(
        (
          await request(undefined, {
            headers: { Origin: origin, 'Content-Type': 'text/plain' },
          })
        ).status,
        405,
      );
    });
    await t.test(
      'relays a same-origin model lookup as a bodyless GET to the configured service',
      async () => {
        const url = '/model-proxy/v1/models/org%2Flocal%3Frevision%3D1';
        const headers = {
          Origin: origin,
          'Content-Type': 'application/json',
          'x-api-key': 'model-key',
        };
        const response = await request(url, { headers, body: undefined });
        assert.deepEqual(await response.json(), { max_tokens: 65536 });
        assert.deepEqual(lastRequest, {
          url: '/v1/models/org%2Flocal%3Frevision%3D1',
          method: 'GET',
          body: null,
          key: 'model-key',
          authorization: undefined,
          version: '2023-06-01',
        });
        const previous = lastRequest;
        assert.equal((await request(url, { headers })).status, 400);
        assert.equal(
          (await request('/model-proxy/v1/models/org%2F..%2Fadmin', { headers, body: undefined }))
            .status,
          400,
        );
        assert.equal(
          (
            await request(url, {
              headers: { ...headers, Origin: 'https://example.invalid' },
              body: undefined,
            })
          ).status,
          403,
        );
        assert.equal(lastRequest, previous);
      },
    );
    await t.test('forwards a complete long reply through the proxy', async () => {
      const response = await request('/model-proxy/long/messages');
      assert.equal(response.status, 200);
      assert.equal((await response.json()).choices[0].message.content, '全'.repeat(350000));
    });
    await t.test('does not follow upstream redirects', async () => {
      assert.equal((await request('/model-proxy/redirect/chat/completions')).status, 502);
    });
    await t.test('reports a missing configured upstream', async () => {
      handler = createProxy(undefined, origin);
      assert.equal((await request()).status, 503);
    });
  } finally {
    await close(server);
    await close(upstream);
  }
});

test('proxy configuration rejects remote plaintext and URL credentials', () => {
  for (const target of [
    'http://example.com',
    'https://user:pass@example.com',
    'https://example.com/v1',
    'file:///tmp/test',
  ])
    assert.throws(() => createProxy(target, 'http://localhost:3000'));
});

test('HTTP/2 proxy accepts local authority while preserving origin and host boundaries', async () => {
  let forwarded = 0;
  const upstream = createServer((req, res) => {
    forwarded++;
    req.resume();
    res.writeHead(200, { 'Content-Type': 'application/json' });
    res.end(JSON.stringify({ choices: [{ message: { content: 'OK' } }] }));
  });
  const upstreamPort = await listen(upstream);
  let proxy;
  const server = createHttp2Server((req, res) => proxy(req, res, () => res.end()));
  const port = await listen(server);
  const origin = `http://localhost:${port}`;
  proxy = createProxy(`http://127.0.0.1:${upstreamPort}`, origin);
  const client = connect(`http://127.0.0.1:${port}`);
  const request = async (headers = {}) => {
    const stream = client.request({
      ':method': 'POST',
      ':path': '/model-proxy/v1/chat/completions',
      ':authority': `localhost:${port}`,
      origin,
      'content-type': 'application/json',
      ...headers,
    });
    const response = once(stream, 'response');
    stream.end(JSON.stringify({ stream: false, messages: [{ role: 'user', content: 'OK' }] }));
    const [result] = await response;
    for await (const chunk of stream) void chunk;
    return result[':status'];
  };
  try {
    assert.equal(await request(), 200);
    assert.equal(await request({ ':authority': `127.0.0.1:${port}` }), 200);
    assert.equal(await request({ ':authority': 'untrusted.example' }), 403);
    assert.equal(await request({ host: 'untrusted.example' }), 403);
    assert.equal(await request({ origin: 'https://untrusted.example' }), 403);
    assert.equal(await request({ origin: undefined }), 403);
    assert.equal(forwarded, 2);
  } finally {
    client.destroy();
    await new Promise((resolve) => server.close(resolve));
    await close(upstream);
  }
});
