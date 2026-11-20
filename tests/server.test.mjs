import { test } from 'node:test';
import assert from 'node:assert/strict';
import { createServer } from 'node:http';
import { once } from 'node:events';
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

test('local model proxy enforces fixed target, source, path and body boundaries', async (t) => {
  let lastRequest;
  const upstream = createServer(async (req, res) => {
    const parts = [];
    for await (const chunk of req) parts.push(chunk);
    lastRequest = {
      url: req.url,
      authorization: req.headers.authorization,
      body: JSON.parse(Buffer.concat(parts)),
    };
    if (req.url === '/redirect/chat/completions') {
      res.writeHead(302, { Location: 'https://example.invalid' });
      res.end();
      return;
    }
    res.writeHead(200, { 'Content-Type': 'application/json' });
    res.end(JSON.stringify({ choices: [{ message: { content: 'Hello' } }] }));
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
