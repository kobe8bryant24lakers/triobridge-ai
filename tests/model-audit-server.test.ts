// @vitest-environment node
import { afterEach, describe, expect, it } from 'vitest';
import { createServer, request as httpRequest, type Server } from 'node:http';
import { connect, createServer as createHttp2Server } from 'node:http2';
import { once } from 'node:events';
import { mkdtemp, readFile, rm, stat } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { createModelAuditMiddleware } from '../scripts/model-audit';
import { MODEL_AUDIT_PATH } from '../src/core/model-audit';

const servers: Server[] = [];
const folders: string[] = [];
afterEach(async () => {
  for (const server of servers.splice(0)) {
    server.closeAllConnections();
    await new Promise<void>((resolve) => server.close(() => resolve()));
  }
  await Promise.all(
    folders.splice(0).map((folder) => rm(folder, { recursive: true, force: true })),
  );
});

async function setup() {
  const folder = await mkdtemp(join(tmpdir(), 'triobridge-audit-'));
  folders.push(folder);
  const output = join(folder, 'private', 'requests.jsonl');
  let handler: ReturnType<typeof createModelAuditMiddleware>;
  const server = createServer((req, res) =>
    handler(req, res, () => {
      res.writeHead(404);
      res.end();
    }),
  );
  servers.push(server);
  server.listen(0, '127.0.0.1');
  await once(server, 'listening');
  const port = (server.address() as { port: number }).port;
  const origin = `http://localhost:${port}`;
  handler = createModelAuditMiddleware(origin, output);
  const record = {
    version: 1,
    auditId: 'synthetic-1',
    startedAt: new Date().toISOString(),
    elapsedMs: 14,
    endpoint: 'https://example.test/v1/chat/completions',
    requestedModel: 'test-model',
    messages: [{ role: 'user', content: '合成数据' }],
    responseContent: '合成响应',
    outcome: 'success',
  };
  const request = (overrides: RequestInit = {}, path = MODEL_AUDIT_PATH) =>
    fetch(`http://127.0.0.1:${port}${path}`, {
      method: 'POST',
      headers: { Origin: origin, 'Content-Type': 'application/json' },
      body: JSON.stringify(record),
      ...overrides,
    });
  return { output, request, record, origin, port };
}

describe('same-origin local audit sink', () => {
  it('appends only whitelisted fields to a private fixed file', async () => {
    const { output, request, record, origin } = await setup();
    const secret = 'secret-extra-field';
    expect(
      (
        await request({
          headers: {
            Origin: origin,
            'Content-Type': 'application/json',
            Authorization: secret,
          },
          body: JSON.stringify({
            ...record,
            apiKey: secret,
            headers: { Authorization: secret },
            outputFile: '/tmp/not-allowed',
            usage: { total_tokens: 42, secret },
          }),
        })
      ).status,
    ).toBe(204);
    expect((await request()).status).toBe(204);
    const content = await readFile(output, 'utf8');
    const lines = content
      .trim()
      .split('\n')
      .map((line) => JSON.parse(line));
    expect(lines).toHaveLength(2);
    expect(lines[0]).toMatchObject({ ...record, usage: { total_tokens: 42 } });
    expect(content).not.toContain(secret);
    expect(content).not.toContain('outputFile');
    expect((await stat(output)).mode & 0o777).toBe(0o600);
  });

  it('rejects cross-origin, missing origin, rebinding hosts and unsupported methods before writing', async () => {
    const { request, output, origin, port, record } = await setup();
    const invalidHeaders: Record<string, string>[] = [
      { Origin: 'https://external.test', 'Content-Type': 'application/json' },
      { 'Content-Type': 'application/json' },
    ];
    for (const headers of invalidHeaders) expect((await request({ headers })).status).toBe(403);
    // fetch normalizes Host, so use a raw HTTP client to exercise rebinding protection.
    const reboundStatus = await new Promise<number | undefined>((resolve, reject) => {
      const req = httpRequest(
        `http://127.0.0.1:${port}${MODEL_AUDIT_PATH}`,
        {
          method: 'POST',
          headers: { Origin: origin, Host: 'evil.test:3000', 'Content-Type': 'application/json' },
        },
        (res) => {
          res.resume();
          res.on('end', () => resolve(res.statusCode));
        },
      );
      req.on('error', reject);
      req.end(JSON.stringify(record));
    });
    expect(reboundStatus).toBe(403);
    expect((await request({ method: 'GET', body: undefined })).status).toBe(405);
    expect(
      (await request({ headers: { Origin: origin, 'Content-Type': 'text/plain' } })).status,
    ).toBe(405);
    await expect(stat(output)).rejects.toMatchObject({ code: 'ENOENT' });
  });

  it('rejects query/path overrides, malformed records and oversized bodies before writing', async () => {
    const { request, output } = await setup();
    expect((await request({}, `${MODEL_AUDIT_PATH}?path=/tmp/unsafe`)).status).toBe(400);
    expect((await request({}, `${MODEL_AUDIT_PATH}/extra`)).status).toBe(400);
    expect((await request({ body: 'invalid' })).status).toBe(400);
    expect((await request({ body: '{}' })).status).toBe(400);
    expect((await request({ body: 'x'.repeat(512_001) })).status).toBe(413);
    await expect(stat(output)).rejects.toMatchObject({ code: 'ENOENT' });
  });

  it('rejects a non-loopback server origin', () => {
    expect(() => createModelAuditMiddleware('https://example.test', '/tmp/not-used')).toThrow();
  });

  it('accepts native Office HTTP/2 authority and rejects conflicting destinations', async () => {
    const folder = await mkdtemp(join(tmpdir(), 'triobridge-audit-http2-'));
    folders.push(folder);
    const output = join(folder, 'requests.jsonl');
    let handler: ReturnType<typeof createModelAuditMiddleware>;
    const server = createHttp2Server((req, res) => handler(req, res, () => res.end()));
    server.listen(0, '127.0.0.1');
    await once(server, 'listening');
    const port = (server.address() as { port: number }).port;
    const origin = `http://localhost:${port}`;
    handler = createModelAuditMiddleware(origin, output);
    const client = connect(`http://127.0.0.1:${port}`);
    const request = (headers: Record<string, string>) =>
      new Promise<number>((resolve, reject) => {
        const stream = client.request({
          ':method': 'POST',
          ':path': MODEL_AUDIT_PATH,
          ':authority': `localhost:${port}`,
          origin,
          'content-type': 'application/json',
          ...headers,
        });
        let status = 0;
        stream.on('response', (responseHeaders) => {
          status = Number(responseHeaders[':status']);
          stream.resume();
        });
        stream.on('error', reject);
        stream.on('end', () => resolve(status));
        stream.end(
          JSON.stringify({
            version: 1,
            auditId: 'http2-test',
            startedAt: new Date().toISOString(),
            elapsedMs: 20,
            endpoint: 'https://example.test/v1/chat/completions',
            requestedModel: 'test',
            messages: [{ role: 'user', content: 'synthetic' }],
            outcome: 'success',
          }),
        );
      });
    try {
      expect(await request({})).toBe(204);
      expect(await request({ ':authority': `127.0.0.1:${port}` })).toBe(204);
      expect(await request({ ':authority': 'evil.test' })).toBe(403);
      expect(await request({ host: 'evil.test' })).toBe(403);
      expect(await request({ origin: '' })).toBe(403);
      expect((await readFile(output, 'utf8')).trim().split('\n')).toHaveLength(2);
    } finally {
      client.close();
      await new Promise<void>((resolve) => server.close(() => resolve()));
    }
  });
});
