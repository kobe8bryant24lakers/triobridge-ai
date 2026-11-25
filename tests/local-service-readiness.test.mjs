import assert from 'node:assert/strict';
import { execFile } from 'node:child_process';
import { once } from 'node:events';
import { createServer as createProxyCanary } from 'node:http';
import { mkdtemp, readFile, rm, writeFile } from 'node:fs/promises';
import { createServer } from 'node:https';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { after, before, test } from 'node:test';
import { promisify } from 'node:util';
import devCerts from 'office-addin-dev-certs';
import { checkLoopbackPort, waitForLocalEndpoint } from '../scripts/local-service-readiness.mjs';
import { storageDetailHeader } from '../scripts/model-storage-diagnostic.mjs';
import { getOrCreateLocalCapability } from '../scripts/local-capability.mjs';
import { enforcePrivatePath } from '../scripts/private-local-permissions.mjs';
import { createSharedModelRuntime } from '../scripts/shared-model-runtime.mjs';

let directory, ca, cert, key;
const version = '1.2.3';
const installationId = 'test-installation';
const token = '1'.repeat(64);
const health = { product: 'triobridge-ai-local', installationId, version };
const state = { version: 2, revision: 0, profiles: [] };

before(async () => {
  directory = await mkdtemp(join(tmpdir(), 'triobridge-readiness-'));
  const paths = ['ca.pem', 'cert.pem', 'key.pem'].map((name) => join(directory, name));
  await devCerts.generateCertificates(...paths, 1, ['localhost']);
  [ca, cert, key] = await Promise.all(paths.map((path) => readFile(path)));
});
after(() => directory && rm(directory, { recursive: true, force: true }));

async function service(t, handler) {
  const server = createServer({ cert, key }, handler);
  server.listen(0, '127.0.0.1');
  await once(server, 'listening');
  t.after(() => new Promise((done) => server.close(done)));
  return server.address().port;
}

test('a cold configuration read taking longer than two seconds succeeds over verified loopback TLS', async (t) => {
  const seen = [];
  const port = await service(t, (req, res) => {
    seen.push([req.url, req.method, req.headers['x-triobridge-local-token']]);
    if (req.url === '/local-health') res.end(JSON.stringify(health));
    else setTimeout(() => res.end(JSON.stringify(state)), 2400);
  });
  await waitForLocalEndpoint({ ca, port, installationId, version });
  await waitForLocalEndpoint({ ca, port, token, timeout: 7000 });
  assert.deepEqual(seen, [
    ['/local-health', 'GET', undefined],
    ['/local-models/state', 'POST', token],
  ]);
});

test('the read-only state probe retries a dropped connection', async (t) => {
  let attempts = 0;
  const port = await service(t, (req, res) => {
    if (++attempts === 1) req.socket.destroy();
    else res.end(JSON.stringify(state));
  });
  await waitForLocalEndpoint({ ca, port, token, timeout: 3000 });
  assert.equal(attempts, 2);
});

test('a known catalog lock is retried, while unrecognized error details never reach diagnostics', async (t) => {
  let requests = 0;
  const port = await service(t, (_req, res) => {
    requests++;
    if (requests === 2) return res.end(JSON.stringify(state));
    res.writeHead(503, {
      [storageDetailHeader]: requests === 1 ? 'CATALOG_BUSY' : 'private-test-value',
    });
    res.end('private-test-body');
  });
  await waitForLocalEndpoint({ ca, port, token, timeout: 3000 });
  await assert.rejects(waitForLocalEndpoint({ ca, port, token }), (error) => {
    assert.equal(error.code, 'MODEL_STORAGE');
    assert.equal(error.storageDetail, undefined);
    assert.doesNotMatch(error.message, /private-test/);
    return true;
  });
  assert.equal(requests, 3);
});

test(
  'the same installed middleware recovers after a bad first catalog read without replacing user data',
  { skip: !['darwin', 'win32'].includes(process.platform) },
  async (t) => {
    const temporary = await mkdtemp(join(tmpdir(), 'triobridge-runtime-models-'));
    t.after(() => rm(temporary, { recursive: true, force: true }));
    const models = join(temporary, 'models');
    const capability = await getOrCreateLocalCapability(models);
    const path = join(models, 'profiles.json');
    const invalid = 'synthetic unreadable catalog';
    await writeFile(path, invalid);
    await enforcePrivatePath(path, 'file', { created: true });
    let runtime;
    const port = await service(t, async (req, res) => (await runtime)(req, res, () => res.end()));
    runtime = createSharedModelRuntime('https://localhost:' + port, {}, { directory: models });
    await assert.rejects(waitForLocalEndpoint({ ca, port, token: capability }), {
      code: 'MODEL_STORAGE',
      storageDetail: 'CATALOG_JSON',
    });
    assert.equal(await readFile(path, 'utf8'), invalid);
    const restored =
      '\uFEFF' +
      JSON.stringify({
        version: 2,
        revision: 4,
        profiles: [],
        activeId: '',
        retiredCredentials: [],
      });
    await writeFile(path, restored);
    await waitForLocalEndpoint({ ca, port, token: capability });
    assert.equal(await readFile(path, 'utf8'), restored);
  },
);

test('TLS and service identity failures are terminal and never send a capability', async (t) => {
  const tokens = [];
  const port = await service(t, (req, res) => {
    tokens.push(req.headers['x-triobridge-local-token']);
    res.end(JSON.stringify({ ...health, installationId: 'other-installation' }));
  });
  await assert.rejects(waitForLocalEndpoint({ ca: [], port, installationId, version }), {
    code: 'TLS_CERTIFICATE',
  });
  await assert.rejects(waitForLocalEndpoint({ ca, port, installationId, version }), {
    code: 'SERVICE_IDENTITY',
  });
  assert.deepEqual(tokens, [undefined]);
});

test('authentication, storage, invalid JSON and wrong state format report safe errors without retry', async (t) => {
  const responses = [
    [401, 'LOCAL_AUTH'],
    [403, 'LOCAL_AUTH'],
    [503, 'MODEL_STORAGE'],
    [200, 'INVALID_RESPONSE'],
    [200, 'INVALID_RESPONSE'],
    [200, 'INVALID_RESPONSE'],
  ];
  const bodies = [
    'private response',
    'private response',
    'private response',
    'private response',
    '{}',
    'x'.repeat(512_001),
  ];
  let requests = 0;
  const port = await service(t, (_req, res) => {
    const index = requests++;
    res.writeHead(responses[index][0]);
    res.end(bodies[index]);
  });
  for (const [, code] of responses) {
    await assert.rejects(waitForLocalEndpoint({ ca, port, token }), (error) => {
      assert.equal(error.code, code);
      assert.doesNotMatch(error.message, /private response/);
      return true;
    });
  }
  assert.equal(requests, responses.length);
});

test('the absolute deadline also bounds a response that keeps sending individual bytes', async (t) => {
  const port = await service(t, (_req, res) => {
    const timer = setInterval(() => res.write(' '), 20);
    res.on('close', () => clearInterval(timer));
  });
  await assert.rejects(waitForLocalEndpoint({ ca, port, token, timeout: 350 }), {
    code: 'TIMEOUT',
  });
});

test('the preflight leaves an occupied port and its owner running', async (t) => {
  const port = await service(t, (_req, res) => res.end(JSON.stringify(health)));
  await assert.rejects(checkLoopbackPort(port), { code: 'PORT_BUSY' });
  await waitForLocalEndpoint({ ca, port, installationId, version });
});

test('only the control request reaches the configured proxy, while authenticated readiness connects directly', async (t) => {
  const port = await service(t, (_req, res) => res.end(JSON.stringify(state)));
  let tunnels = 0;
  const proxy = createProxyCanary();
  proxy.on('connect', (_req, socket) => {
    tunnels++;
    socket.destroy();
  });
  proxy.listen(0, '127.0.0.1');
  await once(proxy, 'listening');
  t.after(() => new Promise((done) => proxy.close(done)));
  const env = Object.fromEntries(
    Object.entries(process.env).filter(
      ([key]) => !/^(?:https?_proxy|all_proxy|no_proxy|node_use_env_proxy)$/i.test(key),
    ),
  );
  const address = 'http://127.0.0.1:' + proxy.address().port;
  const module = new URL('../scripts/local-service-readiness.mjs', import.meta.url).href;
  const child = `
    import { readFile } from 'node:fs/promises';
    import { get } from 'node:https';
    import { waitForLocalEndpoint } from ${JSON.stringify(module)};
    const ca = await readFile(process.argv[1]);
    const port = Number(process.argv[2]);
    if (process.allowedNodeEnvironmentFlags.has('--use-env-proxy')) {
      await new Promise((done) => get({ host: '127.0.0.1', port, servername: 'localhost', ca }, res => {
        res.resume(); res.on('end', done);
      }).on('error', done));
    }
    await waitForLocalEndpoint({ ca, port, token: '1'.repeat(64), timeout: 3000 });
  `;
  await promisify(execFile)(
    process.execPath,
    ['--input-type=module', '-e', child, join(directory, 'ca.pem'), String(port)],
    {
      env: {
        ...env,
        NODE_USE_ENV_PROXY: '1',
        HTTPS_PROXY: address,
        HTTP_PROXY: address,
        NO_PROXY: '',
      },
      timeout: 10_000,
    },
  );
  assert.equal(tunnels, process.allowedNodeEnvironmentFlags.has('--use-env-proxy') ? 1 : 0);
});
