import assert from 'node:assert/strict';
import { execFile, spawn } from 'node:child_process';
import { promisify } from 'node:util';
import { once } from 'node:events';
import { randomUUID } from 'node:crypto';
import { createServer } from 'node:http';
import { request } from 'node:https';
import { copyFile, cp, lstat, mkdtemp, readFile, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join, resolve } from 'node:path';
import { pathToFileURL } from 'node:url';
import { verifyPackage, installedRoot, readInstallation } from './local-install-files.mjs';
import { localInstallerName, packageFiles } from './build-local-installers.mjs';
import { sharedModelsDirectory } from './shared-model-store.mjs';
import { createNativeModelVault } from './shared-model-vault.mjs';
import { windowsPowerShellEnvironment } from './private-local-permissions.mjs';
import { waitForLocalEndpoint } from './local-service-readiness.mjs';
import { storageDetailHeader } from './model-storage-diagnostic.mjs';

const execute = promisify(execFile);
const archive = resolve(process.argv[2]);
const temporary = await mkdtemp(join(tmpdir(), 'triobridge-package-verification-'));
const name = localInstallerName(process.platform, process.arch);
const directory = join(temporary, name);
let nativeReference;
let nativeModelsCreated = false;
let fixtureApi;
let lifecycleRunner;
let simulatedOffice;
async function closeSimulatedOffice() {
  if (simulatedOffice && simulatedOffice.exitCode === null && simulatedOffice.signalCode === null) {
    const exited = once(simulatedOffice, 'exit');
    simulatedOffice.kill();
    await exited;
  }
  simulatedOffice = undefined;
}
const checks = [];
try {
  const listing = await execute(
    process.platform === 'win32' ? 'tar' : 'unzip',
    process.platform === 'win32' ? ['-tf', archive] : ['-Z1', archive],
  );
  for (const path of listing.stdout.trim().split(/\r?\n/)) {
    assert.ok(path.startsWith(name + '/') && !path.split('/').includes('..'));
    assert.doesNotMatch(
      path,
      /(?:^|\/)(?:\.env|local-capability|profiles\.json|server\.(?:key|crt|pfx|password))$/,
    );
  }
  await execute(
    process.platform === 'win32' ? 'tar' : 'unzip',
    process.platform === 'win32'
      ? ['-xf', archive, '-C', temporary]
      : ['-q', archive, '-d', temporary],
  );
  const metadata = await verifyPackage(directory);
  checks.push('package integrity and platform');
  const node = join(directory, process.platform === 'win32' ? 'runtime/node.exe' : 'runtime/node');
  const { stdout } = await execute(
    node,
    [
      '--input-type=module',
      '-e',
      "const{AsyncEntry}=await import('@napi-rs/keyring');if(typeof AsyncEntry!=='function')process.exit(1);console.log(JSON.stringify({version:process.version,arch:process.arch}))",
    ],
    { cwd: join(directory, 'app'), windowsHide: true },
  );
  assert.deepEqual(JSON.parse(stdout), { version: metadata.nodeVersion, arch: process.arch });
  checks.push('bundled Node and native credential binding');
  if (process.argv.includes('--install')) {
    assert.equal(
      process.env.GITHUB_ACTIONS,
      'true',
      'Native lifecycle verification is restricted to a fresh CI account.',
    );
    assert.equal(process.platform, 'win32');
    const root = installedRoot();
    const models = sharedModelsDirectory();
    for (const path of [root, models]) await assert.rejects(lstat(path), { code: 'ENOENT' });
    nativeModelsCreated = true;
    // Windows asks for interactive confirmation when modifying CurrentUser Root.
    // CI has no person to approve it. Exercise the installed core with only this
    // OS confirmation boundary replaced. Every HTTPS request still validates the
    // real generated certificate via its explicit CA; no TLS checks are disabled.
    const createRunner = async (packageDirectory, name) => {
      const runner = join(temporary, name);
      const moduleUrl = (name) =>
        JSON.stringify(pathToFileURL(join(packageDirectory, 'app/scripts', name)).href);
      await writeFile(
        runner,
        `
import { runInstaller } from ${moduleUrl('install-local.mjs')};
import { platformDriver, certificateIdentity } from ${moduleUrl('local-install-platform.mjs')};
import { installedRoot } from ${moduleUrl('local-install-files.mjs')};
import { windowsPowerShellEnvironment } from ${moduleUrl('private-local-permissions.mjs')};
import { execFile } from 'node:child_process';
import { promisify } from 'node:util';
import { readFile } from 'node:fs/promises';
import { join } from 'node:path';
if (process.env.GITHUB_ACTIONS !== 'true' || process.platform !== 'win32') throw new Error('CI only');
const driver = platformDriver(installedRoot());
driver.trustCertificate = async (directory, id) => { await certificateIdentity(directory, id); };
let starts = 0;
async function startupDiagnostics() {
  // Only this freshly created CI installation is inspected. Do not print task
  // arguments, private manifests, model metadata or credential values.
  const command = "$sid = [Security.Principal.WindowsIdentity]::GetCurrent().User.Value; $name = 'TrioBridgeAI-Local-' + $sid; $task = Get-ScheduledTask -TaskName $name -ErrorAction SilentlyContinue; $info = Get-ScheduledTaskInfo -TaskName $name -ErrorAction SilentlyContinue; [PSCustomObject]@{state = [string]$task.State; lastResult = $info.LastTaskResult; session = (Get-Process -Id $PID).SessionId} | ConvertTo-Json -Compress";
  try {
    const result = await promisify(execFile)(join(process.env.SystemRoot, 'System32/WindowsPowerShell/v1.0/powershell.exe'), ['-NoProfile', '-NonInteractive', '-Command', command], { env: windowsPowerShellEnvironment(), windowsHide: true, timeout: 60000 });
    console.log('Native task diagnostic ' + result.stdout.trim());
  } catch { console.log('Native task diagnostic unavailable'); }
  for (const name of ['service.log', 'service-error.log']) {
    try {
      const text = await readFile(join(installedRoot(), 'logs', name), 'utf8');
      console.log('Native service diagnostic ' + name + ': ' + text.slice(-2000));
    } catch { console.log('Native service diagnostic ' + name + ': absent'); }
  }
}
for (const [name, operation] of Object.entries(driver)) {
  driver[name] = async (...args) => {
    if (name === 'stop' && starts > 0) await startupDiagnostics();
    if (name === 'start') starts++;
    const start = Date.now();
    console.log('Native step ' + name + ' started');
    try { return await operation(...args); }
    finally { console.log('Native step ' + name + ' elapsed ' + (Date.now() - start) + ' ms'); }
  };
}
await runInstaller(process.argv[2], { driver });
`,
      );
      return runner;
    };
    lifecycleRunner = await createRunner(directory, 'native-lifecycle.mjs');
    const run = (action, runner = lifecycleRunner, env = process.env) =>
      execute(node, [runner, action], {
        env,
        windowsHide: true,
        timeout: 600_000,
        maxBuffer: 512_000,
      });
    const apiCalls = [];
    fixtureApi = createServer(async (req, res) => {
      const chunks = [];
      for await (const chunk of req) chunks.push(chunk);
      apiCalls.push({
        path: req.url,
        authorization: req.headers.authorization,
        payload: JSON.parse(Buffer.concat(chunks).toString('utf8')),
      });
      res.setHeader('Content-Type', 'application/json');
      res.end(
        JSON.stringify({
          choices: [
            { message: { content: 'synthetic installed transport OK' }, finish_reason: 'stop' },
          ],
        }),
      );
    });
    fixtureApi.listen(0, '127.0.0.1');
    await once(fixtureApi, 'listening');
    const call = async (action, body, expectedStatus = 200) => {
      const ca = await readFile(join(root, 'certificates/server.crt'));
      const token = await readFile(join(models, 'local-capability'), 'utf8');
      return new Promise((done, reject) => {
        const req = request(
          {
            hostname: '127.0.0.1',
            port: 3000,
            servername: 'localhost',
            ca,
            path: '/local-models/' + action,
            method: 'POST',
            headers: {
              Host: 'localhost:3000',
              Origin: 'https://localhost:3000',
              'Content-Type': 'application/json',
              'X-TrioBridge-Local-Token': token,
            },
          },
          (res) => {
            const chunks = [];
            res.on('data', (chunk) => chunks.push(chunk));
            res.on('error', reject);
            res.on('end', () => {
              try {
                assert.equal(res.statusCode, expectedStatus);
                done(
                  expectedStatus === 200
                    ? JSON.parse(Buffer.concat(chunks).toString('utf8'))
                    : res.headers[storageDetailHeader],
                );
              } catch (error) {
                reject(error);
              }
            });
          },
        );
        req.on('error', reject);
        req.setTimeout(15000, () => req.destroy(new Error('Installed API timeout.')));
        req.end(JSON.stringify(body));
      });
    };
    // A renamed bundled Node simulates only the Windows process name/session.
    // It is not Office and never opens a document. Setup must leave it running.
    const officeExecutable = join(temporary, 'WINWORD.exe');
    await copyFile(node, officeExecutable);
    const openSimulatedOffice = async () => {
      simulatedOffice = spawn(
        officeExecutable,
        ['-e', "console.log('SYNTHETIC_OFFICE_READY');setInterval(()=>{},1000)"],
        { windowsHide: true, stdio: ['ignore', 'pipe', 'pipe'] },
      );
      await Promise.race([
        once(simulatedOffice.stdout, 'data'),
        once(simulatedOffice, 'exit').then(() => {
          throw new Error('Synthetic Office process exited before readiness.');
        }),
      ]);
    };
    await openSimulatedOffice();
    await assert.rejects(run('install'), /OFFICE_OPEN/);
    assert.deepEqual([simulatedOffice.exitCode, simulatedOffice.signalCode], [null, null]);
    assert.equal((await readInstallation(root)).active, null);
    await assert.rejects(lstat(models), { code: 'ENOENT' });
    await closeSimulatedOffice();
    checks.push('catalog installation refuses an open simulated Office process without closing it');
    // Simulate a machine with Node's environment proxy enabled and no loopback exemptions.
    const proxyEnv = Object.fromEntries(
      Object.entries(process.env).filter(
        ([key]) => !/^(?:https?_proxy|all_proxy|no_proxy|node_use_env_proxy)$/i.test(key),
      ),
    );
    const proxy = 'http://127.0.0.1:' + fixtureApi.address().port;
    let tunnels = 0;
    fixtureApi.on('connect', (_req, socket) => {
      tunnels++;
      socket.destroy();
    });
    await run('install', lifecycleRunner, {
      ...proxyEnv,
      NODE_USE_ENV_PROXY: '1',
      HTTPS_PROXY: proxy,
      HTTP_PROXY: proxy,
      NO_PROXY: '',
    });
    assert.equal(tunnels, 0);
    assert.equal(apiCalls.length, 0);
    const first = await readInstallation(root);
    checks.push('native install, generated TLS certificate, login service and Office catalog');
    checks.push('local readiness bypasses the environment proxy without sending it any traffic');
    // A stale or replaced PID record must never terminate an unrelated process.
    // Point it at this verifier, whose executable is outside the installed app.
    const processRecordPath = join(root, 'service-process.json');
    await openSimulatedOffice();
    const originalProcessRecord = await readFile(processRecordPath, 'utf8');
    await writeFile(
      processRecordPath,
      JSON.stringify({ ...JSON.parse(originalProcessRecord), pid: process.pid }),
    );
    try {
      await run('restart');
    } catch (error) {
      assert.match(error.message, /3000 端口仍被占用/);
      await writeFile(processRecordPath, originalProcessRecord);
      await run('restart');
    }
    assert.equal((await call('state', {})).profiles.length, 0);
    assert.deepEqual([simulatedOffice.exitCode, simulatedOffice.signalCode], [null, null]);
    await closeSimulatedOffice();
    checks.push('stale PID protection leaves an unrelated process alive');
    checks.push('service restart remains available with a simulated Office process open');
    const id = 'installer-test-' + randomUUID();
    const saved = await call('upsert', {
      expectedRevision: 0,
      profile: {
        id,
        name: 'Installer transport fixture',
        baseUrl: `http://127.0.0.1:${fixtureApi.address().port}/v1`,
        model: 'synthetic-model',
        timeoutSeconds: 10,
      },
      credential: { action: 'replace', value: 'synthetic-installer-api-key' },
      activate: true,
      testConnection: true,
    });
    assert.equal(saved.profiles[0].hasKey, true);
    assert.ok(!JSON.stringify(saved).includes('synthetic-installer-api-key'));
    nativeReference = JSON.parse(await readFile(join(models, 'profiles.json'), 'utf8')).profiles[0]
      .credentialRef;
    checks.push('connect and save through HTTP with native credential storage');
    const diagnose = () =>
      execute(
        join(process.env.SystemRoot, 'System32/WindowsPowerShell/v1.0/powershell.exe'),
        [
          '-NoProfile',
          '-ExecutionPolicy',
          'Bypass',
          '-File',
          join(directory, 'app/scripts/windows-office-discovery.ps1'),
          '-Diagnose',
        ],
        { env: windowsPowerShellEnvironment(), windowsHide: true, timeout: 120_000 },
      );
    const diagnosis = await diagnose();
    const report = JSON.parse(diagnosis.stdout.replace(/^\uFEFF/, ''));
    assert.deepEqual(Object.keys(report).sort(), [
      'catalogs',
      'detectedVersions',
      'format',
      'localService',
      'runningOffice',
    ]);
    assert.equal(report.format, 2);
    assert.equal(report.localService.tcp3000, true);
    assert.equal(report.localService.task.state, 'Running');
    assert.deepEqual(report.localService.lastStartup, { phase: 'ready', code: 'OK' });
    for (const [entry, type] of [
      ['directory', 'directory'],
      ['catalog', 'file'],
      ['capability', 'file'],
    ]) {
      assert.deepEqual(report.localService.storage[entry], {
        status: type,
        ownerMatches: true,
        aclProtected: true,
      });
    }
    assert.deepEqual(
      report.catalogs.find((catalog) => catalog.edition === 'local'),
      {
        edition: 'local',
        registered: true,
        urlMatches: true,
        showInMenu: true,
        reachable: true,
        manifest: { validXml: true, matchesId: true, Word: true, Excel: true, PowerPoint: true },
      },
    );
    assert.ok(!diagnosis.stdout.includes('synthetic-installer-api-key'));
    assert.ok(!diagnosis.stdout.includes(root));
    assert.ok(!diagnosis.stdout.includes(models));
    checks.push('read-only discovery confirms the reachable catalog and all three manifest hosts');
    const catalogPath = join(models, 'profiles.json');
    const originalCatalog = await readFile(catalogPath);
    const held = spawn(
      join(process.env.SystemRoot, 'System32/WindowsPowerShell/v1.0/powershell.exe'),
      [
        '-NoProfile',
        '-NonInteractive',
        '-Command',
        "$ErrorActionPreference='Stop'; $file=[IO.File]::Open($env:TRIOBRIDGE_TEST_CATALOG,'Open','Read','None'); try { [Console]::WriteLine('LOCKED'); [void][Console]::ReadLine() } finally { $file.Dispose() }",
      ],
      {
        windowsHide: true,
        stdio: ['pipe', 'pipe', 'pipe'],
        env: { ...windowsPowerShellEnvironment(), TRIOBRIDGE_TEST_CATALOG: catalogPath },
      },
    );
    const released = once(held, 'exit');
    const lockDeadline = setTimeout(() => held.kill(), 45_000);
    try {
      const output = await Promise.race([
        once(held.stdout, 'data'),
        released.then(() => {
          throw new Error('Catalog lock fixture exited.');
        }),
      ]);
      assert.equal(output[0].toString().trim(), 'LOCKED');
      assert.equal(await call('state', {}, 503), 'CATALOG_BUSY');
      const unlock = setTimeout(() => held.stdin.end('\n'), 1500);
      try {
        await waitForLocalEndpoint({
          ca: await readFile(join(root, 'certificates/server.crt')),
          token: await readFile(join(models, 'local-capability'), 'utf8'),
          timeout: 30_000,
        });
      } finally {
        clearTimeout(unlock);
      }
    } finally {
      if (held.stdin.writable && !held.stdin.writableEnded) held.stdin.end('\n');
      await released;
      clearTimeout(lockDeadline);
    }
    assert.deepEqual(await readFile(catalogPath), originalCatalog);
    checks.push(
      'installed state reports a real Windows file-sharing violation and recovers after unlock',
    );
    const damaged = 'synthetic malformed catalog';
    const restoredBytes = Buffer.concat([Buffer.from([0xef, 0xbb, 0xbf]), originalCatalog]);
    try {
      await writeFile(catalogPath, damaged);
      await assert.rejects(run('restart'), /\[models\/MODEL_STORAGE\].*\[CATALOG_JSON\]/);
      const failedCheck = JSON.parse((await diagnose()).stdout.trim()).localService.lastStartup;
      assert.equal(failedCheck.storageDetail, 'CATALOG_JSON');
      assert.equal(await readFile(catalogPath, 'utf8'), damaged);
    } finally {
      await writeFile(catalogPath, restoredBytes);
    }
    assert.equal((await call('state', {})).profiles[0].id, id);
    assert.deepEqual(await readFile(catalogPath), restoredBytes);
    checks.push(
      'a corrected first-read catalog recovers in the same service process with the original credential',
    );
    await run('install');
    assert.equal((await readInstallation(root)).active, first.active);
    assert.equal((await call('state', {})).profiles[0].id, id);
    checks.push('idempotent repair retains configuration');
    // A valid, separately staged test package fails only when its service starts.
    // Recovery must restore the real scheduled service, not just an in-memory stub.
    const broken = join(temporary, 'broken-upgrade');
    await cp(directory, broken, { recursive: true });
    await writeFile(
      join(broken, 'app/scripts/serve-installed.mjs'),
      "throw new Error('Synthetic upgrade startup failure');\n",
    );
    const brokenMetadata = JSON.parse(
      await readFile(join(broken, 'package-manifest.json'), 'utf8'),
    );
    brokenMetadata.files = (await packageFiles(broken)).filter(
      (file) => file.path !== 'package-manifest.json',
    );
    await writeFile(join(broken, 'package-manifest.json'), JSON.stringify(brokenMetadata) + '\n');
    const brokenRunner = await createRunner(broken, 'broken-lifecycle.mjs');
    await assert.rejects(run('install', brokenRunner), /本机服务未能启动/);
    assert.equal((await readInstallation(root)).active, first.active);
    assert.equal((await call('state', {})).profiles[0].id, id);
    checks.push('failed upgrade activation restores the previous native service and model catalog');
    const afterRollback = JSON.parse((await diagnose()).stdout.trim()).localService;
    assert.equal(afterRollback.tcp3000, true);
    assert.equal(afterRollback.lastStartup.phase, 'https');
    assert.equal(afterRollback.lastStartup.code, 'NO_LISTENER');
    assert.match(afterRollback.lastStartup.taskResult, /^0x[\dA-F]{8}$/);
    checks.push('the failing startup phase and scheduler result survive rollback');
    await run('restart');
    const result = await call('complete', {
      profileId: id,
      profileRevision: saved.profiles[0].revision,
      messages: [{ role: 'user', content: 'synthetic document' }],
    });
    assert.equal(result.choices[0].message.content, 'synthetic installed transport OK');
    assert.equal(apiCalls.length, 2);
    for (const request of apiCalls) {
      assert.equal(request.path, '/v1/chat/completions');
      assert.equal(request.authorization, 'Bearer synthetic-installer-api-key');
    }
    checks.push('generation after managed service restart');
    const before = await readFile(join(models, 'profiles.json'));
    await run('remove');
    await assert.rejects(lstat(root), { code: 'ENOENT' });
    assert.deepEqual(await readFile(join(models, 'profiles.json')), before);
    assert.equal(
      await createNativeModelVault().get(nativeReference),
      'synthetic-installer-api-key',
    );
    checks.push('native uninstall retains model metadata and credentials');
  }
  console.log(
    JSON.stringify(
      {
        platform: process.platform,
        arch: process.arch,
        version: metadata.version,
        nodeVersion: metadata.nodeVersion,
        checks,
        modelResponses: 'synthetic transport fixture; not native Office acceptance',
        interactiveCertificateTrust: process.argv.includes('--install')
          ? 'not exercised in unattended CI; requires manual Windows acceptance'
          : 'not exercised by package integrity verification',
      },
      null,
      2,
    ),
  );
} finally {
  await closeSimulatedOffice();
  if (nativeModelsCreated) {
    const node = join(directory, 'runtime/node.exe');
    try {
      await execute(node, [lifecycleRunner, 'remove'], {
        windowsHide: true,
        timeout: 180_000,
        maxBuffer: 512_000,
      });
    } catch {
      /* Preserve CI failure evidence. */
    }
    if (nativeReference) await createNativeModelVault().delete(nativeReference);
    await rm(sharedModelsDirectory(), { recursive: true, force: true });
  }
  if (fixtureApi) {
    fixtureApi.closeAllConnections();
    await new Promise((done) => fixtureApi.close(done));
  }
  await rm(temporary, { recursive: true, force: true });
}
