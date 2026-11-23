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
    const run = (action, runner = lifecycleRunner) =>
      execute(node, [runner, action], {
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
    const call = async (action, body) => {
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
                assert.equal(res.statusCode, 200);
                done(JSON.parse(Buffer.concat(chunks).toString('utf8')));
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
    await run('install');
    const first = await readInstallation(root);
    checks.push('native install, generated TLS certificate, login service and Office catalog');
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
    const diagnosis = await execute(
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
    const report = JSON.parse(diagnosis.stdout.replace(/^\uFEFF/, ''));
    assert.deepEqual(Object.keys(report).sort(), [
      'catalogs',
      'detectedVersions',
      'format',
      'runningOffice',
    ]);
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
