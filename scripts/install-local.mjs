import { execFile } from 'node:child_process';
import { promisify } from 'node:util';
import { randomUUID } from 'node:crypto';
import { request } from 'node:https';
import { createConnection } from 'node:net';
import { lstat, open, readFile, readdir, rename, rm, writeFile } from 'node:fs/promises';
import { dirname, join, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import {
  installedRoot,
  initializeInstallation,
  readInstallation,
  writeInstallation,
  stagePackage,
  verifyPackage,
  assertRegularOwned,
  VERSION_DIRECTORY,
} from './local-install-files.mjs';
import { platformDriver, certificateIdentity } from './local-install-platform.mjs';
import { getOrCreateLocalCapability } from './local-capability.mjs';
import { sharedModelsDirectory } from './shared-model-store.mjs';
import { renderLocalManifest, writePrivateLocalManifest } from './local-manifest.mjs';
import { enforcePrivatePath } from './private-local-permissions.mjs';

const execute = promisify(execFile);
const delay = (milliseconds) => new Promise((done) => setTimeout(done, milliseconds));
const packageRoot = resolve(dirname(fileURLToPath(import.meta.url)), '../..');

export async function portOpen() {
  return new Promise((done) => {
    const socket = createConnection({ host: '127.0.0.1', port: 3000 });
    const finish = (value) => {
      socket.destroy();
      done(value);
    };
    socket.once('connect', () => finish(true));
    socket.once('error', () => finish(false));
    socket.setTimeout(1000, () => finish(false));
  });
}

async function requestLocal(root, path, token) {
  const ca = await readFile(join(root, 'certificates/server.crt'));
  return new Promise((done, reject) => {
    const req = request(
      {
        hostname: '127.0.0.1',
        port: 3000,
        servername: 'localhost',
        ca,
        path,
        method: token ? 'POST' : 'GET',
        headers: {
          Host: 'localhost:3000',
          ...(token
            ? {
                Origin: 'https://localhost:3000',
                'Content-Type': 'application/json',
                'X-TrioBridge-Local-Token': token,
              }
            : {}),
        },
      },
      (res) => {
        const chunks = [];
        let size = 0;
        res.on('data', (chunk) => {
          size += chunk.length;
          if (size > 512_000) req.destroy(new Error('Local response too large.'));
          else chunks.push(chunk);
        });
        res.on('error', reject);
        res.on('end', () => {
          try {
            if (res.statusCode !== 200) throw new Error('Local service did not become ready.');
            done(JSON.parse(Buffer.concat(chunks).toString('utf8')));
          } catch (error) {
            reject(error);
          }
        });
      },
    );
    req.on('error', reject);
    req.setTimeout(2000, () => req.destroy(new Error('Local service readiness timeout.')));
    req.end(token ? '{}' : undefined);
  });
}

export async function waitForHealth(root, state, version, timeout = 30_000) {
  const until = Date.now() + timeout;
  do {
    try {
      const health = await requestLocal(root, '/local-health');
      if (
        health.product === 'triobridge-ai-local' &&
        health.installationId === state.id &&
        (!version || health.version === version)
      )
        return health;
    } catch {
      /* The managed service may still be starting. */
    }
    await delay(300);
  } while (Date.now() < until);
  throw new Error(
    '本机服务未能启动。请查看安装目录 logs 中的日志，并检查 3000 端口是否被其他程序占用。',
  );
}

async function stopAndWait(driver) {
  await driver.stop();
  for (let attempt = 0; attempt < 20; attempt++) {
    if (!(await portOpen())) return;
    await delay(200);
  }
  throw new Error(
    '3000 端口仍被占用。请关闭此前手动启动的 TrioBridge 服务或其他占用程序，然后重新双击 Install。未终止其他程序。',
  );
}

async function exists(path) {
  try {
    await lstat(path);
    return true;
  } catch (error) {
    if (error.code === 'ENOENT') return false;
    throw error;
  }
}

async function writeJournal(root, journal) {
  const path = join(root, 'transaction.json');
  await writeFile(path, JSON.stringify(journal) + '\n', { mode: 0o600, flag: 'wx' });
  await enforcePrivatePath(path, 'file', { created: true });
}

export async function recoverInstallation(root, driver, waitReady = waitForHealth) {
  const path = join(root, 'transaction.json');
  if (!(await exists(path))) return;
  await assertRegularOwned(path);
  const journal = JSON.parse(await readFile(path, 'utf8'));
  const state = await readInstallation(root);
  if (
    journal.id !== state.id ||
    (journal.previous !== null && !VERSION_DIRECTORY.test(journal.previous)) ||
    !/^\.cert-next-[a-f0-9-]{36}$/.test(journal.nextCertificate) ||
    !/^\.cert-before-[a-f0-9-]{36}$/.test(journal.previousCertificate)
  )
    throw new Error('未完成安装的恢复记录无效，请保留文件并检查安装日志。');
  await driver.stop();
  const current = join(root, 'certificates');
  const next = join(root, journal.nextCertificate);
  const previous = join(root, journal.previousCertificate);
  if (await exists(previous)) {
    if (await exists(current)) {
      await driver.removeCertificate(current, state.id);
      await rm(current, { recursive: true });
    }
    await rename(previous, current);
  } else if (!journal.hadCertificate && (await exists(current))) {
    await driver.removeCertificate(current, state.id);
    await rm(current, { recursive: true });
  }
  if (await exists(next)) {
    if (await exists(join(next, 'identity.json'))) await driver.removeCertificate(next, state.id);
    await rm(next, { recursive: true });
  }
  const restored = { ...state, active: journal.previous };
  await writeInstallation(root, restored);
  if (restored.active) {
    await driver.start(restored);
    await waitReady(root, restored);
    await driver.registerUninstaller();
  } else await driver.removeService();
  await rm(path);
}

async function withInstallationLock(root, operation) {
  const path = join(root, 'install.lock');
  let handle;
  try {
    handle = await open(path, 'wx', 0o600);
  } catch (error) {
    if (error.code !== 'EEXIST') throw error;
    await assertRegularOwned(path);
    const pid = Number(await readFile(path, 'utf8'));
    if (!Number.isSafeInteger(pid) || pid <= 0)
      throw new Error('安装锁无效，请检查是否有其他安装窗口仍在运行。');
    try {
      process.kill(pid, 0);
      throw new Error('另一个安装或卸载窗口仍在运行，请先完成该操作。');
    } catch (failure) {
      if (failure.code !== 'ESRCH') throw failure;
    }
    await rm(path);
    handle = await open(path, 'wx', 0o600);
  }
  try {
    await handle.writeFile(String(process.pid));
    await handle.close();
    handle = undefined;
    await operation();
  } finally {
    await handle?.close();
    await rm(path, { force: true });
  }
}

async function install(root, driver) {
  console.log('1/5 正在校验安装包及自带运行环境…');
  const metadata = await verifyPackage(packageRoot);
  const previous = await readInstallation(root);
  await driver.preflight();
  if (!previous.active && (await portOpen()))
    throw new Error(
      '3000 端口已被使用。请关闭此前手动启动的 TrioBridge 服务或其他占用程序，然后重新双击 Install。',
    );
  const active = await stagePackage(packageRoot, root, metadata);
  const version = join(root, 'versions', active);
  await execute(
    join(version, process.platform === 'win32' ? 'runtime/node.exe' : 'runtime/node'),
    [
      '--input-type=module',
      '-e',
      "const binding=await import('@napi-rs/keyring');if(typeof binding.AsyncEntry!=='function')process.exit(1)",
    ],
    { cwd: join(version, 'app'), windowsHide: true, timeout: 20_000 },
  );
  const certificate = join(root, 'certificates');
  const hadCertificate = await exists(certificate);
  const identity = hadCertificate ? await certificateIdentity(certificate, previous.id) : undefined;
  const renew = !identity || identity.expiresAt < Date.now() + 30 * 24 * 60 * 60 * 1000;
  const journal = {
    id: previous.id,
    previous: previous.active,
    hadCertificate,
    nextCertificate: `.cert-next-${randomUUID()}`,
    previousCertificate: `.cert-before-${randomUUID()}`,
  };
  await writeJournal(root, journal);
  try {
    console.log('2/5 正在配置本机 HTTPS；首次安装请批准系统证书信任提示…');
    if (renew) {
      const next = join(root, journal.nextCertificate);
      await driver.generateCertificate(next, previous.id);
      await driver.trustCertificate(next, previous.id);
    } else await driver.trustCertificate(certificate, previous.id);
    console.log('3/5 正在启动随登录运行的本机服务…');
    await stopAndWait(driver);
    if (renew) {
      if (hadCertificate) await rename(certificate, join(root, journal.previousCertificate));
      await rename(join(root, journal.nextCertificate), certificate);
    }
    const nextState = { ...previous, active };
    await writeInstallation(root, nextState);
    await driver.start(nextState);
    await waitForHealth(root, nextState, metadata.version);
    const directory = sharedModelsDirectory();
    const capability = await getOrCreateLocalCapability(directory);
    await requestLocal(root, '/local-models/state', capability);
    console.log('4/5 正在为 Word、Excel、PowerPoint 安装当前账号专用的共享入口…');
    const manifest = renderLocalManifest(
      await readFile(join(version, 'app/manifest.xml'), 'utf8'),
      capability,
    );
    const manifestPath = join(directory, 'manifest/triobridge-ai.xml');
    await writePrivateLocalManifest(manifestPath, manifest);
    await driver.installManifest(manifestPath);
    await driver.registerUninstaller();
    console.log('5/5 本机共享服务与配置接口均已通过启动检查。');
    // Commit before optional cleanup: a cleanup failure must not roll back a working upgrade.
    await rm(join(root, 'transaction.json'));
    if (renew && hadCertificate) {
      const oldCertificate = join(root, journal.previousCertificate);
      try {
        await driver.removeCertificate(oldCertificate, previous.id);
        await rm(oldCertificate, { recursive: true });
      } catch {
        console.warn('新服务已可用；旧证书清理未完成，可在下一次修复安装时处理。');
      }
    }
    for (const name of await readdir(join(root, 'versions'))) {
      if (!VERSION_DIRECTORY.test(name) || [active, previous.active].includes(name)) continue;
      try {
        const directory = join(root, 'versions', name);
        await assertRegularOwned(directory, 'directory');
        const old = JSON.parse(await readFile(join(directory, 'package-manifest.json'), 'utf8'));
        if (old.product === 'triobridge-ai-local') await rm(directory, { recursive: true });
      } catch {
        console.warn('部分旧版本文件仍保留，不影响新版本使用。');
      }
    }
    console.log('安装完成。请保存文档并完全退出、重开 Office，选择“三桥 AI（本机共享）”。');
    console.log(
      '模型设置页应显示“本机共享”。本地 API 可填写 http://127.0.0.1:8001/v1；无需密钥时选择“不使用密钥”。',
    );
    console.log('本机服务会随当前账号登录启动，无需保持此窗口打开。');
  } catch (error) {
    try {
      await recoverInstallation(root, driver);
    } catch {
      console.error(
        '自动恢复未完成，请保留安装目录并重新双击 Install 修复。模型配置和密钥未删除。',
      );
    }
    throw error;
  }
}

export async function runInstaller(action) {
  if (!['install', 'remove', 'restart'].includes(action))
    throw new Error('支持 install、remove 或 restart。');
  const root = installedRoot();
  if (action === 'install') await initializeInstallation(root);
  else if (!(await exists(root))) {
    console.log('未找到本机共享版安装。模型配置和密钥未修改。');
    return;
  } else await readInstallation(root);
  const driver = platformDriver(root);
  await withInstallationLock(root, async () => {
    await recoverInstallation(root, driver);
    if (action === 'install') return install(root, driver);
    await driver.preflight();
    const state = await readInstallation(root);
    if (action === 'restart') {
      if (!state.active) throw new Error('安装尚未完成，请重新双击 Install。');
      await stopAndWait(driver);
      await driver.start(state);
      await waitForHealth(root, state);
      console.log('本机共享服务已重启，原有模型配置和密钥已保留。');
      return;
    }
    console.log('正在停止本机服务并移除共享入口…');
    await driver.stop();
    await driver.removeManifest();
    await driver.removeService();
    for (const name of await readdir(root)) {
      if (name !== 'certificates' && !/^\.cert-(?:before|next)-[a-f0-9-]{36}$/.test(name)) continue;
      const directory = join(root, name);
      if (await exists(join(directory, 'identity.json')))
        await driver.removeCertificate(directory, state.id);
    }
    await rm(root, { recursive: true });
    console.log('本机共享服务、安装文件和本次安装的证书已移除。模型配置与系统密钥已保留。');
    console.log('请保存文档并完全退出、重开 Office。');
  });
}

if (process.argv[1] && resolve(process.argv[1]) === fileURLToPath(import.meta.url)) {
  try {
    await runInstaller(process.argv[2]);
  } catch (error) {
    console.error(
      `操作未完成：${error instanceof Error ? error.message : '请检查安装环境后重试。'}`,
    );
    console.error('可重新双击 Install 修复；Uninstall 可移除安装组件。模型配置和密钥会保留。');
    process.exitCode = 1;
  }
}
