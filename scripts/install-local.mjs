import { execFile } from 'node:child_process';
import { promisify } from 'node:util';
import { randomUUID } from 'node:crypto';
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
import {
  checkLoopbackPort,
  ReadinessError,
  waitForLocalEndpoint,
} from './local-service-readiness.mjs';
import { platformDriver, certificateIdentity } from './local-install-platform.mjs';
import { getOrCreateLocalCapability } from './local-capability.mjs';
import { storageDetailMessage } from './model-storage-diagnostic.mjs';
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

async function waitForHealth(root, state, version) {
  const path = join(root, 'certificates/server.crt');
  await assertRegularOwned(path);
  const ca = await readFile(path);
  await waitForLocalEndpoint({ ca, installationId: state.id, version });
  return ca;
}

const startupPhases = {
  stop: ['正在等待旧服务释放 3000 端口', 'TASK_STOP'],
  activate: ['正在启用本次安装的证书和运行文件', 'ACTIVATION'],
  credential: ['正在准备当前账号的共享配置访问凭据', 'CAPABILITY_ACCESS'],
  start: ['正在启动当前账号的登录任务', 'TASK_START'],
  https: ['正在验证 localhost:3000 的 HTTPS 和安装标识', 'CERTIFICATE_READ'],
  models: ['正在检查本机共享模型配置', 'MODEL_STORAGE'],
};

async function saveStartupResult(root, report) {
  if (process.platform !== 'win32') return;
  const temporary = join(root, `.startup-${randomUUID()}.tmp`);
  try {
    // Only fixed phase/error codes and scheduler status are kept, even after rollback.
    await writeFile(temporary, JSON.stringify(report), { flag: 'wx', mode: 0o600 });
    await rename(temporary, join(root, 'last-startup.json'));
  } catch {
    console.warn('本次启动诊断未写入，安装窗口中的阶段和代码仍可用于排查。');
  } finally {
    await rm(temporary, { force: true }).catch(() => {});
  }
}

async function startAndVerify(root, driver, state, version, activate) {
  let phase;
  let since;
  function progress(next) {
    phase = next;
    since = Date.now();
    console.log('  ' + startupPhases[phase][0] + '…');
  }
  const heartbeat = setInterval(() => {
    console.log(
      `  ${startupPhases[phase][0]}，已等待 ${Math.floor((Date.now() - since) / 1000)} 秒。`,
    );
  }, 10_000);
  try {
    progress('stop');
    await stopAndWait(driver);
    await checkLoopbackPort();
    if (activate) {
      progress('activate');
      await activate();
    }
    // Finish protecting the directory before the new process can open it on first launch.
    progress('credential');
    const capability = await getOrCreateLocalCapability(sharedModelsDirectory());
    progress('start');
    await driver.start(state);
    progress('https');
    const ca = await waitForHealth(root, state, version);
    progress('models');
    await waitForLocalEndpoint({ ca, token: capability });
    await saveStartupResult(root, { phase: 'ready', code: 'OK' });
    return capability;
  } catch (error) {
    const code = error instanceof ReadinessError ? error.code : startupPhases[phase][1];
    let task;
    if (driver.taskStatus) task = await driver.taskStatus().catch(() => undefined);
    await saveStartupResult(root, {
      phase,
      code,
      task,
      hresult: error.taskHresult,
      storageDetail: error.storageDetail,
    });
    let detail = startupPhases[phase][0] + '时出错。';
    if (error instanceof ReadinessError) detail = error.message;
    else if (storageDetailMessage(error.storageDetail))
      detail += ` [${error.storageDetail}] ${storageDetailMessage(error.storageDetail)}`;
    const taskMessage =
      task?.state === 'Running' && task.result === '0x00041301'
        ? ' Windows 登录任务正在运行（0x00041301）。'
        : task
          ? ` Windows 登录任务 ${task.state}，结果 ${task.result ?? '未知'}。`
          : '';
    throw new Error(
      (phase === 'models' ? '本机共享配置检查未通过。' : '本机服务未能启动。') +
        `[${phase}/${code}] ${detail}` +
        (error.taskHresult ? ` Windows 错误码 ${error.taskHresult}。` : '') +
        taskMessage +
        (process.platform === 'win32' ? ' 可双击包内 Diagnose.cmd 查看本次启动诊断。' : ''),
    );
  } finally {
    clearInterval(heartbeat);
  }
}

async function stopAndWait(driver) {
  await driver.stop();
  for (let attempt = 0; attempt < 20; attempt++) {
    if (!(await portOpen())) return;
    await delay(200);
  }
  throw new ReadinessError('PORT_BUSY');
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
    // Recovery restores the previous executable. An existing catalog error must not
    // leave a pending journal that blocks the user's next repair or uninstall.
    await waitReady(root, restored, restored.active.split('-')[0]);
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
  if (process.platform === 'darwin')
    console.log('首次访问 Office 目录时，请留意并处理 macOS 的数据访问提示。');
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
    const nextState = { ...previous, active };
    const capability = await startAndVerify(root, driver, nextState, metadata.version, async () => {
      if (renew) {
        if (hadCertificate) await rename(certificate, join(root, journal.previousCertificate));
        await rename(join(root, journal.nextCertificate), certificate);
      }
      await writeInstallation(root, nextState);
    });
    const directory = sharedModelsDirectory();
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
    console.error(error.message);
    console.log('正在恢复安装前的状态…');
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

export async function runInstaller(action, { driver: injectedDriver } = {}) {
  if (!['install', 'remove', 'restart'].includes(action))
    throw new Error('支持 install、remove 或 restart。');
  const root = installedRoot();
  if (action === 'install') await initializeInstallation(root);
  else if (!(await exists(root))) {
    console.log('未找到本机共享版安装。模型配置和密钥未修改。');
    return;
  } else await readInstallation(root);
  const driver = injectedDriver || platformDriver(root);
  await withInstallationLock(root, async () => {
    await recoverInstallation(root, driver);
    if (action === 'install') return install(root, driver);
    await driver.preflight(action);
    const state = await readInstallation(root);
    if (action === 'restart') {
      if (!state.active) throw new Error('安装尚未完成，请重新双击 Install。');
      await startAndVerify(root, driver, state, state.active.split('-')[0]);
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
