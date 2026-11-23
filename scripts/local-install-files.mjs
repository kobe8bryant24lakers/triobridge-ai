import { createHash, randomUUID } from 'node:crypto';
import {
  chmod,
  copyFile,
  lstat,
  mkdir,
  readFile,
  readdir,
  rename,
  rm,
  writeFile,
} from 'node:fs/promises';
import { homedir } from 'node:os';
import { dirname, join, resolve, sep } from 'node:path';
import { enforcePrivatePath } from './private-local-permissions.mjs';

export const INSTALL_PRODUCT = 'triobridge-ai-local';
export const VERSION_DIRECTORY = /^\d+\.\d+\.\d+-[a-f0-9]{12}(?:-[a-f0-9]{8})?$/;

export function installedRoot(platform = process.platform) {
  if (platform === 'darwin')
    return join(homedir(), 'Library/Application Support/TrioBridge AI/local-app');
  if (platform === 'win32' && process.env.LOCALAPPDATA)
    return join(process.env.LOCALAPPDATA, 'TrioBridgeAI/LocalApp');
  throw new Error('本机共享安装包支持 macOS 和 Windows。');
}

export async function assertRegularOwned(path, kind = 'file') {
  const info = await lstat(path);
  if (
    info.isSymbolicLink() ||
    (kind === 'directory' ? !info.isDirectory() : !info.isFile()) ||
    (process.getuid && info.uid !== process.getuid())
  )
    throw new Error('安装路径不是当前账号拥有的普通文件或目录，已停止操作。');
  return info;
}

export async function readInstallation(root) {
  await assertRegularOwned(root, 'directory');
  await assertRegularOwned(join(root, 'installation.json'));
  const state = JSON.parse(await readFile(join(root, 'installation.json'), 'utf8'));
  if (
    state.product !== INSTALL_PRODUCT ||
    state.format !== 1 ||
    !/^[a-f0-9-]{36}$/.test(state.id) ||
    (state.active !== null && !VERSION_DIRECTORY.test(state.active))
  )
    throw new Error('安装记录无效，未修改已有文件。');
  return state;
}

export async function writeInstallation(root, state) {
  const temporary = join(root, `.installation-${randomUUID()}.tmp`);
  try {
    await writeFile(temporary, JSON.stringify(state, null, 2) + '\n', { mode: 0o600, flag: 'wx' });
    await enforcePrivatePath(temporary, 'file', { created: true });
    await rename(temporary, join(root, 'installation.json'));
    await enforcePrivatePath(join(root, 'installation.json'), 'file');
  } finally {
    await rm(temporary, { force: true });
  }
}

export async function initializeInstallation(root) {
  // Refuse linked ancestors before creating any private installation files.
  const ancestors = [];
  for (let current = root; current !== dirname(current); current = dirname(current))
    ancestors.push(current);
  for (const path of ancestors.reverse()) {
    try {
      const info = await lstat(path);
      if (info.isSymbolicLink() || !info.isDirectory())
        throw new Error('安装目录包含链接或非目录路径。');
    } catch (error) {
      if (error.code !== 'ENOENT') throw error;
    }
  }
  try {
    return await readInstallation(root);
  } catch (error) {
    if (error.code !== 'ENOENT') throw error;
    try {
      if ((await readdir(root)).length) throw new Error('安装目录已有其他内容，未覆盖。');
    } catch (failure) {
      if (failure.code !== 'ENOENT') throw failure;
    }
  }
  await mkdir(root, { recursive: true, mode: 0o700 });
  await enforcePrivatePath(root, 'directory', { created: true });
  const state = { format: 1, product: INSTALL_PRODUCT, id: randomUUID(), active: null };
  await writeInstallation(root, state);
  return state;
}

export async function verifyPackage(directory, platform = process.platform, arch = process.arch) {
  await assertRegularOwned(join(directory, 'package-manifest.json'));
  const raw = await readFile(join(directory, 'package-manifest.json'));
  const metadata = JSON.parse(raw);
  if (
    metadata.product !== INSTALL_PRODUCT ||
    metadata.format !== 1 ||
    metadata.platform !== platform ||
    metadata.arch !== arch ||
    !/^\d+\.\d+\.\d+$/.test(metadata.version) ||
    !Array.isArray(metadata.files) ||
    !metadata.files.length ||
    metadata.files.length > 2000
  )
    throw new Error('安装包不适用于当前系统或架构，或安装包信息无效。');
  const seen = new Set();
  for (const file of metadata.files) {
    if (
      typeof file.path !== 'string' ||
      !/^[A-Za-z0-9@_./-]+$/.test(file.path) ||
      file.path.split('/').some((part) => !part || part === '.' || part === '..') ||
      seen.has(file.path) ||
      !/^[a-f0-9]{64}$/.test(file.sha256) ||
      !Number.isSafeInteger(file.size) ||
      file.size < 0
    )
      throw new Error('安装包文件清单无效。');
    seen.add(file.path);
    const path = resolve(directory, file.path);
    if (!path.startsWith(resolve(directory) + sep)) throw new Error('安装包路径越界。');
    for (let parent = dirname(path); parent !== resolve(directory); parent = dirname(parent))
      await assertRegularOwned(parent, 'directory');
    const info = await assertRegularOwned(path);
    const data = await readFile(path);
    if (info.size !== file.size || createHash('sha256').update(data).digest('hex') !== file.sha256)
      throw new Error(`安装包文件损坏，请重新下载：${file.path}`);
  }
  for (const required of [
    'app/dist/index.html',
    'app/scripts/serve-installed.mjs',
    'app/scripts/install-local.mjs',
    'app/manifest.xml',
    platform === 'win32' ? 'runtime/node.exe' : 'runtime/node',
    ...(platform === 'win32' ? ['app/scripts/start-installed.ps1'] : []),
  ])
    if (!seen.has(required)) throw new Error('安装包缺少必要组件。');
  return { ...metadata, digest: createHash('sha256').update(raw).digest('hex') };
}

export async function stagePackage(source, root, metadata) {
  const versions = join(root, 'versions');
  const created = await mkdir(versions, { recursive: true, mode: 0o700 });
  await enforcePrivatePath(versions, 'directory', { created: created !== undefined });
  let name = `${metadata.version}-${metadata.digest.slice(0, 12)}`;
  try {
    const previous = await verifyPackage(join(versions, name));
    if (previous.digest === metadata.digest) return name;
  } catch (error) {
    if (error.code !== 'ENOENT') name += `-${randomUUID().replaceAll('-', '').slice(0, 8)}`;
  }
  const temporary = join(versions, `.stage-${randomUUID()}`);
  try {
    await mkdir(temporary, { mode: 0o700 });
    for (const file of metadata.files) {
      const target = join(temporary, file.path);
      await mkdir(dirname(target), { recursive: true, mode: 0o700 });
      await copyFile(join(source, file.path), target);
      if (process.platform !== 'win32') await chmod(target, file.executable ? 0o755 : 0o600);
    }
    await copyFile(join(source, 'package-manifest.json'), join(temporary, 'package-manifest.json'));
    await verifyPackage(temporary);
    await rename(temporary, join(versions, name));
    return name;
  } finally {
    await rm(temporary, { recursive: true, force: true });
  }
}
