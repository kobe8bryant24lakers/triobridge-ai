import { randomBytes, randomUUID, timingSafeEqual } from 'node:crypto';
import { link, lstat, mkdir, open, readFile, unlink } from 'node:fs/promises';
import { join } from 'node:path';
import { enforcePrivatePath } from './private-local-permissions.mjs';

const preparingDirectories = new Map();

export function ensurePrivateModelDirectory(directory) {
  const preparing = preparingDirectories.get(directory);
  if (preparing) return preparing;
  // Concurrent first opens must see the ACL installed by the caller that created the directory.
  const operation = (async () => {
    const created = await mkdir(directory, { recursive: true, mode: 0o700 });
    await enforcePrivatePath(directory, 'directory', { created: created !== undefined });
  })().finally(() => preparingDirectories.delete(directory));
  preparingDirectories.set(directory, operation);
  return operation;
}

/** Provision once per OS-user installation. Never place this token in logs or command arguments. */
export async function getOrCreateLocalCapability(directory) {
  await ensurePrivateModelDirectory(directory);
  const file = join(directory, 'local-capability');
  const temporary = join(directory, `.capability-${randomUUID()}.tmp`);
  let exists = false;
  try {
    await lstat(file);
    exists = true;
  } catch (error) {
    if (error.code !== 'ENOENT') throw error;
  }
  if (!exists) {
    let handle;
    try {
      handle = await open(temporary, 'wx', 0o600);
      await handle.writeFile(randomBytes(32).toString('hex'));
      await handle.sync();
      await handle.close();
      handle = undefined;
      await enforcePrivatePath(temporary, 'file', { created: true });
      // Atomic create-without-replace: another process never observes a half-written capability.
      try {
        await link(temporary, file);
      } catch (error) {
        if (error.code !== 'EEXIST') throw error;
      }
    } finally {
      await handle?.close().catch(() => {});
      await unlink(temporary).catch(() => {});
    }
  }
  const info = await lstat(file);
  if (
    !info.isFile() ||
    info.isSymbolicLink() ||
    info.size !== 64 ||
    (process.getuid && info.uid !== process.getuid())
  )
    throw new Error('Local capability is invalid or is not owned by the current OS user.');
  await enforcePrivatePath(file, 'file');
  const token = await readFile(file, 'utf8');
  if (!/^[a-f0-9]{64}$/.test(token)) throw new Error('Local capability has an invalid format.');
  return token;
}

export function matchesLocalCapability(candidate, expected) {
  if (
    typeof candidate !== 'string' ||
    typeof expected !== 'string' ||
    !/^[a-f0-9]{64}$/.test(candidate) ||
    !/^[a-f0-9]{64}$/.test(expected)
  )
    return false;
  return timingSafeEqual(Buffer.from(candidate, 'hex'), Buffer.from(expected, 'hex'));
}
