import { readFile, writeFile, rename, lstat, mkdir, unlink } from 'node:fs/promises';
import { dirname, join } from 'node:path';
import { randomUUID } from 'node:crypto';
import { enforcePrivatePath } from './private-local-permissions.mjs';

export const LOCAL_MANIFEST_ID = 'd7cf6e8b-6d03-4a54-9c31-5b7d5427618e';

export function renderLocalManifest(template, capability) {
  if (!/^[a-f0-9]{64}$/.test(capability)) throw new Error('Invalid local installation capability.');
  if (!template.includes(`<Id>${LOCAL_MANIFEST_ID}</Id>`))
    throw new Error('The manifest is not the TrioBridge local app.');
  const entries = [...template.matchAll(/<SourceLocation\s+DefaultValue="([^"]+)"\s*\/>/g)];
  if (entries.length !== 1) throw new Error('Expected one local app source location.');
  const url = new URL(entries[0][1].replaceAll('&amp;', '&'));
  if (
    url.origin !== 'https://localhost:3000' ||
    url.pathname !== '/' ||
    url.hash ||
    url.search !== '?office=1'
  )
    throw new Error('The local app must use its fixed HTTPS loopback origin.');
  url.hash = `local-token=${capability}`;
  return template
    .replace(
      entries[0][0],
      `<SourceLocation DefaultValue="${url.href.replaceAll('&', '&amp;')}" />`,
    )
    .replace(
      'DisplayName DefaultValue="TrioBridge AI · 三桥 AI"',
      'DisplayName DefaultValue="TrioBridge AI · 三桥 AI（本机共享）"',
    );
}

export async function preflightLocalManifest(destination) {
  try {
    const parent = await lstat(dirname(destination));
    if (
      !parent.isDirectory() ||
      parent.isSymbolicLink() ||
      (process.getuid && parent.uid !== process.getuid())
    )
      throw new Error('The local manifest directory is not owned by the current user.');
  } catch (error) {
    if (error.code !== 'ENOENT') throw error;
  }
  try {
    const info = await lstat(destination);
    if (
      !info.isFile() ||
      info.isSymbolicLink() ||
      (process.getuid && info.uid !== process.getuid())
    )
      throw new Error('Refusing to replace an unsafe local manifest.');
    if (!(await readFile(destination, 'utf8')).includes(`<Id>${LOCAL_MANIFEST_ID}</Id>`))
      throw new Error('A different app owns the destination manifest.');
  } catch (error) {
    if (error.code !== 'ENOENT') throw error;
  }
}

/** These files contain an installation capability, not a model key; keep them private. */
export async function writePrivateLocalManifest(destination, manifest) {
  await preflightLocalManifest(destination);
  await mkdir(dirname(destination), { recursive: true });
  await enforcePrivatePath(dirname(destination), 'directory');
  const temporary = join(dirname(destination), `.triobridge-${randomUUID()}.tmp`);
  try {
    await writeFile(temporary, manifest, { mode: 0o600, flag: 'wx' });
    await rename(temporary, destination);
    await enforcePrivatePath(destination, 'file');
  } finally {
    await unlink(temporary).catch(() => {});
  }
}
