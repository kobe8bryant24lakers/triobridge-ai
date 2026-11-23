import test from 'node:test';
import assert from 'node:assert/strict';
import { mkdtemp, readFile, stat, symlink, writeFile, rm } from 'node:fs/promises';
import { join } from 'node:path';
import { tmpdir } from 'node:os';
import {
  preflightLocalManifest,
  renderLocalManifest,
  writePrivateLocalManifest,
} from '../scripts/local-manifest.mjs';

const template = await readFile(new URL('../manifest.xml', import.meta.url), 'utf8');
const token = 'a'.repeat(64);

test('the private local manifest bootstraps the same capability without changing app identity or hosts', () => {
  const result = renderLocalManifest(template, token);
  assert.match(result, new RegExp(`https://localhost:3000/\\?office=1#local-token=${token}`));
  for (const host of ['Document', 'Workbook', 'Presentation'])
    assert.ok(result.includes(`<Host Name="${host}" />`));
  assert.equal(result.match(/<Id>([^<]+)<\/Id>/)[1], template.match(/<Id>([^<]+)<\/Id>/)[1]);
  assert.ok(!template.includes(token));
  assert.throws(() => renderLocalManifest(template, 'short'));
  assert.throws(() => renderLocalManifest(template.replace('?office=1', '?demo=1'), token));
  assert.throws(() =>
    renderLocalManifest(
      template.replaceAll('https://localhost:3000', 'https://example.com'),
      token,
    ),
  );
});

test('capability manifests are private and refuse foreign files or symlink destinations', async () => {
  const dir = await mkdtemp(join(tmpdir(), 'triobridge-local-manifest-'));
  try {
    const target = join(dir, 'triobridge.xml');
    const manifest = renderLocalManifest(template, token);
    await writePrivateLocalManifest(target, manifest);
    assert.equal(await readFile(target, 'utf8'), manifest);
    if (process.platform !== 'win32') assert.equal((await stat(target)).mode & 0o777, 0o600);
    await writeFile(target, '<OfficeApp><Id>different</Id></OfficeApp>');
    await assert.rejects(preflightLocalManifest(target), /different app/);
    await assert.rejects(writePrivateLocalManifest(target, manifest), /different app/);
    if (process.platform !== 'win32') {
      const link = join(dir, 'alias.xml');
      await symlink(target, link);
      await assert.rejects(writePrivateLocalManifest(link, manifest), /unsafe/);
    }
  } finally {
    await rm(dir, { recursive: true, force: true });
  }
});
