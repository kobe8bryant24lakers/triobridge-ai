import assert from 'node:assert/strict';
import { test } from 'node:test';
import { mkdtemp, mkdir, readFile, readdir, rm, symlink, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { dirname, join } from 'node:path';
import {
  APP_URL,
  HOSTED_MANIFEST_ID,
  SITE_FILES,
  copySiteFiles,
  hostedManifest,
} from '../scripts/build-pages.mjs';

test('hosted manifest uses separate identity, HTTPS icons and the Office entry query', async () => {
  const local = await readFile(new URL('../manifest.xml', import.meta.url), 'utf8');
  const hosted = hostedManifest(local);
  assert.match(local, /https:\/\/localhost:3000\/\?office=1/);
  assert.ok(!local.includes(HOSTED_MANIFEST_ID));
  assert.ok(hosted.includes(`<Id>${HOSTED_MANIFEST_ID}</Id>`));
  assert.ok(hosted.includes(`SourceLocation DefaultValue="${APP_URL}?office=1"`));
  assert.ok(hosted.includes(`IconUrl DefaultValue="${APP_URL}icon-32.png"`));
  assert.ok(hosted.includes(`HighResolutionIconUrl DefaultValue="${APP_URL}icon-80.png"`));
  assert.doesNotMatch(hosted, /localhost|127\.0\.0\.1/);
  assert.throws(() => hostedManifest(local.replace('<Id>', '<MissingId>')), /exactly one/);
});

test('site packaging copies only approved files for both languages and refuses symlinks', async () => {
  const root = await mkdtemp(join(tmpdir(), 'triobridge-pages-'));
  try {
    for (const [source] of SITE_FILES) {
      await mkdir(dirname(join(root, source)), { recursive: true });
      await writeFile(join(root, source), 'approved synthetic fixture');
    }
    await writeFile(join(root, 'docs/private-test-data.json'), 'must never be published');
    await writeFile(join(root, 'docs/TESTING.md'), 'not a site asset');
    const output = join(root, 'output');
    await copySiteFiles(root, output);
    assert.deepEqual((await readdir(output)).sort(), [
      'assets',
      'index.html',
      'site.css',
      'zh-CN.html',
    ]);
    assert.deepEqual((await readdir(join(output, 'assets'))).sort(), [
      'TrioBridge-AI-User-Guide-zh-CN.pptx',
      'icon.svg',
      'workbench.png',
    ]);
    await rm(join(root, 'docs/assets/icon.svg'));
    await symlink(join(root, 'docs/private-test-data.json'), join(root, 'docs/assets/icon.svg'));
    await assert.rejects(copySiteFiles(root, output), /regular, explicitly approved/);
  } finally {
    await rm(root, { recursive: true, force: true });
  }
});
