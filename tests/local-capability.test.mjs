import { test } from 'node:test';
import assert from 'node:assert/strict';
import { mkdtemp, readFile, rm, stat, symlink, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import {
  getOrCreateLocalCapability,
  matchesLocalCapability,
} from '../scripts/local-capability.mjs';

test('installation capability is stable, 256-bit, private, and atomically provisioned', async (t) => {
  const root = await mkdtemp(join(tmpdir(), 'triobridge-capability-'));
  t.after(() => rm(root, { recursive: true, force: true }));
  const directory = join(root, 'private');
  const results = await Promise.all(
    Array.from({ length: 8 }, () => getOrCreateLocalCapability(directory)),
  );
  assert.equal(new Set(results).size, 1);
  assert.match(results[0], /^[a-f0-9]{64}$/);
  assert.equal((await stat(directory)).mode & 0o777, 0o700);
  assert.equal((await stat(join(directory, 'local-capability'))).mode & 0o777, 0o600);
  assert.equal(await readFile(join(directory, 'local-capability'), 'utf8'), results[0]);
  assert.equal(matchesLocalCapability(results[0], results[0]), true);
  for (const candidate of [undefined, [], '', 'short', '0'.repeat(64)])
    assert.equal(matchesLocalCapability(candidate, results[0]), false);
});

test('linked and corrupt capabilities fail closed without being replaced', async (t) => {
  const root = await mkdtemp(join(tmpdir(), 'triobridge-capability-invalid-'));
  t.after(() => rm(root, { recursive: true, force: true }));
  const target = join(root, 'unrelated');
  const capability = join(root, 'local-capability');
  await writeFile(target, 'unrelated contents');
  await symlink(target, capability);
  await assert.rejects(getOrCreateLocalCapability(root));
  assert.equal(await readFile(target, 'utf8'), 'unrelated contents');
  await rm(capability);
  await writeFile(capability, 'invalid');
  await assert.rejects(getOrCreateLocalCapability(root));
  assert.equal(await readFile(capability, 'utf8'), 'invalid');
});
