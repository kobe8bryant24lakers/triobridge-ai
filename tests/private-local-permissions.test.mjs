import { test } from 'node:test';
import assert from 'node:assert/strict';
import { mkdtemp, rm, stat, symlink, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import {
  enforcePrivatePath,
  protectWindowsPath,
  WINDOWS_PRIVATE_ACL_SCRIPT,
  windowsPowerShellEnvironment,
} from '../scripts/private-local-permissions.mjs';

test('Windows PowerShell reconstructs module paths inherited through Node from PowerShell 7', () => {
  const environment = {
    SystemRoot: 'C:\\Windows',
    PSModulePath: 'incompatible-v7',
    psMODULEpath: 'v7',
    OTHER: 'keep',
  };
  assert.deepEqual(windowsPowerShellEnvironment(environment), {
    SystemRoot: 'C:\\Windows',
    OTHER: 'keep',
  });
  assert.equal(environment.PSModulePath, 'incompatible-v7');
});

test('Windows private ACL invocation uses a fixed script and passes only a literal path through environment', async () => {
  const path = "C:\\Users\\Synthetic User\\a'; Write-Host PWNED; #\\模型配置";
  const calls = [];
  await protectWindowsPath(path, 'directory', {
    environment: { SystemRoot: 'C:\\Windows' },
    runner: async (...args) => calls.push(args),
  });
  assert.equal(calls.length, 1);
  assert.equal(calls[0][0], 'C:\\Windows\\System32\\WindowsPowerShell\\v1.0\\powershell.exe');
  assert.deepEqual(calls[0][1], [
    '-NoProfile',
    '-NonInteractive',
    '-EncodedCommand',
    Buffer.from(WINDOWS_PRIVATE_ACL_SCRIPT, 'utf16le').toString('base64'),
  ]);
  assert.equal(calls[0][1].join(' ').includes(path), false);
  assert.equal(calls[0][2].env.TRIOBRIDGE_PRIVATE_PATH, path);
  assert.equal(calls[0][2].env.TRIOBRIDGE_PRIVATE_KIND, 'directory');
  assert.equal(calls[0][2].env.TRIOBRIDGE_PRIVATE_CREATED, '0');
  assert.equal(calls[0][2].windowsHide, true);
  assert.equal(calls[0][2].timeout, 15_000);
  assert.match(WINDOWS_PRIVATE_ACL_SCRIPT, /GetCurrent\(\)/);
  assert.match(WINDOWS_PRIVATE_ACL_SCRIPT, /SetAccessRuleProtection\(\$true, \$false\)/);
  assert.match(WINDOWS_PRIVATE_ACL_SCRIPT, /'S-1-5-18', 'S-1-5-32-544'/);
  assert.match(WINDOWS_PRIVATE_ACL_SCRIPT, /GetOwner/);
  assert.match(WINDOWS_PRIVATE_ACL_SCRIPT, /GetAccessRules/);
  assert.equal(WINDOWS_PRIVATE_ACL_SCRIPT.includes('apiKey'), false);
  assert.equal(WINDOWS_PRIVATE_ACL_SCRIPT.includes('local-token'), false);
});

test('only explicitly new files may adopt the elevated Windows creator group', async () => {
  const calls = [];
  const environment = { SystemRoot: 'C:\\Windows', TRIOBRIDGE_PRIVATE_CREATED: '1' };
  const runner = async (...args) => calls.push(args);
  await protectWindowsPath('C:\\existing', 'file', { environment, runner });
  await protectWindowsPath('C:\\new', 'file', { environment, runner, created: true });
  assert.equal(calls[0][2].env.TRIOBRIDGE_PRIVATE_CREATED, '0');
  assert.equal(calls[1][2].env.TRIOBRIDGE_PRIVATE_CREATED, '1');
  assert.match(WINDOWS_PRIVATE_ACL_SCRIPT, /\$fresh -and \$owner -eq 'S-1-5-32-544'/);
  assert.match(WINDOWS_PRIVATE_ACL_SCRIPT, /IsInRole/);
});

test('Windows permission failures are sanitized and never silently fall back to chmod', async () => {
  await assert.rejects(
    protectWindowsPath('C:\\private', 'file', {
      environment: { SystemRoot: 'C:\\Windows' },
      runner: async () => {
        throw new Error('private process environment detail');
      },
    }),
    (error) => {
      assert.equal(
        error.message,
        'Could not enforce private permissions for the current Windows account.',
      );
      assert.equal(error.message.includes('environment detail'), false);
      return true;
    },
  );
  let calls = 0;
  await assert.rejects(
    protectWindowsPath('C:\\private', 'file', { environment: {}, runner: async () => calls++ }),
  );
  assert.equal(calls, 0);
});

test(
  'native private-path permissions reject links and constrain POSIX files',
  { skip: process.platform === 'win32' },
  async (t) => {
    const root = await mkdtemp(join(tmpdir(), 'triobridge-permissions-'));
    t.after(() => rm(root, { recursive: true, force: true }));
    const file = join(root, 'private');
    await writeFile(file, 'synthetic');
    await enforcePrivatePath(root, 'directory');
    await enforcePrivatePath(file, 'file');
    assert.equal((await stat(root)).mode & 0o777, 0o700);
    assert.equal((await stat(file)).mode & 0o777, 0o600);
    const linked = join(root, 'linked');
    await symlink(file, linked);
    await assert.rejects(enforcePrivatePath(linked, 'file'));
  },
);
