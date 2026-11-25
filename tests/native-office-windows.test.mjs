import { test } from 'node:test';
import assert from 'node:assert/strict';
import { execFileSync } from 'node:child_process';
import { mkdtemp, rm } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { fileURLToPath } from 'node:url';

test(
  'Windows native Office preserves PowerPoint append targets and Word write boundaries',
  {
    skip: process.platform !== 'win32' && 'Requires the Windows .NET Framework compiler.',
  },
  async () => {
    const directory = await mkdtemp(join(tmpdir(), 'triobridge-office-test-'));
    try {
      const executable = join(directory, 'OfficeTests.exe');
      execFileSync(
        join(process.env.SystemRoot, 'Microsoft.NET/Framework/v4.0.30319/csc.exe'),
        [
          '/nologo',
          '/utf8output',
          '/codepage:65001',
          '/target:exe',
          '/platform:anycpu',
          '/main:TrioBridge.WindowsOfficeTests',
          '/out:' + executable,
          '/reference:Microsoft.CSharp.dll',
          '/reference:System.Web.Extensions.dll',
          '/reference:System.Core.dll',
          fileURLToPath(new URL('../offline/windows/Office.cs', import.meta.url)),
          fileURLToPath(new URL('./native-office-windows.cs', import.meta.url)),
        ],
        { encoding: 'utf8', timeout: 30000 },
      );
      const output = execFileSync(executable, { encoding: 'utf8', timeout: 30000 });
      assert.match(output, /Windows PowerPoint insertion: 9 scenarios passed/);
      assert.match(output, /Windows Word writing: 32 scenarios passed/);
      assert.match(output, /Windows pane dispatch: 8 scenarios passed/);
    } finally {
      await rm(directory, { recursive: true, force: true });
    }
  },
);
