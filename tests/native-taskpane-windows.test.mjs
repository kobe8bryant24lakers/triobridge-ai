import { test } from 'node:test';
import assert from 'node:assert/strict';
import { execFileSync } from 'node:child_process';
import { existsSync } from 'node:fs';
import { mkdtemp, readFile, writeFile, rm } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { unzipSync } from 'fflate';

const sdk = fileURLToPath(
  new URL('../.cache/webview2/Microsoft.Web.WebView2.1.0.4370.26.nupkg', import.meta.url),
);

test(
  'Windows task pane compiles against the shipped WebView2 SDK and rejects untrusted requests',
  {
    skip:
      process.platform !== 'win32'
        ? 'Requires the Windows .NET Framework compiler.'
        : !existsSync(sdk) && 'Run the Windows installer build to cache the pinned WebView2 SDK.',
  },
  async () => {
    const directory = await mkdtemp(join(tmpdir(), 'triobridge-pane-test-'));
    try {
      const files = unzipSync(await readFile(sdk));
      const libraries = ['Microsoft.Web.WebView2.Core.dll', 'Microsoft.Web.WebView2.WinForms.dll'];
      for (const library of libraries) {
        await writeFile(join(directory, library), files['lib/net462/' + library]);
      }
      const executable = join(directory, 'TaskPaneTests.exe');
      execFileSync(
        join(process.env.SystemRoot, 'Microsoft.NET/Framework/v4.0.30319/csc.exe'),
        [
          '/nologo',
          '/utf8output',
          '/codepage:65001',
          '/target:exe',
          '/platform:anycpu',
          '/main:TrioBridge.TaskPaneTests',
          '/out:' + executable,
          ...[
            'Microsoft.CSharp.dll',
            'System.Web.Extensions.dll',
            'System.Core.dll',
            'System.Drawing.dll',
            'System.Windows.Forms.dll',
            ...libraries.map((library) => join(directory, library)),
          ].map((library) => '/reference:' + library),
          ...[
            '../offline/windows/Office.cs',
            '../offline/windows/TaskPane.cs',
            './native-taskpane-windows.cs',
          ].map((file) => fileURLToPath(new URL(file, import.meta.url))),
        ],
        { encoding: 'utf8', timeout: 30000 },
      );
      const output = execFileSync(executable, { encoding: 'utf8', timeout: 30000 });
      assert.match(output, /Windows task pane policies passed/);
    } finally {
      await rm(directory, { recursive: true, force: true });
    }
  },
);
