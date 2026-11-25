import { createHash } from 'node:crypto';
import { execFileSync } from 'node:child_process';
import { createWriteStream } from 'node:fs';
import { mkdir, readFile, writeFile } from 'node:fs/promises';
import { dirname, join } from 'node:path';
import { pipeline } from 'node:stream/promises';
import { unzipSync } from 'fflate';

const version = '1.0.4370.26';
const sdkSha256 = '1c0dacac84b70109b3fd42f90c53a0ba942abc3a87590d7a4e0ccd86e4a06cf4';
const managed = ['Microsoft.Web.WebView2.Core.dll', 'Microsoft.Web.WebView2.WinForms.dll'];

export async function stageWebViewSdk(root, pane) {
  const cache = join(root, '.cache/webview2', 'Microsoft.Web.WebView2.' + version + '.nupkg');
  const digest = (data) => createHash('sha256').update(data).digest('hex');
  let pkg = await readFile(cache).catch((e) => {
    if (e.code !== 'ENOENT') throw e;
  });
  if (!pkg || digest(pkg) !== sdkSha256) {
    const url =
      'https://api.nuget.org/v3-flatcontainer/microsoft.web.webview2/' +
      version +
      '/microsoft.web.webview2.' +
      version +
      '.nupkg';
    const response = await fetch(url, { signal: AbortSignal.timeout(120000) });
    if (!response.ok) throw Error('Unable to download the WebView2 SDK: ' + response.status);
    pkg = Buffer.from(await response.arrayBuffer());
    if (digest(pkg) !== sdkSha256) throw Error('WebView2 SDK checksum mismatch.');
    await mkdir(dirname(cache), { recursive: true });
    await writeFile(cache, pkg);
  }
  const files = new Map(managed.map((name) => ['lib/net462/' + name, name]));
  for (const arch of ['x86', 'x64', 'arm64']) {
    const path = 'runtimes/win-' + arch + '/native/WebView2Loader.dll';
    files.set(path, path);
  }
  files.set('LICENSE.txt', 'WebView2-LICENSE.txt');
  files.set('NOTICE.txt', 'WebView2-NOTICE.txt');
  const contents = unzipSync(pkg, { filter: (file) => files.has(file.name) });
  for (const [from, to] of files) {
    if (!contents[from]) throw Error('Missing WebView2 SDK file: ' + from);
    await mkdir(dirname(join(pane, to)), { recursive: true });
    await writeFile(join(pane, to), contents[from]);
  }
}

export async function addWindowsTaskPane({ root, payload, dest, arch }) {
  const pane = join(dest, 'pane');
  await stageWebViewSdk(root, pane);
  const refs = [
    'System',
    'System.Core',
    'System.Drawing',
    'System.Windows.Forms',
    'System.Web.Extensions',
    'Microsoft.CSharp',
  ];
  execFileSync(
    join(process.env.SystemRoot, 'Microsoft.NET/Framework/v4.0.30319/csc.exe'),
    [
      '/nologo',
      '/codepage:65001',
      '/optimize+',
      '/platform:anycpu',
      '/target:library',
      '/out:' + join(pane, 'TrioBridge.TaskPane.dll'),
      ...refs.map((r) => '/reference:' + r + '.dll'),
      ...managed.map((file) => '/reference:' + join(pane, file)),
      join(root, 'offline/windows/Office.cs'),
      join(root, 'offline/windows/TaskPane.cs'),
    ],
    { stdio: 'inherit' },
  );

  // Ship Microsoft's full per-OS installer. Office and the application may use
  // different CPU architectures. The loader above follows the Office process.
  const file = 'MicrosoftEdgeWebView2RuntimeInstaller' + arch.toUpperCase() + '.exe';
  const installer = join(payload, file);
  const url =
    'https://go.microsoft.com/fwlink/?linkid=' + { x64: '2124701', arm64: '2099616' }[arch];
  const runtime = await fetch(url, { signal: AbortSignal.timeout(300000) });
  if (
    !runtime.ok ||
    !new URL(runtime.url).pathname.toLowerCase().endsWith('/' + file.toLowerCase())
  )
    throw Error('Expected the Microsoft offline WebView2 installer for ' + arch);
  await pipeline(runtime.body, createWriteStream(installer, { flags: 'wx' }));
  execFileSync(
    'powershell.exe',
    [
      '-NoProfile',
      '-NonInteractive',
      '-Command',
      [
        '$signature = Get-AuthenticodeSignature -LiteralPath $env:TRIO_WEBVIEW_INSTALLER',
        'if ($signature.Status -ne "Valid" -or $signature.SignerCertificate.GetNameInfo("SimpleName", $false) -ne "Microsoft Corporation") { throw "Invalid WebView2 installer signature." }',
      ].join('\n'),
    ],
    { env: { ...process.env, TRIO_WEBVIEW_INSTALLER: installer }, stdio: 'inherit' },
  );
}
