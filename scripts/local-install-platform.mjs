import { execFile } from 'node:child_process';
import { promisify } from 'node:util';
import { X509Certificate } from 'node:crypto';
import { mkdir, readFile, rm, writeFile } from 'node:fs/promises';
import { homedir } from 'node:os';
import { dirname, join, win32 } from 'node:path';
import { fileURLToPath } from 'node:url';
import { enforcePrivatePath, windowsPowerShellEnvironment } from './private-local-permissions.mjs';
import { assertRegularOwned } from './local-install-files.mjs';
import {
  LOCAL_MANIFEST_ID,
  preflightLocalManifest,
  writePrivateLocalManifest,
} from './local-manifest.mjs';

const execute = promisify(execFile);
const label = 'ai.triobridge.local-service';
const xml = (value) =>
  String(value)
    .replaceAll('&', '&amp;')
    .replaceAll('<', '&lt;')
    .replaceAll('>', '&gt;')
    .replaceAll('"', '&quot;');

export async function certificateIdentity(directory, installationId) {
  let path = join(directory, 'server.crt');
  try {
    await assertRegularOwned(path);
  } catch (error) {
    if (error.code !== 'ENOENT') throw error;
    path = join(directory, 'server.cer');
    await assertRegularOwned(path);
  }
  await assertRegularOwned(join(directory, 'identity.json'));
  const certificate = new X509Certificate(await readFile(path));
  const identity = JSON.parse(await readFile(join(directory, 'identity.json'), 'utf8'));
  const thumbprint = certificate.fingerprint.replaceAll(':', '').toUpperCase();
  if (
    identity.installationId !== installationId ||
    identity.thumbprint !== thumbprint ||
    !certificate.checkHost('localhost')
  )
    throw new Error('本机证书的安装归属不匹配，未修改证书。');
  return { ...identity, expiresAt: new Date(certificate.validTo).getTime() };
}

export function platformDriver(root, platform = process.platform) {
  const scripts = dirname(fileURLToPath(import.meta.url));
  const plist = join(homedir(), 'Library/LaunchAgents', `${label}.plist`);
  const service = `gui/${process.getuid?.()}/${label}`;
  const native = async (action, extra = {}) => {
    const systemRoot = process.env.SystemRoot || process.env.SYSTEMROOT;
    if (!systemRoot || !win32.isAbsolute(systemRoot)) throw new Error('Windows 系统目录不可用。');
    const executable = win32.join(systemRoot, 'System32/WindowsPowerShell/v1.0/powershell.exe');
    const args = [
      '-NoProfile',
      '-ExecutionPolicy',
      'Bypass',
      '-File',
      join(scripts, 'native-local.ps1'),
      '-Action',
      action,
      '-Root',
      root,
    ];
    for (const [key, value] of Object.entries(extra)) args.push(`-${key}`, value);
    try {
      const { stdout } = await execute(executable, args, {
        env: windowsPowerShellEnvironment(),
        windowsHide: true,
        timeout: action === 'TaskStatus' ? 15_000 : 180_000,
        maxBuffer: 256_000,
      });
      return stdout;
    } catch (error) {
      if (action === 'Start' || action === 'Stop') {
        const failure = new Error('Windows 登录任务操作失败。');
        failure.taskHresult = String(error.stderr || '').match(
          /TRIOBRIDGE_TASK_ERROR (0x[0-9A-F]{8})\b/,
        )?.[1];
        throw failure;
      }
      throw new Error(
        `Windows 安装步骤 ${action} 未完成。${String(error.stderr || error.stdout || '').slice(-1600)}`,
      );
    }
  };
  async function checkPlist() {
    try {
      await assertRegularOwned(plist);
      const { stdout } = await execute('/usr/bin/plutil', ['-convert', 'json', '-o', '-', plist]);
      const value = JSON.parse(stdout);
      if (
        value.Label !== label ||
        !value.ProgramArguments?.[0]?.startsWith(join(root, 'versions') + '/')
      )
        throw new Error('已有同名后台服务不属于本安装，未修改。');
    } catch (error) {
      if (error.code !== 'ENOENT') throw error;
    }
  }
  return {
    ...(platform === 'win32' && {
      async taskStatus() {
        const { state, result } = JSON.parse((await native('TaskStatus')).trim());
        if (
          !['Unknown', 'Disabled', 'Queued', 'Ready', 'Running', 'Missing', 'Unavailable'].includes(
            state,
          )
        )
          return undefined;
        return { state, result: /^0x[0-9A-F]{8}$/.test(result) ? result : null };
      },
    }),
    async preflight(action = 'install') {
      if (platform === 'win32')
        return native('Preflight', { RequireClosedOffice: action === 'restart' ? '0' : '1' });
      await checkPlist();
      for (const app of ['Word', 'Excel', 'Powerpoint'])
        await preflightLocalManifest(
          join(
            homedir(),
            `Library/Containers/com.microsoft.${app}/Data/Documents/wef/triobridge-ai.xml`,
          ),
        );
    },
    async generateCertificate(directory, id) {
      const created = await mkdir(directory, { recursive: true, mode: 0o700 });
      await enforcePrivatePath(directory, 'directory', { created: created !== undefined });
      if (platform === 'win32')
        return native('Certificate', { CertificateDirectory: directory, InstallationId: id });
      const config = join(directory, 'openssl.cnf');
      await writeFile(
        config,
        '[req]\ndistinguished_name=dn\nx509_extensions=extensions\nprompt=no\n[dn]\nCN=localhost\nO=TrioBridge AI\nOU=' +
          id +
          '\n[extensions]\nsubjectAltName=DNS:localhost,IP:127.0.0.1,IP:::1\nbasicConstraints=critical,CA:FALSE\nkeyUsage=critical,digitalSignature,keyEncipherment\nextendedKeyUsage=serverAuth\n',
        { mode: 0o600 },
      );
      await execute(
        '/usr/bin/openssl',
        [
          'req',
          '-x509',
          '-newkey',
          'rsa:2048',
          '-nodes',
          '-sha256',
          '-days',
          '365',
          '-keyout',
          join(directory, 'server.key'),
          '-out',
          join(directory, 'server.crt'),
          '-config',
          config,
        ],
        { timeout: 60_000 },
      );
      await enforcePrivatePath(join(directory, 'server.key'), 'file');
      await enforcePrivatePath(join(directory, 'server.crt'), 'file');
      const certificate = new X509Certificate(await readFile(join(directory, 'server.crt')));
      await writeFile(
        join(directory, 'identity.json'),
        JSON.stringify({
          installationId: id,
          thumbprint: certificate.fingerprint.replaceAll(':', '').toUpperCase(),
        }) + '\n',
        { mode: 0o600 },
      );
      await rm(config);
    },
    async trustCertificate(directory, id) {
      await certificateIdentity(directory, id);
      if (platform === 'win32')
        return native('Trust', { CertificateDirectory: directory, InstallationId: id });
      const keychain = join(homedir(), 'Library/Keychains/login.keychain-db');
      try {
        await execute(
          '/usr/bin/security',
          ['verify-cert', '-c', join(directory, 'server.crt'), '-p', 'ssl', '-s', 'localhost'],
          { timeout: 15_000 },
        );
        return;
      } catch {
        /* Ask macOS to trust this installation's localhost certificate. */
      }
      await execute(
        '/usr/bin/security',
        [
          'add-trusted-cert',
          '-r',
          'trustRoot',
          '-p',
          'ssl',
          '-s',
          'localhost',
          '-k',
          keychain,
          join(directory, 'server.crt'),
        ],
        { timeout: 180_000 },
      );
      await execute(
        '/usr/bin/security',
        ['verify-cert', '-c', join(directory, 'server.crt'), '-p', 'ssl', '-s', 'localhost'],
        { timeout: 15_000 },
      );
    },
    async removeCertificate(directory, id) {
      const identity = await certificateIdentity(directory, id);
      if (platform === 'win32')
        return native('RemoveCertificate', { CertificateDirectory: directory, InstallationId: id });
      const keychain = join(homedir(), 'Library/Keychains/login.keychain-db');
      try {
        await execute('/usr/bin/security', [
          'delete-certificate',
          '-t',
          '-Z',
          identity.thumbprint,
          keychain,
        ]);
      } catch (error) {
        if (!String(error.stderr).includes('could not be found')) throw error;
      }
    },
    async stop() {
      if (platform === 'win32') return native('Stop');
      await checkPlist();
      try {
        await execute('/bin/launchctl', ['print', service]);
      } catch (error) {
        if ([3, 113].includes(error.code)) return;
        throw error;
      }
      await execute('/bin/launchctl', ['bootout', service]);
    },
    async start(state) {
      if (platform === 'win32') return native('Start');
      await checkPlist();
      await mkdir(dirname(plist), { recursive: true });
      await assertRegularOwned(dirname(plist), 'directory');
      await mkdir(join(root, 'logs'), { recursive: true, mode: 0o700 });
      const version = join(root, 'versions', state.active);
      await writeFile(
        plist,
        `<?xml version="1.0" encoding="UTF-8"?>\n<!DOCTYPE plist PUBLIC "-//Apple//DTD PLIST 1.0//EN" "http://www.apple.com/DTDs/PropertyList-1.0.dtd">\n<plist version="1.0"><dict><key>Label</key><string>${label}</string><key>ProgramArguments</key><array><string>${xml(join(version, 'runtime/node'))}</string><string>${xml(join(version, 'app/scripts/serve-installed.mjs'))}</string></array><key>WorkingDirectory</key><string>${xml(join(version, 'app'))}</string><key>RunAtLoad</key><true/><key>KeepAlive</key><true/><key>ThrottleInterval</key><integer>10</integer><key>StandardOutPath</key><string>${xml(join(root, 'logs/service.log'))}</string><key>StandardErrorPath</key><string>${xml(join(root, 'logs/service-error.log'))}</string></dict></plist>\n`,
        { mode: 0o600 },
      );
      await execute('/bin/launchctl', ['enable', service]);
      await execute('/bin/launchctl', ['bootstrap', `gui/${process.getuid()}`, plist]);
    },
    async installManifest(manifest) {
      if (platform === 'win32') return native('Catalog', { ManifestPath: manifest });
      const text = await readFile(manifest, 'utf8');
      for (const app of ['Word', 'Excel', 'Powerpoint'])
        await writePrivateLocalManifest(
          join(
            homedir(),
            `Library/Containers/com.microsoft.${app}/Data/Documents/wef/triobridge-ai.xml`,
          ),
          text,
        );
    },
    async removeManifest() {
      if (platform === 'win32') return native('RemoveCatalog');
      for (const app of ['Word', 'Excel', 'Powerpoint']) {
        const file = join(
          homedir(),
          `Library/Containers/com.microsoft.${app}/Data/Documents/wef/triobridge-ai.xml`,
        );
        await preflightLocalManifest(file);
        try {
          if ((await readFile(file, 'utf8')).includes(`<Id>${LOCAL_MANIFEST_ID}</Id>`))
            await rm(file);
        } catch (error) {
          if (error.code !== 'ENOENT') throw error;
        }
      }
    },
    async removeService() {
      if (platform === 'win32') return native('RemoveService');
      await checkPlist();
      await rm(plist, { force: true });
    },
    async registerUninstaller() {
      if (platform === 'win32') return native('RegisterUninstaller');
      const launcher =
        '#!/bin/bash\ninstaller_root=$(cd "$(dirname "$0")" && pwd)\nversion=$("/usr/bin/plutil" -extract active raw -o - "$installer_root/installation.json")\n"$installer_root/versions/$version/Uninstall.command"\n';
      await writeFile(join(root, 'Uninstall.command'), launcher, { mode: 0o700 });
    },
  };
}
