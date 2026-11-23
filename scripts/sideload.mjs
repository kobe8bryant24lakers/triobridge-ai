import { readFile, unlink } from 'node:fs/promises';
import { join, resolve } from 'node:path';
import { homedir } from 'node:os';
import { getOrCreateLocalCapability } from './local-capability.mjs';
import { sharedModelsDirectory } from './shared-model-store.mjs';
import {
  preflightLocalManifest,
  renderLocalManifest,
  writePrivateLocalManifest,
} from './local-manifest.mjs';

if (process.platform !== 'darwin')
  throw new Error('此脚本仅适用于 macOS。Windows 请按 README 中的受信任目录步骤安装。');
const action = process.argv[2];
if (!['install', 'remove'].includes(action)) throw new Error('使用 install 或 remove。');
const manifest =
  action === 'install'
    ? renderLocalManifest(
        await readFile(resolve('manifest.xml'), 'utf8'),
        await getOrCreateLocalCapability(sharedModelsDirectory()),
      )
    : undefined;
const destinations = ['Word', 'Excel', 'Powerpoint'].map((app) =>
  join(
    homedir(),
    'Library',
    'Containers',
    `com.microsoft.${app}`,
    'Data',
    'Documents',
    'wef',
    'triobridge-ai.xml',
  ),
);
// Reject ownership conflicts in every host before installing or removing anything.
for (const destination of destinations) await preflightLocalManifest(destination);
for (const destination of destinations) {
  if (action === 'install') {
    await writePrivateLocalManifest(destination, manifest);
  } else {
    try {
      await unlink(destination);
    } catch (error) {
      if (error.code !== 'ENOENT') throw error;
    }
  }
  console.log(`${action}: ${destination}`);
}
console.log('请重新打开 Office。本机共享版需保持 npm start 或 npm run dev:office 运行。');
console.log('在任一应用迁移或配置模型一次，即可在同一电脑、同一系统账号下的三类 Office 中共用。');
console.log('卸载仅管理 TrioBridge 清单，不删除共享模型或系统安全存储中的密钥。');
