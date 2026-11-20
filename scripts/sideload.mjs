import { copyFile, mkdir, unlink } from 'node:fs/promises';
import { join, resolve } from 'node:path';
import { homedir } from 'node:os';

if (process.platform !== 'darwin')
  throw new Error('此脚本仅适用于 macOS。Windows 请按 README 中的受信任目录步骤安装。');
const action = process.argv[2];
if (!['install', 'remove'].includes(action)) throw new Error('使用 install 或 remove。');
for (const app of ['Word', 'Excel', 'Powerpoint']) {
  const directory = join(
    homedir(),
    'Library',
    'Containers',
    `com.microsoft.${app}`,
    'Data',
    'Documents',
    'wef',
  );
  const destination = join(directory, 'triobridge-ai.xml');
  if (action === 'install') {
    await mkdir(directory, { recursive: true });
    await copyFile(resolve('manifest.xml'), destination);
  } else {
    try {
      await unlink(destination);
    } catch (error) {
      if (error.code !== 'ENOENT') throw error;
    }
  }
  console.log(`${action}: ${destination}`);
}
console.log('请重新打开 Office。此脚本只管理 TrioBridge 自己的清单文件。');
