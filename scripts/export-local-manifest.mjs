import { readFile } from 'node:fs/promises';
import { resolve, join } from 'node:path';
import { getOrCreateLocalCapability } from './local-capability.mjs';
import { sharedModelsDirectory } from './shared-model-store.mjs';
import { renderLocalManifest, writePrivateLocalManifest } from './local-manifest.mjs';

const directory = sharedModelsDirectory();
const manifest = renderLocalManifest(
  await readFile(resolve('manifest.xml'), 'utf8'),
  await getOrCreateLocalCapability(directory),
);
const destination = join(directory, 'manifest', 'triobridge-ai.xml');
await writePrivateLocalManifest(destination, manifest);
console.log(`本机专用清单已生成：${destination}`);
console.log('仅供当前系统账号的 Word / Excel / PowerPoint 安装，请勿上传或公开分享此清单。');
console.log('保持本机 HTTPS 服务运行；在一个应用配置或迁移模型后，其余应用会同步。');
