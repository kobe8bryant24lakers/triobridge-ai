import { copyFile, lstat, mkdir, readFile, rm } from 'node:fs/promises';
import { dirname, join, resolve } from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';

export const SITE_URL = 'https://kobe8bryant24lakers.github.io/triobridge-ai/';
export const APP_BASE = '/triobridge-ai/app/';
export const APP_URL = new URL('app/', SITE_URL).href;
export const HOSTED_MANIFEST_ID = '9d977989-12ea-46cf-920e-b54a12b4d728';
export const SITE_FILES = [
  ['docs/index.html', 'index.html'],
  ['docs/zh-CN.html', 'zh-CN.html'],
  ['docs/site.css', 'site.css'],
  ['docs/assets/icon.svg', 'assets/icon.svg'],
  ['docs/assets/workbench.png', 'assets/workbench.png'],
  ['docs/assets/TrioBridge-AI-User-Guide-en.pptx', 'assets/TrioBridge-AI-User-Guide-en.pptx'],
  ['docs/assets/TrioBridge-AI-User-Guide-zh-CN.pptx', 'assets/TrioBridge-AI-User-Guide-zh-CN.pptx'],
];
export const APP_ICONS = ['icon.svg', 'icon-32.png', 'icon-80.png'];
const PROJECT_ROOT = resolve(dirname(fileURLToPath(import.meta.url)), '..');

function replaceOne(xml, pattern, replacement) {
  if ([...xml.matchAll(pattern)].length !== 1)
    throw new Error(`Expected exactly one manifest field: ${pattern}`);
  return xml.replace(pattern, replacement);
}

export function hostedManifest(localManifest) {
  if (localManifest.includes(HOSTED_MANIFEST_ID))
    throw new Error('The hosted manifest ID must differ from the local manifest ID.');
  let xml = replaceOne(localManifest, /<Id>[^<]+<\/Id>/g, `<Id>${HOSTED_MANIFEST_ID}</Id>`);
  xml = replaceOne(
    xml,
    /<DisplayName\b[^>]*\/>/g,
    '<DisplayName DefaultValue="TrioBridge AI · 三桥 AI（在线版）" />',
  );
  xml = replaceOne(xml, /<IconUrl\b[^>]*\/>/g, `<IconUrl DefaultValue="${APP_URL}icon-32.png" />`);
  xml = replaceOne(
    xml,
    /<HighResolutionIconUrl\b[^>]*\/>/g,
    `<HighResolutionIconUrl DefaultValue="${APP_URL}icon-80.png" />`,
  );
  xml = replaceOne(
    xml,
    /<AppDomains>[\s\S]*?<\/AppDomains>/g,
    `<AppDomains><AppDomain>${new URL(SITE_URL).origin}</AppDomain></AppDomains>`,
  );
  xml = replaceOne(
    xml,
    /<SourceLocation\b[^>]*\/>/g,
    `<SourceLocation DefaultValue="${APP_URL}?office=1" />`,
  );
  if (/localhost|127\.0\.0\.1/.test(xml)) throw new Error('Hosted manifest contains a local URL.');
  return xml;
}

async function assertRegularFile(path) {
  const info = await lstat(path);
  if (!info.isFile() || info.isSymbolicLink())
    throw new Error(`Only regular, explicitly approved files can be published: ${path}`);
}

export async function copySiteFiles(root, destination) {
  for (const [source] of SITE_FILES) await assertRegularFile(join(root, source));
  for (const [source, output] of SITE_FILES) {
    await mkdir(dirname(join(destination, output)), { recursive: true });
    await copyFile(join(root, source), join(destination, output));
  }
}

export async function buildPages(root = PROJECT_ROOT) {
  const destination = join(root, 'dist-pages');
  for (const [source] of SITE_FILES) await assertRegularFile(join(root, source));
  await rm(destination, { recursive: true, force: true });
  await copySiteFiles(root, destination);
  for (const file of ['index.html', 'zh-CN.html']) {
    const page = await readFile(join(destination, file), 'utf8');
    if (
      !page.includes('TrioBridge-AI-Windows-x64.zip') ||
      !page.includes('TrioBridge-AI-Mac-arm64.zip')
    )
      throw new Error('Missing desktop installer links.');
  }
}
if (process.argv[1] && import.meta.url === pathToFileURL(resolve(process.argv[1])).href)
  await buildPages();
