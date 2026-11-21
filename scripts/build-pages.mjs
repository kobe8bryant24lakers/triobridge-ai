import { copyFile, lstat, mkdir, readFile, rm, writeFile } from 'node:fs/promises';
import { dirname, join, resolve } from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';
import { buildInstallers, INSTALLER_FILES } from './build-installers.mjs';

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

export async function verifyPagesOutput(destination) {
  for (const file of INSTALLER_FILES) await assertRegularFile(join(destination, 'install', file));
  for (const [, output] of SITE_FILES) await assertRegularFile(join(destination, output));
  for (const icon of APP_ICONS) await assertRegularFile(join(destination, 'app', icon));
  const html = await readFile(join(destination, 'app/index.html'), 'utf8');
  const references = [...html.matchAll(/(?:src|href)="([^"]+)"/g)].map((match) => match[1]);
  if (!references.length) throw new Error('Application HTML has no asset references.');
  for (const reference of references) {
    if (!reference.startsWith(APP_BASE) || reference.includes('..'))
      throw new Error(`Application asset escapes the hosted base path: ${reference}`);
    await assertRegularFile(join(destination, 'app', reference.slice(APP_BASE.length)));
  }
  const manifest = await readFile(join(destination, 'manifest.xml'), 'utf8');
  if (!manifest.includes(`DefaultValue="${APP_URL}?office=1"`))
    throw new Error('Hosted manifest does not preserve the Office entry query.');
}

export async function buildPages(root = PROJECT_ROOT) {
  const destination = join(root, 'dist-pages');
  // Verify the allowlist before replacing the previous build. Never copy docs/ or public/ recursively.
  for (const [source] of SITE_FILES) await assertRegularFile(join(root, source));
  for (const icon of APP_ICONS) await assertRegularFile(join(root, 'public', icon));
  const manifest = hostedManifest(await readFile(join(root, 'manifest.xml'), 'utf8'));
  await rm(destination, { recursive: true, force: true });
  await copySiteFiles(root, destination);
  const { build } = await import('vite');
  await build({
    root,
    configFile: join(root, 'vite.config.ts'),
    base: APP_BASE,
    envDir: false,
    publicDir: false,
    build: { outDir: join(destination, 'app'), emptyOutDir: true },
  });
  for (const icon of APP_ICONS)
    await copyFile(join(root, 'public', icon), join(destination, 'app', icon));
  await writeFile(join(destination, 'manifest.xml'), manifest);
  await buildInstallers(root, join(destination, 'install'), manifest, HOSTED_MANIFEST_ID);
  await writeFile(join(destination, '.nojekyll'), '');
  await verifyPagesOutput(destination);
  console.log(`Pages / Release files: ${destination}`);
  console.log(`Office manifest: ${SITE_URL}manifest.xml (separate hosted ID)`);
}

if (process.argv[1] && import.meta.url === pathToFileURL(resolve(process.argv[1])).href)
  await buildPages();
