import { readFile } from 'node:fs/promises';
import { expect, test } from '@playwright/test';
import { strFromU8, unzipSync } from 'fflate';

for (const [path, lang] of [
  ['', 'en'],
  ['zh-CN.html', 'zh-CN'],
]) {
  test(`${lang} publishes four desktop packages and the two installation paths`, async ({
    page,
  }) => {
    await page.goto(path);
    await expect(page.locator('html')).toHaveAttribute('lang', lang);
    const install = page.locator('#install');
    for (const name of ['Windows-x64', 'Windows-arm64', 'Mac-x64', 'Mac-arm64']) {
      await expect(install.locator('a[href$="TrioBridge-AI-' + name + '.zip"]')).toHaveCount(1);
    }
    await expect(install).toContainText('Install.cmd');
    await expect(install).toContainText('Install.command');
    for (const platform of ['Windows', 'Mac']) {
      const card = install.getByRole('article', { name: platform, exact: true });
      const instructions = card.locator('details p');
      await expect(instructions).toBeHidden();
      await card.locator('summary').press('Enter');
      await expect(instructions).toBeVisible();
    }
    await expect(page.locator('a[href*="/app/"], a[href$=".xml"], a[href*="-Local-"]')).toHaveCount(
      0,
    );
    await expect(page.locator('#guide')).toContainText('127.0.0.1:8001/v1');
  });
  test(`${lang} downloads the guide in the page language`, async ({ page }) => {
    await page.goto(path);
    const filename = `TrioBridge-AI-User-Guide-${lang}.pptx`;
    const link = page.locator('#guide a[download]');
    await expect(link).toHaveAttribute('href', `./assets/${filename}`);
    const pendingDownload = page.waitForEvent('download');
    await link.click();
    const download = await pendingDownload;
    expect(download.suggestedFilename()).toBe(filename);
    const file = await download.path();
    if (!file) throw new Error('The guide download did not complete.');
    const cover = 'ppt/slides/slide1.xml';
    const archive = unzipSync(await readFile(file), { filter: (entry) => entry.name === cover });
    expect(strFromU8(archive[cover])).toMatch(lang === 'en' ? /user guide/i : /使用指南/);
  });
  test(`${lang} keeps downloads and language links usable at phone and desktop widths`, async ({
    page,
  }) => {
    await page.goto(path);
    for (const width of [360, 768, 1280]) {
      await page.setViewportSize({ width, height: 900 });
      await expect(page.locator('.primary-nav')).toBeVisible();
      for (const name of ['English', '中文'])
        await expect(
          page.locator('.language-switch').getByRole('link', { name, exact: true }),
        ).toBeInViewport();
      expect(await page.evaluate(() => document.documentElement.scrollWidth <= innerWidth)).toBe(
        true,
      );
    }
  });
}
test('publishes only the public site, with no hosted application, manifest, installer or private docs', async ({
  request,
}) => {
  for (const path of ['', 'zh-CN.html', 'site.css', 'assets/icon.svg', 'assets/workbench.png'])
    expect((await request.get(path)).status()).toBe(200);
  for (const path of [
    'app/',
    'manifest.xml',
    'install/TrioBridge-AI-Mac.zip',
    'LOCAL-MODELS.md',
    'TESTING.md',
    '.env',
  ])
    expect((await request.get(path)).status()).toBe(404);
});
test.describe('without JavaScript', () => {
  test.use({ javaScriptEnabled: false });
  for (const path of ['', 'zh-CN.html']) {
    test(`app previews respond to click and arrow keys on ${path || 'the English page'}`, async ({
      page,
    }) => {
      await page.emulateMedia({ reducedMotion: 'reduce' });
      await page.goto(path);
      const showcase = page.locator('.showcase');
      const previewLink = showcase.getByRole('link');
      await expect(showcase.getByRole('radio', { name: 'Word', exact: true })).toBeChecked();
      await expect(previewLink).toHaveAttribute('href', '#word');
      for (const image of await showcase.locator('img').all()) {
        await expect(image).toHaveCSS('transition-duration', '0s');
        expect(await image.evaluate((node: HTMLImageElement) => node.naturalWidth)).toBeGreaterThan(
          0,
        );
      }

      const excelModel = showcase.locator('.excel-model');
      const backWidth = await excelModel.evaluate((image) => image.getBoundingClientRect().width);
      await showcase.locator('label[for="show-excel"]').click();
      const excel = showcase.getByRole('radio', { name: 'Excel', exact: true });
      await expect(excel).toBeChecked();
      await expect(previewLink).toHaveAttribute('href', '#excel');
      expect(
        await excelModel.evaluate((image) => image.getBoundingClientRect().width),
      ).toBeGreaterThan(backWidth);
      await excel.press('ArrowRight');
      await expect(showcase.getByRole('radio', { name: 'PowerPoint', exact: true })).toBeChecked();
      await expect(previewLink).toHaveAttribute('href', '#powerpoint');
      await previewLink.click();
      await expect(page).toHaveURL(/#powerpoint$/);
      await expect(page.locator('#powerpoint h3')).toBeInViewport();
    });
  }

  test('both languages link to downloads', async ({ page }) => {
    await page.goto('');
    await page.locator('.language-switch').getByRole('link', { name: '中文' }).click();
    await page.locator('.hero').getByRole('link', { name: '下载并安装' }).click();
    await expect(page.locator('#install a').first()).toBeVisible();
  });
});
