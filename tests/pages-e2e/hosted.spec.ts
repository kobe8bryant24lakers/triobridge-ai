import { expect, test } from '@playwright/test';

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
    await expect(page.locator('a[href*="/app/"], a[href$=".xml"], a[href*="-Local-"]')).toHaveCount(
      0,
    );
    await expect(page.locator('#guide')).toContainText('127.0.0.1:8001/v1');
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
test('publishes only the guide, with no hosted application, manifest, installer or private docs', async ({
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
  test('both languages link to downloads', async ({ page }) => {
    await page.goto('');
    await page.locator('.language-switch').getByRole('link', { name: '中文' }).click();
    await page.locator('.hero').getByRole('link', { name: '下载并安装' }).click();
    await expect(page.locator('#install a').first()).toBeVisible();
  });
});
