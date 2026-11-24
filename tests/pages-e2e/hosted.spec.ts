import { test, expect } from '@playwright/test';
import { syntheticPdf } from '../pdf-fixture';

test('hosted demo loads assets under the repository/app base and completes a preview', async ({
  page,
}) => {
  const failures: string[] = [];
  page.on('response', (response) => {
    if (response.status() >= 400) failures.push(`${response.status()} ${response.url()}`);
  });
  await page.goto('app/');
  await expect(page.getByRole('heading', { name: 'Word 工作台' })).toBeVisible();
  const brand = page.locator('.brand img');
  await expect(brand).toHaveAttribute('src', '/triobridge-ai/app/icon.svg');
  await expect
    .poll(() => brand.evaluate((image: HTMLImageElement) => image.naturalWidth))
    .toBeGreaterThan(0);
  await page.getByRole('button', { name: '读取示例选区' }).click();
  await page.getByRole('button', { name: '生成演示预览' }).click();
  await expect(page.getByText('固定演示样例 · 未调用 AI')).toBeVisible();
  expect(failures).toEqual([]);
});

for (const host of ['Word', 'Excel', 'PowerPoint']) {
  test(`hosted Office query starts the ${host} adapter without exposing the demo model`, async ({
    page,
  }) => {
    await page.route('https://appsforoffice.microsoft.com/lib/1/hosted/office.js', (route) =>
      route.fulfill({
        contentType: 'application/javascript',
        body: `window.Office = {onReady: () => Promise.resolve({host: ${JSON.stringify(host)}, platform: 'OfficeOnline'}), context: {requirements: {isSetSupported: () => false}, document: {}}};`,
      }),
    );
    await page.goto('app/?office=1');
    await expect(page.getByRole('heading', { name: `${host} 工作台` })).toBeVisible();
    await expect(page.getByText('Office 加载项', { exact: true })).toBeVisible();
    await expect(page.getByRole('option', { name: '演示模型 · 无需 API Key' })).toHaveCount(0);
    await expect(
      page.getByRole('button', {
        name: host === 'PowerPoint' ? '生成整套 PPT' : '生成预览',
        exact: true,
      }),
    ).toBeDisabled();
    await expect(page.locator('.brand img')).toHaveAttribute('src', '/triobridge-ai/app/icon.svg');
  });
}

test('production PDF worker loads from the hosted app base and extracts selected material locally', async ({
  page,
}) => {
  const posts: string[] = [],
    failures: string[] = [];
  page.on('request', (request) => {
    if (request.method() === 'POST') posts.push(request.url());
  });
  page.on('response', (response) => {
    if (response.status() >= 400) failures.push(response.url());
  });
  await page.goto('app/?demo-host=powerpoint');
  const worker = page.waitForResponse((response) => /pdf\.worker.*\.mjs$/.test(response.url()));
  await page.getByLabel('选择素材文件', { exact: true }).setInputFiles({
    name: 'synthetic-survey.pdf',
    mimeType: 'application/pdf',
    buffer: syntheticPdf(),
  });
  expect((await worker).headers()['content-type']).toContain('javascript');
  await expect(page.getByRole('button', { name: '素材上下文 · 1 份已选' })).toBeVisible();
  await page.getByText('查看内容', { exact: true }).click();
  await expect(page.locator('.material-text')).toContainText('120 responses');
  expect(posts).toEqual([]);
  expect(failures).toEqual([]);
});

test('site resources and hosted manifest are served while unapproved docs stay private', async ({
  request,
}) => {
  for (const path of [
    '',
    'zh-CN.html',
    'site.css',
    'assets/icon.svg',
    'assets/workbench.png',
    'app/icon-32.png',
    'app/icon-80.png',
    'install/TrioBridge-AI-Mac.zip',
    'install/TrioBridge-AI-Windows.zip',
  ])
    expect((await request.get(path)).status(), path).toBe(200);
  const manifest = await request.get('manifest.xml');
  expect(manifest.status()).toBe(200);
  expect(await manifest.text()).toContain(
    'https://kobe8bryant24lakers.github.io/triobridge-ai/app/?office=1',
  );
  expect((await request.get('TESTING.md')).status()).toBe(404);
  expect((await request.get('docs/TESTING.md')).status()).toBe(404);
  expect((await request.get('app/ASSETS.md')).status()).toBe(404);
});

test('quick setup works in the hosted pane and installer links download real ZIPs', async ({
  page,
}) => {
  await page.goto('');
  await page.locator('.hosted-edition > summary').click();
  for (const platform of ['Mac', 'Windows']) {
    const download = page.waitForEvent('download');
    await page
      .getByRole('link', { name: `Download ${platform} hosted installer`, exact: true })
      .click();
    expect((await download).suggestedFilename()).toBe(`TrioBridge-AI-${platform}.zip`);
  }
  await page.setViewportSize({ width: 360, height: 800 });
  await page.goto('app/');
  await page.getByRole('button', { name: '模型设置', exact: true }).click();
  await page.getByLabel('模型服务').selectOption('deepseek');
  await expect(page.getByLabel('API Base URL')).not.toBeAttached();
  await page.getByLabel('API Key', { exact: true }).fill('synthetic-key');
  await page.route('https://api.deepseek.com/chat/completions', (route) =>
    route.fulfill({ json: { choices: [{ message: { content: 'OK' } }] } }),
  );
  await page.getByRole('button', { name: '连接并保存', exact: true }).click();
  await expect(page.getByRole('status')).toContainText('连接成功，配置已保存');
  expect(await page.evaluate(() => document.documentElement.scrollWidth <= innerWidth)).toBe(true);
});

test('site defaults to English and supports Chinese navigation, reloads, and direct links', async ({
  page,
}) => {
  await page.goto('');
  await expect(page.locator('html')).toHaveAttribute('lang', 'en');
  const languages = page.locator('.language-switch');
  await expect(languages.getByRole('link', { name: 'English', exact: true })).toHaveAttribute(
    'aria-current',
    'page',
  );
  await expect(page.locator('head link[rel="alternate"][hreflang="zh-CN"]')).toHaveAttribute(
    'href',
    './zh-CN.html',
  );
  const englishTitle = await page.title();
  const englishHero = await page.locator('#hero-title').innerText();

  await languages.getByRole('link', { name: '中文', exact: true }).click();
  await expect(page).toHaveURL(/\/zh-CN\.html$/);
  await expect(page.locator('html')).toHaveAttribute('lang', 'zh-CN');
  await expect(languages.getByRole('link', { name: '中文', exact: true })).toHaveAttribute(
    'aria-current',
    'page',
  );
  await expect(page.locator('#hero-title')).toContainText('接上你的 AI 模型');
  expect(await page.title()).not.toBe(englishTitle);

  await page.reload();
  await expect(page.locator('html')).toHaveAttribute('lang', 'zh-CN');
  await page.goto('zh-CN.html#install');
  await expect(page.locator('#install')).toBeInViewport();
  await expect(
    page.getByRole('link', { name: '下载 Mac Apple 芯片版', exact: true }),
  ).toBeVisible();

  await languages.getByRole('link', { name: 'English', exact: true }).click();
  await expect(page.locator('html')).toHaveAttribute('lang', 'en');
  await expect(page).toHaveTitle(englishTitle);
  await expect(page.locator('#hero-title')).toHaveText(englishHero, { useInnerText: true });
});

test('both languages offer four distinct local packages before the hosted edition', async ({
  page,
}) => {
  for (const path of ['', 'zh-CN.html']) {
    await page.goto(path + '#install');
    const links = page.locator('.local-downloads a');
    await expect(links).toHaveCount(4);
    for (const [index, asset] of [
      'Windows-x64',
      'Windows-arm64',
      'Mac-arm64',
      'Mac-x64',
    ].entries()) {
      await expect(links.nth(index)).toBeVisible();
      await expect(links.nth(index)).toHaveAttribute(
        'href',
        `https://github.com/kobe8bryant24lakers/triobridge-ai/releases/download/v0.5.0/TrioBridge-AI-Local-${asset}.zip`,
      );
    }
    await expect(page.locator('.hosted-edition')).not.toHaveAttribute('open');
    await page.locator('.hosted-edition > summary').focus();
    await page.keyboard.press('Enter');
    await expect(page.locator('.hosted-downloads')).toBeVisible();
  }
});

test.describe('site without JavaScript', () => {
  test.use({ javaScriptEnabled: false, locale: 'zh-CN' });

  test('still defaults to English and switches languages using ordinary links', async ({
    page,
  }) => {
    await page.goto('');
    await expect(page.locator('html')).toHaveAttribute('lang', 'en');
    await page.locator('.language-switch').getByRole('link', { name: '中文', exact: true }).click();
    await expect(page.locator('html')).toHaveAttribute('lang', 'zh-CN');
    await expect(page.locator('#hero-title')).toContainText('接上你的 AI 模型');
    await page
      .locator('.language-switch')
      .getByRole('link', { name: 'English', exact: true })
      .click();
    await expect(page.locator('html')).toHaveAttribute('lang', 'en');
  });
});

test('both site languages keep navigation visible without horizontal overflow', async ({
  page,
}) => {
  for (const path of ['', 'zh-CN.html']) {
    await page.goto(path);
    for (const width of [360, 768, 1280]) {
      await page.setViewportSize({ width, height: 900 });
      await expect(page.locator('.primary-nav')).toBeVisible();
      for (const name of ['English', '中文']) {
        const language = page.locator('.language-switch').getByRole('link', { name, exact: true });
        await expect(language).toBeInViewport();
      }
      expect(
        await page.evaluate(() => document.documentElement.scrollWidth <= innerWidth),
        `${path || 'English home'} at ${width}px`,
      ).toBe(true);
    }
  }
});
