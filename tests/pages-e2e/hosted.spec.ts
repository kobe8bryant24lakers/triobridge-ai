import { test, expect } from '@playwright/test';

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
    await expect(page.getByRole('button', { name: '生成预览', exact: true })).toBeDisabled();
    await expect(page.locator('.brand img')).toHaveAttribute('src', '/triobridge-ai/app/icon.svg');
  });
}

test('site resources and hosted manifest are served while unapproved docs stay private', async ({
  request,
}) => {
  for (const path of [
    '',
    'site.css',
    'assets/icon.svg',
    'assets/workbench.png',
    'app/icon-32.png',
    'app/icon-80.png',
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
