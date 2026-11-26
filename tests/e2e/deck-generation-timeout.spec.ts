import { expect, test, type Route } from '@playwright/test';

test('a deck built from uploaded sources can finish after the model timeout and regeneration can still be stopped', async ({
  page,
}) => {
  await page.addInitScript(() => {
    const profile = {
      id: 'delayed',
      name: 'Delayed test model',
      model: 'fixture',
      apiFormat: 'anthropic',
      baseUrl: location.origin + '/delayed-model/v1',
      timeoutSeconds: 60,
    };
    localStorage.setItem(
      'triobridge.profiles.v1',
      JSON.stringify({
        version: 1,
        profiles: [profile],
        activeId: profile.id,
      }),
    );
  });
  await page.goto('/?demo-host=powerpoint');
  await page.getByLabel('选择素材文件', { exact: true }).setInputFiles({
    name: 'workshop.md',
    mimeType: 'text/markdown',
    buffer: Buffer.from('Workshop attendance was 24.'),
  });
  await expect(page.getByRole('button', { name: '素材上下文 · 1 份已选' })).toBeVisible();
  await page.getByLabel('演示要求', { exact: true }).fill('Summarize the workshop.');
  await page.getByLabel('生成页数', { exact: true }).fill('2');
  await page.clock.install();
  const endpoint = '**/delayed-model/v1/messages';
  async function generate(button: string) {
    let hold!: (route: Route) => void;
    const pending = new Promise<Route>((resolve) => {
      hold = resolve;
    });
    await page.route(endpoint, hold, { times: 1 });
    await page.getByRole('button', { name: button, exact: true }).click();
    return pending;
  }
  const request = await generate('生成整套 PPT');
  expect(request.request().postData()).toContain('Workshop attendance was 24.');
  await page.clock.fastForward(65_000);
  const deck = {
    title: 'Workshop review',
    slides: [
      { title: 'Workshop attendance', layout: 'cover', subtitle: '24 participants' },
      { title: 'Next steps', layout: 'closing', subtitle: 'Collect feedback' },
    ],
  };
  await request.fulfill({
    json: {
      stop_reason: 'end_turn',
      content: [{ type: 'text', text: JSON.stringify(deck) }],
    },
  });
  const preview = page.getByRole('region', { name: '整稿预览' });
  await expect(preview.getByLabel('文件标题')).toHaveValue(deck.title);
  await expect(preview.getByRole('paragraph').filter({ hasText: '24 participants' })).toBeVisible();

  const retry = await generate('重新生成整套 PPT');
  await page.getByRole('button', { name: '停止', exact: true }).click();
  await expect(page.getByText('已停止，之前的资料与预览仍保留。')).toBeVisible();
  await expect(preview.getByLabel('文件标题')).toHaveValue(deck.title);
  await retry.abort();
});
