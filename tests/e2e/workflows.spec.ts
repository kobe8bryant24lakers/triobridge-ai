import { test, expect } from '@playwright/test';

test('Word: preview, cancel write, then confirm; manual edits invalidate source binding', async ({
  page,
}) => {
  await page.goto('/');
  await page.getByRole('button', { name: '读取示例选区' }).click();
  await page.getByRole('button', { name: '生成演示预览' }).click();
  await expect(page.getByText('固定演示样例 · 未调用 AI')).toBeVisible();
  await expect(page.getByText('查看模拟文档 · 已应用的内容')).not.toBeVisible();
  await page.getByRole('button', { name: '替换已读取选区' }).click();
  await expect(page.getByRole('dialog')).toBeVisible();
  await page.getByRole('button', { name: '返回检查' }).click();
  await expect(page.getByText('查看模拟文档 · 已应用的内容')).not.toBeVisible();
  await page.getByRole('button', { name: '替换已读取选区' }).click();
  await page.getByRole('button', { name: '确认写入', exact: true }).click();
  await expect(page.getByText('已写入', { exact: true })).toBeVisible();
  await page.getByText('查看模拟文档 · 已应用的内容').click();
  await expect(page.locator('.demo-document pre')).toContainText('本季度，团队');
  await page.getByLabel('参考内容').fill('手动粘贴的文本');
  await page.getByRole('button', { name: '生成演示预览' }).click();
  await expect(page.getByText('固定演示样例 · 未调用 AI')).toBeVisible();
  await expect(page.getByRole('button', { name: '替换已读取选区' })).not.toBeVisible();
});

test('Excel: generate formula and write to the reviewed cell', async ({ page }) => {
  await page.goto('/');
  await page.getByRole('button', { name: 'Excel', exact: true }).click();
  await page.getByLabel('具体要求').fill('计算 B2:B4 合计');
  await page.getByRole('button', { name: '生成演示预览' }).click();
  await expect(page.locator('.formula-text')).toHaveText('=SUM(B2:B4)');
  await page.getByRole('button', { name: '写入选中单元格' }).click();
  await expect(page.getByRole('dialog')).toContainText('B6');
  await page.getByRole('button', { name: '确认写入', exact: true }).click();
  await expect(page.getByText('已写入', { exact: true })).toBeVisible();
});

test('PowerPoint: structured outline, edit validation, and slide creation', async ({ page }) => {
  await page.goto('/');
  await page.getByRole('button', { name: 'PowerPoint', exact: true }).click();
  await page.getByLabel('具体要求').fill('季度工作回顾');
  await page.getByRole('button', { name: '生成演示预览' }).click();
  await expect(page.locator('.slide-outline > li')).toHaveCount(4);
  await page.getByRole('button', { name: '编辑', exact: true }).click();
  const original = await page.getByLabel('编辑结果').inputValue();
  await page.getByLabel('编辑结果').fill('{"slides":[]}');
  await expect(page.getByRole('button', { name: '创建幻灯片' })).toBeDisabled();
  await page.getByLabel('编辑结果').fill(original);
  await page.getByRole('button', { name: '完成编辑' }).click();
  await page.getByRole('button', { name: '创建幻灯片' }).click();
  await page.getByRole('button', { name: '确认写入', exact: true }).click();
  await expect(page.getByText('已写入', { exact: true })).toBeVisible();
});

test('cancellation and model HTTP errors remain recoverable', async ({ page }) => {
  await page.goto('/');
  await page.getByRole('button', { name: '读取示例选区' }).click();
  await page.getByRole('button', { name: '生成演示预览' }).click();
  await page.getByRole('button', { name: '停止', exact: true }).click();
  await expect(page.getByText('已取消生成。', { exact: true })).toBeVisible();
  await page.getByRole('button', { name: '模型设置', exact: true }).click();
  await page.getByLabel('模型服务').selectOption('custom');
  await page.getByLabel('服务名称').fill('Test model');
  await page.getByLabel('API Base URL').fill('https://example.test/v1');
  await page.getByLabel('模型 ID').fill('example-model');
  await page.getByLabel(/API Key 本地/).fill('test-only-key');
  await page.getByRole('button', { name: '仅保存配置' }).click();
  await page.route('https://example.test/v1/chat/completions', (route) =>
    route.fulfill({
      status: 401,
      contentType: 'application/json',
      body: '{"error":"invalid"}',
    }),
  );
  await page.getByRole('button', { name: '连接并保存' }).click();
  await expect(page.getByRole('alert')).toContainText('API Key 无效');
  await page.getByRole('button', { name: '返回工作台' }).click();
  await page.getByRole('button', { name: '读取示例选区' }).click();
  await page.getByRole('button', { name: '生成预览', exact: true }).click();
  await expect(page.getByRole('alert')).toContainText('API Key 无效');
  await expect(page.getByRole('button', { name: '生成预览', exact: true })).toBeEnabled();
});

test('360px Office pane has no horizontal overflow in workbench or settings', async ({ page }) => {
  await page.setViewportSize({ width: 360, height: 800 });
  await page.goto('/');
  expect(await page.evaluate(() => document.documentElement.scrollWidth <= innerWidth)).toBe(true);
  await page.getByRole('button', { name: '模型设置', exact: true }).click();
  expect(await page.evaluate(() => document.documentElement.scrollWidth <= innerWidth)).toBe(true);
  await expect(page.getByRole('button', { name: '连接并保存' })).toBeAttached();
});

test('keyboard users can inspect both overflowing confirmation previews', async ({ page }) => {
  await page.setViewportSize({ width: 360, height: 800 });
  await page.goto('/');
  await page.getByRole('button', { name: '读取示例选区' }).click();
  await page.getByRole('button', { name: '生成演示预览' }).click();
  await page.getByRole('button', { name: '编辑', exact: true }).click();
  await page
    .getByLabel('编辑结果')
    .fill(
      Array.from({ length: 30 }, (_, i) => `第 ${i + 1} 行：用于检查键盘滚动的长段落。`).join('\n'),
    );
  await page.getByRole('button', { name: '完成编辑' }).click();
  await page.getByRole('button', { name: '替换已读取选区' }).click();
  await page.getByRole('button', { name: '确认写入', exact: true }).click();
  await page.getByRole('button', { name: '读取示例选区' }).click();
  await page.getByRole('button', { name: '生成演示预览' }).click();
  await page.getByRole('button', { name: '替换已读取选区' }).click();
  await expect(page.getByRole('button', { name: '返回检查' })).toBeFocused();
  await page.keyboard.press('Tab');
  await expect(page.getByRole('button', { name: '确认写入', exact: true })).toBeFocused();
  await page.keyboard.press('Tab');
  const before = page.getByRole('region', { name: '目标当前内容' });
  const after = page.getByRole('region', { name: '即将写入的结果' });
  for (const region of [before, after]) {
    await expect(region).toBeFocused();
    await page.keyboard.press('End');
    await expect
      .poll(() =>
        region.evaluate(
          (el) => el.scrollTop > 0 && el.scrollTop + el.clientHeight >= el.scrollHeight - 1,
        ),
      )
      .toBe(true);
    await expect
      .poll(() => region.evaluate((el) => getComputedStyle(el).outlineStyle))
      .toBe('solid');
    await page.keyboard.press('Tab');
  }
  await page.keyboard.press('Escape');
  await expect(page.getByRole('dialog')).not.toBeVisible();
});
