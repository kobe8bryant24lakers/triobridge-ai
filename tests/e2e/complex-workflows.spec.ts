import { test, expect } from '@playwright/test';

test('Word structured report: inspect, edit paragraphs and tables, review readable content and write', async ({
  page,
}) => {
  await page.goto('/');
  await page.getByLabel('示例文档').selectOption('complex');
  await page.getByRole('button', { name: '读取示例选区' }).click();
  await page.getByRole('button', { name: '结构化报告', exact: true }).click();
  await page.getByLabel('具体要求').fill('生成分节验收报告，保留所有指标和待办表格');
  await page.getByRole('button', { name: '生成演示预览' }).click();
  await expect(page.locator('.report-preview section')).toHaveCount(6);
  await expect(page.locator('.report-preview table')).toHaveCount(2);
  await page.getByRole('button', { name: '编辑', exact: true }).click();
  await page.getByLabel('文档标题', { exact: true }).fill('验收报告修订版');
  await page
    .getByLabel('第 2 节表格第 1 行第 4 列', { exact: true })
    .fill('按相同口径复核，保留原始记录');
  await page.getByRole('button', { name: '完成编辑' }).click();
  await page.getByRole('button', { name: '插入排版报告' }).click();
  await expect(page.getByRole('dialog')).toContainText('验收报告修订版');
  await expect(page.getByRole('dialog')).toContainText('按相同口径复核，保留原始记录');
  await expect(page.getByRole('dialog')).not.toContainText('"sections"');
  await page.getByRole('button', { name: '确认写入', exact: true }).click();
  await expect(page.getByText('已写入', { exact: true })).toBeVisible();
  await expect(page.getByRole('region', { name: '结果预览', exact: true })).toBeFocused();
});

test('Excel complex table: typed cell editing cannot silently lose an unapplied change', async ({
  page,
}) => {
  await page.goto('/');
  await page.getByRole('button', { name: 'Excel', exact: true }).click();
  await page.getByRole('button', { name: '创建数据表', exact: true }).click();
  await page.getByLabel('具体要求').fill('48 行运营数据，差额与差额率用公式');
  await page.getByRole('button', { name: '生成演示预览' }).click();
  await expect(page.locator('.result-summary')).toHaveText('48 行数据 · 8 列 · 96 个公式');
  await page.getByRole('button', { name: '编辑', exact: true }).click();
  await page.getByRole('button', { name: 'D2 1200', exact: true }).click();
  await page.getByRole('button', { name: 'D2 1200', exact: true }).focus();
  await page.keyboard.press('ArrowRight');
  await expect(page.getByRole('button', { name: 'E2 800', exact: true })).toBeFocused();
  await page.keyboard.press('ArrowLeft');
  await expect(page.getByRole('button', { name: 'D2 1200', exact: true })).toBeFocused();
  await page.getByLabel('单元格内容').fill('1500.25');
  await expect(page.getByRole('button', { name: '创建新工作表' })).toBeDisabled();
  await expect(page.getByRole('button', { name: '完成编辑' })).toBeDisabled();
  await page.getByRole('button', { name: '应用单元格修改' }).click();
  await page.getByRole('button', { name: '完成编辑' }).click();
  await expect(page.locator('.table-preview')).toContainText('1500.25');
  await page.getByRole('button', { name: '编辑', exact: true }).click();
  await page.getByLabel('新工作表名称').fill('不保留的修改');
  await page.getByLabel('单元格内容').fill('尚未应用的修改');
  await page.getByRole('button', { name: '取消本轮编辑' }).click();
  await expect(page.locator('.table-preview h3')).toHaveText('区域运营演示');
  await expect(page.locator('.table-preview')).toContainText('1500.25');
  await page.getByRole('button', { name: '创建新工作表' }).click();
  await expect(page.getByRole('dialog')).toContainText('区域运营演示');
  await expect(page.getByRole('dialog')).toContainText('=D49-E49');
  await page.getByRole('button', { name: '确认写入', exact: true }).click();
  await expect(page.getByText('已写入', { exact: true })).toBeVisible();
});

test('PowerPoint twelve-page outline can be reordered and repaired without JSON', async ({
  page,
}) => {
  await page.goto('/');
  await page.getByRole('button', { name: 'PowerPoint', exact: true }).click();
  await page.getByRole('button', { name: '文案与大纲', exact: true }).click();
  await page.getByLabel('示例文档').selectOption('complex');
  await page.getByLabel('具体要求').fill('12 页验收汇报');
  await page.getByRole('button', { name: '生成演示预览' }).click();
  await expect(page.locator('.slide-outline > li')).toHaveCount(12);
  await page.getByRole('button', { name: '编辑', exact: true }).click();
  await expect(page.getByRole('button', { name: '添加一页（12 / 12）' })).toBeDisabled();
  await page.getByRole('button', { name: '下移第 1 页', exact: true }).click();
  await expect(page.getByLabel('第 1 页标题', { exact: true })).toHaveValue('验收范围');
  await page.getByLabel('第 1 页标题', { exact: true }).fill('');
  await expect(page.getByRole('button', { name: '创建幻灯片' })).toBeDisabled();
  await expect(page.getByLabel('第 1 页标题', { exact: true })).toBeVisible();
  await page.getByLabel('第 1 页标题', { exact: true }).fill('验收范围修订');
  await page.getByRole('button', { name: '完成编辑' }).click();
  await page.getByRole('button', { name: '创建幻灯片' }).click();
  await expect(page.getByRole('region', { name: '即将写入的结果' })).toContainText(
    '12. 验收确认与行动清单',
  );
  await expect(page.getByRole('region', { name: '即将写入的结果' })).not.toContainText('"bullets"');
  await page.keyboard.press('Escape');
  await expect(page.getByRole('button', { name: '创建幻灯片' })).toBeFocused();
});

test('settings and help preserve drafts; changing requirements retains a visibly outdated preview', async ({
  page,
}) => {
  await page.goto('/');
  await page.getByRole('button', { name: '读取示例选区' }).click();
  const source = await page.getByLabel('参考内容').inputValue();
  await page.getByRole('button', { name: '生成演示预览' }).click();
  await expect(page.getByText('固定演示样例 · 未调用 AI')).toBeVisible();
  const result = await page.locator('.result-text').textContent();
  await page.getByRole('button', { name: '模型设置', exact: true }).click();
  await page.getByRole('button', { name: '返回工作台' }).click();
  await expect(page.getByLabel('参考内容')).toHaveValue(source);
  await expect(page.locator('.result-text')).toHaveText(result!);
  await page.getByRole('button', { name: '使用指南', exact: true }).click();
  await page.getByRole('button', { name: '工作台', exact: true }).click();
  await page.getByLabel('补充要求').fill('保留原文数字');
  await expect(page.getByText('旧预览', { exact: true })).toBeVisible();
  await expect(page.getByRole('button', { name: '替换已读取选区' })).toBeDisabled();
  await expect(page.locator('.result-text')).toHaveText(result!);
  await page.getByRole('button', { name: '生成演示预览' }).click();
  await page.getByRole('button', { name: '停止', exact: true }).click();
  await expect(page.getByText('已取消生成。', { exact: true })).toBeVisible();
  await expect(page.locator('.result-text')).toHaveText(result!);
});

test('oversized paste is retained and clearing the result never removes its editor', async ({
  page,
}) => {
  // Filling 4,000 lines through the browser's native insertText path can exceed
  // 30s on CI even in a plain textarea. Keep assertion deadlines unchanged.
  test.setTimeout(90_000);
  await page.goto('/');
  const large = '多语言文本🙂\n'.repeat(4000);
  await page.getByLabel('参考内容').fill(large);
  await expect(page.getByLabel('参考内容')).toHaveValue(large);
  await expect(page.getByRole('alert')).toContainText('内容已完整保留');
  await expect(page.getByRole('button', { name: '生成演示预览' })).toBeDisabled();
  await page.getByRole('button', { name: '读取示例选区' }).click();
  await page.getByRole('button', { name: '生成演示预览' }).click();
  await page.getByRole('button', { name: '编辑', exact: true }).click();
  await page.getByLabel('编辑结果').fill('');
  await expect(page.getByLabel('编辑结果')).toBeVisible();
  await expect(page.getByRole('button', { name: '替换已读取选区' })).toBeDisabled();
  await page.getByLabel('编辑结果').fill('修订后的正文');
  await page.getByRole('button', { name: '完成编辑' }).click();
  await expect(page.getByRole('button', { name: '替换已读取选区' })).toBeEnabled();
});

test('a failed mocked remote-model regeneration keeps the last successful result and sends only visible input', async ({
  page,
}) => {
  await page.addInitScript(() => {
    localStorage.setItem(
      'triobridge.profiles.v1',
      JSON.stringify({
        version: 1,
        activeId: 'test',
        profiles: [
          {
            id: 'test',
            name: '合成测试模型',
            baseUrl: 'https://example.test/v1',
            model: 'test',
            apiKey: '',
            rememberKey: false,
            timeoutSeconds: 5,
          },
        ],
      }),
    );
  });
  let calls = 0;
  await page.route('https://example.test/v1/chat/completions', async (route) => {
    const request = route.request().postDataJSON();
    const input = JSON.parse(request.messages[1].content);
    expect(input.documentContent).toBe('仅此内容可以发送');
    await route.fulfill({
      status: ++calls === 1 ? 200 : 503,
      contentType: 'application/json',
      body: JSON.stringify({ choices: [{ message: { content: '第一次成功结果' } }] }),
    });
  });
  await page.goto('/');
  await page.getByLabel('参考内容').fill('仅此内容可以发送');
  await page.getByRole('button', { name: '生成预览', exact: true }).click();
  await expect(page.locator('.result-text')).toHaveText('第一次成功结果');
  await page.getByRole('button', { name: '生成预览', exact: true }).click();
  await expect(page.getByRole('alert')).toContainText('HTTP 503');
  await expect(page.locator('.result-text')).toHaveText('第一次成功结果');
  expect(calls).toBe(2);
  expect(await page.evaluate(() => localStorage.getItem('triobridge.profiles.v1'))).not.toContain(
    '仅此内容',
  );
});

for (const width of [320, 360, 768, 1280]) {
  test(`${width}px: complex table stays inside the workbench and its confirmation`, async ({
    page,
  }) => {
    await page.setViewportSize({ width, height: 900 });
    await page.goto('/');
    await page.getByRole('button', { name: 'Excel', exact: true }).click();
    await page.getByRole('button', { name: '创建数据表', exact: true }).click();
    await page.getByLabel('具体要求').fill('创建复杂数据表');
    await page.getByRole('button', { name: '生成演示预览' }).click();
    await expect(page.locator('.table-preview tbody tr')).toHaveCount(48);
    expect(await page.evaluate(() => document.documentElement.scrollWidth <= innerWidth)).toBe(
      true,
    );
    await page.getByRole('button', { name: '创建新工作表' }).click();
    expect(
      await page
        .getByRole('dialog')
        .evaluate((element) => element.scrollWidth <= element.clientWidth),
    ).toBe(true);
    await page.screenshot({
      path: `test-results/complex-table-${width}-${test.info().project.name}.png`,
      fullPage: true,
    });
  });
}
