import { expect, test } from '@playwright/test';

for (const host of ['word', 'excel', 'powerpoint'] as const) {
  test(`${host} runs its English demo and keeps the draft when switching back to Chinese`, async ({
    page,
  }) => {
    await page.goto('/?demo-host=' + host);
    const language = page.getByRole('combobox', { name: /Interface language/i });
    await language.selectOption('en');
    if (host === 'powerpoint')
      await page.getByRole('button', { name: 'Text and Outlines' }).click();
    await page.getByRole('button', { name: 'Read Sample Selection' }).click();
    const source = page.getByLabel('Source text', { exact: true });
    expect(await source.inputValue()).toMatch(
      host === 'word'
        ? /This quarter/
        : host === 'excel'
          ? /Sales Summary!A1:D5/
          : /Quarterly review/,
    );
    const originalSource = await source.inputValue();
    if (host !== 'word')
      await page
        .getByLabel(/Requirements/)
        .first()
        .fill(host === 'excel' ? 'Sum B2:B4' : 'Create a quarterly review outline');
    await page.getByRole('button', { name: 'Generate Demo Preview' }).click();
    const result = page.getByRole('region', { name: 'Result content' });
    await expect(result).toContainText(
      host === 'word' ? 'This quarter' : host === 'excel' ? '=SUM(B2:B4)' : 'Quarterly review',
    );
    const draft = await result.innerText();
    await page
      .getByRole('button', {
        name:
          host === 'word'
            ? 'Replace Read Selection'
            : host === 'excel'
              ? 'Write to Selected Cell'
              : 'Create Slides',
      })
      .click();
    const confirm = page.getByRole('dialog');
    await expect(confirm.getByRole('heading', { name: 'Confirm Write' })).toBeVisible();
    await confirm.getByRole('button', { name: 'Back to Review' }).click();
    await language.selectOption('zh-CN');
    await expect(page.getByLabel('参考内容', { exact: true })).toHaveValue(originalSource);
    await expect(page.getByRole('region', { name: '结果内容' })).toHaveText(draft, {
      useInnerText: true,
    });
  });
}
