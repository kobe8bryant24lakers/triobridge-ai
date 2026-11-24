import { test, expect } from '@playwright/test';
import { mkdir, readFile, writeFile } from 'node:fs/promises';
import { strFromU8, unzipSync } from 'fflate';

test.beforeEach(async ({ page }) => {
  await page.goto('/');
  await page.getByRole('button', { name: 'PowerPoint', exact: true }).click();
});
async function capture(page: import('@playwright/test').Page, name: string) {
  if (process.env.PPT_REVIEW_SCREENSHOTS !== '1') return;
  await mkdir('.impeccable/review/ppt-materials-final', { recursive: true });
  await page.evaluate(() => window.scrollTo(0, 0));
  await page.screenshot({
    path: `.impeccable/review/ppt-materials-final/${name}.png`,
    fullPage: true,
    animations: 'disabled',
  });
}
test('materials are reviewed locally and full decks export editable native pages; changed inputs invalidate export', async ({
  page,
}, testInfo) => {
  const pickerEvent = page.waitForEvent('filechooser');
  await page.getByRole('button', { name: '选择素材目录', exact: true }).click();
  const picker = await pickerEvent;
  expect(await picker.element().getAttribute('webkitdirectory')).not.toBeNull();
  const folderPath = testInfo.outputPath('materials');
  await mkdir(folderPath, { recursive: true });
  await writeFile(`${folderPath}/overview.md`, '合成目录素材：项目甲季度复盘。');
  await picker.setFiles(folderPath);
  await expect(page.getByRole('button', { name: '素材上下文 · 1 份已选' })).toBeVisible();
  const requests: string[] = [];
  page.on('request', (r) => {
    if (r.method() === 'POST') requests.push(r.url());
  });
  await page.getByLabel('选择素材文件', { exact: true }).setInputFiles([
    {
      name: '项目说明.md',
      mimeType: 'text/markdown',
      buffer: Buffer.from('项目甲：季度经营汇报，给管理层演示。'),
    },
    {
      name: '经营.csv',
      mimeType: 'text/csv',
      buffer: Buffer.from('月份,营收\n一月,128\n二月,156'),
    },
    {
      name: 'private.txt',
      mimeType: 'text/plain',
      buffer: Buffer.from('api_key="private-material-example"'),
    },
  ]);
  await expect(page.getByRole('button', { name: '素材上下文 · 2 份已选' })).toBeVisible();
  expect(requests).toEqual([]);
  await page.getByText('1 项资料读取提示').click();
  await expect(page.getByText(/检测到疑似凭据/)).toBeVisible();
  await page.getByLabel('演示要求', { exact: true }).fill('生成项目经营汇报，8页。');
  await page.getByRole('button', { name: '生成整套 PPT', exact: true }).click();
  await expect(page.getByRole('region', { name: '整稿预览' })).toBeVisible();
  await capture(page, 'generated-preview');
  const downloadEvent = page.waitForEvent('download');
  await page.getByRole('button', { name: '下载完整 PPT', exact: true }).click();
  const download = await downloadEvent;
  const bytes = await readFile((await download.path())!);
  const files = unzipSync(new Uint8Array(bytes));
  expect(Object.keys(files).filter((k) => /^ppt\/slides\/slide\d+\.xml$/.test(k))).toHaveLength(8);
  expect(strFromU8(files['ppt/slides/slide2.xml'])).toContain('本季度进展');
  expect(strFromU8(files['ppt/slides/slide3.xml'])).toContain('复盘改进');
  await page.getByLabel('演示要求', { exact: true }).fill('新的受众和要求');
  await expect(page.getByRole('button', { name: '下载完整 PPT', exact: true })).toBeDisabled();
  await expect(page.getByText(/素材、要求或模型已变化/)).toBeVisible();
  await capture(page, 'outdated-preview');
  expect(requests).toEqual([]);
});
test('existing pages support selective real object changes, quick replacements and stale protection in demo bridge', async ({
  page,
}) => {
  await page.getByRole('button', { name: '编辑现有', exact: true }).click();
  await page.getByRole('button', { name: '读取页面与对象', exact: true }).click();
  await page.getByText('批量编辑与排版工具', { exact: true }).click();
  await page.getByLabel('查找文字', { exact: true }).fill('24 小时');
  await page.getByLabel('替换为', { exact: true }).fill('12 小时');
  await page.getByRole('button', { name: '预览批量替换', exact: true }).click();
  await expect(page.locator('.change-row')).toHaveCount(1);
  await expect(page.locator('.change-after')).toContainText('12 小时');
  await page.getByRole('button', { name: '应用所选 1 项改动' }).click();
  await capture(page, 'edit-confirmation');
  await page.getByRole('button', { name: '确认应用改动', exact: true }).click();
  await expect(page.getByText('改动已应用并核对。继续编辑前请重新读取。')).toBeVisible();
  await expect(page.getByRole('button', { name: '应用所选 1 项改动' })).toBeDisabled();
  await page.getByRole('button', { name: '重新读取页面', exact: true }).click();
  await page.getByLabel('编辑要求', { exact: true }).fill('根据素材更新标题并精简文字');
  await page.getByRole('button', { name: '生成局部修改方案', exact: true }).click();
  await page.getByLabel('编辑要求', { exact: true }).fill('改成另一个要求');
  await expect(page.getByRole('button', { name: '应用所选 1 项改动' })).toBeDisabled();
});
test('data charts preserve underlying zero/negative/missing values and reject invalid pie data', async ({
  page,
}) => {
  await page.getByRole('button', { name: '数据图表', exact: true }).click();
  await page
    .getByLabel('含表头的数据', { exact: true })
    .fill('月份\t营收\t利润率\n一月\t0\t12.5%\n二月\t-2.5\t20%\n三月\t\t27.5%');
  await page.getByRole('combobox', { name: '图表类型', exact: true }).selectOption('line');
  await expect(page.getByRole('region', { name: '图表预览' })).toBeVisible();
  const downloadEvent = page.waitForEvent('download');
  await page.getByRole('button', { name: '下载图表 PPT', exact: true }).click();
  const download = await downloadEvent,
    files = unzipSync(new Uint8Array(await readFile((await download.path())!)));
  const chart = strFromU8(files['ppt/charts/chart1.xml']);
  expect(chart).toContain('lineChart');
  expect(chart).toContain('<c:v>-2.5</c:v>');
  expect(chart).toContain('<c:v>0</c:v>');
  expect(Object.keys(files).some((k) => k.endsWith('.xlsx'))).toBe(true);
  const embedded = unzipSync(files[Object.keys(files).find((k) => k.endsWith('.xlsx'))!]);
  const sheet = strFromU8(embedded['xl/worksheets/sheet1.xml']);
  expect(sheet).toMatch(/r="B2"[^>]*><v>0<\/v>/);
  expect(sheet).toMatch(/r="B4"[^>]*\/>/);
  await page.getByRole('combobox', { name: '图表类型', exact: true }).selectOption('bar');
  await capture(page, 'bar-chart');
  await page.getByRole('combobox', { name: '图表类型', exact: true }).selectOption('pie');
  await expect(page.getByText(/饼图和圆环图需要非负完整数据/)).toBeVisible();
  await expect(page.getByRole('button', { name: '下载图表 PPT', exact: true })).not.toBeVisible();
  await capture(page, 'chart-error');
});
test('narrow Office pane keeps workflows and previews reachable without horizontal scrolling', async ({
  page,
}) => {
  await page.setViewportSize({ width: 320, height: 720 });
  await page.getByLabel('演示要求', { exact: true }).fill('季度工作汇报');
  await page.getByRole('button', { name: '生成整套 PPT', exact: true }).click();
  await page.getByRole('combobox', { name: '预览页面', exact: true }).selectOption('1');
  await expect(page.locator('.slide-columns')).toBeVisible();
  expect(await page.evaluate(() => document.documentElement.scrollWidth <= innerWidth)).toBe(true);
  await page.getByRole('button', { name: '编辑现有', exact: true }).click();
  await page.getByRole('button', { name: '读取页面与对象', exact: true }).click();
  await page.getByLabel('编辑要求', { exact: true }).fill('精简标题');
  await page.getByRole('button', { name: '生成局部修改方案', exact: true }).click();
  expect(await page.evaluate(() => document.documentElement.scrollWidth <= innerWidth)).toBe(true);
  await capture(page, 'narrow-edit');
});
