import assert from 'node:assert/strict';
import { _electron as electron } from '@playwright/test';
import { mkdtemp, rm } from 'node:fs/promises';
import { join, resolve } from 'node:path';
import { tmpdir } from 'node:os';

const mac = process.platform === 'darwin';
const payload = resolve('build', 'native-' + process.platform + '-' + process.arch, 'bundle');
const executablePath = mac
  ? join(payload, 'TrioBridge AI.app/Contents/MacOS/TrioBridge AI')
  : join(payload, 'Workbench/TrioBridge AI.exe');
const directory = await mkdtemp(join(tmpdir(), 'triobridge-ui-'));
try {
  for (const [host, heading] of [
    ['word', 'Word 工作台'],
    ['excel', 'Excel 工作台'],
    ['powerpoint', 'PowerPoint 工作台'],
  ]) {
    const application = await electron.launch({
      executablePath,
      args: ['--office=' + host, '--user-data-dir=' + join(directory, host)],
      timeout: 50000,
    });
    try {
      const page = await application.firstWindow({ timeout: 30000 });
      await page.getByRole('heading', { name: heading, exact: true }).waitFor({ timeout: 20000 });
      assert.equal(new URL(page.url()).protocol, 'file:');
      const result = await page.evaluate(async () => {
        const native = window.trioWorkbench;
        const { status, body } = await native.model(crypto.randomUUID(), 'state', '{}');
        return {
          host: native.host,
          status,
          hasProfileList: Array.isArray(body.profiles),
          resources: performance.getEntriesByType('resource').map((r) => new URL(r.name).protocol),
        };
      });
      assert.equal(result.host, host);
      assert.equal(result.status, 200);
      assert.ok(result.hasProfileList);
      assert.ok(result.resources.every((p) => ['file:', 'data:', 'blob:'].includes(p)));
      await page.getByRole('button', { name: '模型设置', exact: true }).first().click();
      await page.getByRole('button', { name: '返回工作台' }).waitFor();
      if (host === 'word') {
        await page.getByRole('button', { name: '新增', exact: true }).click();
        const provider = page.getByLabel('模型服务', { exact: true });
        await provider.selectOption('device');
        await page.getByLabel('模型 ID', { exact: true }).fill('acceptance-local');
        await page.getByLabel('API Base URL', { exact: true }).fill('http://127.0.0.1:8012/v1');
        assert.equal(await provider.inputValue(), 'device');
        assert.equal(await page.getByLabel('密钥操作', { exact: true }).inputValue(), 'remove');
      }
      console.log(host + ': packaged file page, host, model catalog, and settings verified');
    } finally {
      await application.close();
    }
  }
} finally {
  await rm(directory, { recursive: true, force: true });
}
