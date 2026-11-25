import assert from 'node:assert/strict';
import { _electron as electron } from '@playwright/test';
import { createServer } from 'node:http';
import { once } from 'node:events';
import { randomUUID } from 'node:crypto';
import { mkdtemp, rm } from 'node:fs/promises';
import { join, resolve } from 'node:path';
import { tmpdir } from 'node:os';

const mac = process.platform === 'darwin';
const extracted = resolve(
  'build',
  'native-' + process.platform + '-' + process.arch,
  'verified-archive',
);
const payload = mac ? join(extracted, 'bundle') : extracted;
const executablePath = mac
  ? join(payload, 'TrioBridge AI.app/Contents/MacOS/TrioBridge AI')
  : join(payload, 'Workbench/TrioBridge AI.exe');
const directory = await mkdtemp(join(tmpdir(), 'triobridge-ui-'));
// Only the disposable hosted runner writes a synthetic shared profile, never a developer account.
const testModel = process.argv.includes('--test-model');
if (testModel)
  assert.equal(
    process.env.RUNNER_ENVIRONMENT,
    'github-hosted',
    'Use an ephemeral hosted runner for the loopback model fixture.',
  );
const requests = [];
const fixtureKey = randomUUID();
let model;
try {
  if (testModel) {
    model = createServer(async (request, response) => {
      const body = [];
      for await (const chunk of request) body.push(chunk);
      const input = JSON.parse(Buffer.concat(body));
      const claude = input.model === 'acceptance-second';
      requests.push({
        url: request.url,
        method: request.method,
        model: input.model,
        hasKey: Boolean(request.headers['x-api-key'] || request.headers.authorization),
      });
      const valid =
        request.url === (claude ? '/v1/messages' : '/v1/chat/completions') &&
        request.headers.authorization === undefined &&
        request.headers['x-api-key'] === (claude && !mac ? fixtureKey : undefined) &&
        request.headers['anthropic-version'] === (claude ? '2023-06-01' : undefined) &&
        input.max_tokens === undefined;
      const status = valid ? 200 : 400;
      response.writeHead(status, { 'Content-Type': 'application/json' });
      response.end(
        JSON.stringify(
          claude
            ? { stop_reason: 'end_turn', content: [{ type: 'text', text: 'OK' }] }
            : { choices: [{ message: { role: 'assistant', content: 'OK' } }] },
        ),
      );
    });
    model.listen(0, '127.0.0.1');
    await once(model, 'listening');
  }
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
      if (testModel && host === 'excel') {
        await page
          .getByRole('heading', { name: 'Excel Workspace', exact: true })
          .waitFor({ timeout: 20000 });
        const selector = page.getByRole('combobox', { name: 'Interface language' });
        assert.equal(await selector.inputValue(), 'en');
        await selector.selectOption('zh-CN');
        await page.waitForFunction(
          () =>
            document.documentElement.lang === 'zh-CN' &&
            !document.querySelector('.language-control select').disabled,
        );
      }
      await page.getByRole('heading', { name: heading, exact: true }).waitFor({ timeout: 20000 });
      assert.equal(new URL(page.url()).protocol, 'file:');
      const result = await page.evaluate(async () => {
        const native = window.trioWorkbench;
        const replies = await Promise.all(
          [1, 2, 3].map(() => native.model(crypto.randomUUID(), 'state', '{}')),
        );
        const { status, body } = replies.find((r) => r.status !== 200) || replies[0];
        return {
          host: native.host,
          status,
          detail: body.error?.detail || body.error?.code,
          hasProfileList: Array.isArray(body.profiles),
          activeModel: body.profiles?.find((p) => p.id === body.activeId)?.model,
          resources: performance.getEntriesByType('resource').map((r) => new URL(r.name).protocol),
        };
      });
      assert.equal(result.host, host);
      assert.equal(
        result.status,
        200,
        'Shared catalog first-open check: ' + (result.detail || 'unknown'),
      );
      assert.ok(result.hasProfileList);
      if (testModel && host !== 'word')
        assert.equal(
          result.activeModel,
          host === 'excel' ? 'acceptance-local' : 'acceptance-second',
        );
      assert.ok(result.resources.every((p) => ['file:', 'data:', 'blob:'].includes(p)));
      await page.getByRole('button', { name: '模型设置', exact: true }).first().click();
      await page.getByRole('button', { name: '返回工作台' }).waitFor();
      if (host === 'word' || (testModel && host === 'excel')) {
        const claude = host === 'excel';
        await page.getByRole('button', { name: '新增', exact: true }).click();
        const provider = page.getByLabel('模型服务', { exact: true });
        await provider.selectOption('device');
        const format = page.getByRole('combobox', { name: '接口格式' });
        assert.equal(await format.inputValue(), 'openai');
        if (claude) await format.selectOption('anthropic');
        await page
          .getByLabel('模型 ID', { exact: true })
          .fill(claude ? 'acceptance-second' : 'acceptance-local');
        await page
          .getByLabel('服务名称', { exact: true })
          .fill(claude ? 'Claude 兼容验收' : 'OpenAI 兼容验收');
        const port = model?.address().port || 8012;
        await page
          .getByLabel('API Base URL', { exact: true })
          .fill('http://127.0.0.1:' + port + '/v1');
        assert.equal(await provider.inputValue(), 'device');
        assert.equal(await page.getByLabel('密钥操作', { exact: true }).inputValue(), 'remove');
        if (claude && !mac) {
          await page.getByLabel('密钥操作', { exact: true }).selectOption('replace');
          await page.getByLabel('API Key', { exact: true }).fill(fixtureKey);
        }
        if (testModel) {
          await page.getByRole('button', { name: '连接并保存', exact: true }).click();
          await page
            .getByText('连接成功，配置已在三个应用间共享。', { exact: true })
            .waitFor({ timeout: 60000 });
          assert.deepEqual(requests.at(-1), {
            url: claude ? '/v1/messages' : '/v1/chat/completions',
            method: 'POST',
            model: claude ? 'acceptance-second' : 'acceptance-local',
            hasKey: claude && !mac,
          });
        }
      }
      if (testModel && host === 'powerpoint') {
        assert.equal(await page.getByLabel('接口格式', { exact: true }).inputValue(), 'anthropic');
        const checked = await page.evaluate(async () => {
          const call = (action, payload) =>
            window.trioWorkbench.model(crypto.randomUUID(), action, JSON.stringify(payload));
          const { body: state } = await call('state', {});
          const active = state.profiles.find((p) => p.id === state.activeId);
          const first = state.profiles.find((p) => p.model === 'acceptance-local');
          const completion = await call('complete', {
            profileId: active.id,
            profileRevision: active.revision,
            messages: [{ role: 'user', content: 'Reply with OK.' }],
          });
          const switched = await call('activate', {
            expectedRevision: state.revision,
            id: first.id,
          });
          return {
            count: state.profiles.length,
            completion: completion.status,
            switched: switched.status === 200 && switched.body.activeId === first.id,
          };
        });
        assert.deepEqual(checked, { count: 2, completion: 200, switched: true });
        assert.deepEqual(
          requests.map((r) => [r.model, r.hasKey]),
          [
            ['acceptance-local', false],
            ['acceptance-second', !mac],
            ['acceptance-second', !mac],
          ],
        );
        console.log(
          'Verified two model profiles, cross-host generation, and switching the active model.',
        );
      }
      if (testModel && host === 'word') {
        await page.getByRole('combobox', { name: /Interface language/ }).selectOption('en');
        await page.getByRole('button', { name: 'Back to workspace' }).waitFor();
        await page.waitForFunction(
          () => !document.querySelector('.language-control select').disabled,
        );
        const language = await page.evaluate(
          async () =>
            (await window.trioWorkbench.model(crypto.randomUUID(), 'state', '{}')).body.uiLanguage,
        );
        assert.equal(language, 'en');
      }
      console.log(host + ': packaged file page, host, model catalog, and settings verified');
    } finally {
      await application.close();
    }
  }
} finally {
  if (model?.listening) await new Promise((resolve) => model.close(resolve));
  await rm(directory, { recursive: true, force: true });
}
