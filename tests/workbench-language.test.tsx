import { act, fireEvent, render, screen } from '@testing-library/react';
import userEvent from '@testing-library/user-event';
import { afterEach, expect, it, vi } from 'vitest';
import { Workbench } from '../src/components/Workbench';
import { emptyProfile } from '../src/core/profile-store';
import { buildMessages, TASKS } from '../src/core/tasks';
import { setUiLanguage } from '../src/i18n';
import { getComplexSamples } from '../src/office/complex-samples';
import { parseTable, parseReport } from '../src/core/structured-output';
import { parseOutline } from '../src/core/output';
import { createDemoBridge } from '../src/office/demo';

afterEach(() => {
  act(() => setUiLanguage('zh-CN'));
  vi.unstubAllGlobals();
});

it('switches Word controls in place and preserves the source, model name, target, and preview', async () => {
  const profile = {
    ...emptyProfile(),
    name: '模型设置',
    model: 'test-model',
    baseUrl: 'http://localhost:8001/v1',
  };
  const fetcher = vi
    .fn()
    .mockResolvedValue(
      Response.json({ choices: [{ message: { content: '模型设置' }, finish_reason: 'stop' }] }),
    );
  vi.stubGlobal('fetch', fetcher);
  const user = userEvent.setup();
  render(
    <Workbench
      bridge={createDemoBridge('word', vi.fn())}
      profiles={{ profiles: [profile], activeId: profile.id }}
      onChooseModel={vi.fn()}
      onSettings={vi.fn()}
      onBusyChange={vi.fn()}
    />,
  );
  await user.click(screen.getByRole('button', { name: '翻译文本' }));
  await user.selectOptions(screen.getByLabelText('目标语言'), '法语');
  fireEvent.change(screen.getByLabelText('参考内容'), { target: { value: '生成预览' } });
  act(() => setUiLanguage('en'));
  expect(screen.getByRole('button', { name: 'Translate' })).toHaveAttribute('aria-pressed', 'true');
  expect(screen.getByLabelText('Source text')).toHaveValue('生成预览');
  expect(screen.getByRole('option', { name: 'French' })).toHaveProperty('selected', true);
  expect(screen.getByRole('option', { name: /模型设置/ })).toBeInTheDocument();
  await user.click(screen.getByRole('button', { name: 'Generate Preview' }));
  expect(await screen.findByRole('region', { name: 'Result content' })).toHaveTextContent(
    '模型设置',
  );
  const request = JSON.parse(fetcher.mock.calls[0][1].body);
  expect(JSON.parse(request.messages[1].content)).toMatchObject({
    targetLanguage: 'French',
    documentContent: '生成预览',
  });
  expect(request.messages[0].content).toContain('Use English by default');
  act(() => setUiLanguage('zh-CN'));
  expect(screen.getByRole('region', { name: '结果内容' })).toHaveTextContent('模型设置');
  expect(screen.getByLabelText('目标语言')).toHaveValue('法语');
});

it('localizes Excel tasks and keeps explicit user instructions and workbook values intact', () => {
  setUiLanguage('en');
  const task = TASKS.excel.find((t) => t.id === 'table')!;
  expect(task.label).toBe('Create a Table');
  const messages = buildMessages(task, '费用\t16', '用日语生成表格。', '');
  expect(messages[0].content).toContain('Follow any language explicitly requested');
  expect(JSON.parse(messages[1].content)).toMatchObject({
    documentContent: '费用\t16',
    requirements: '用日语生成表格。',
  });
  expect(messages[0].content).toContain('2,000 cells');
  setUiLanguage('zh-CN');
  expect(TASKS.excel.find((t) => t.id === 'table')?.label).toBe('创建数据表');
});

it('creates complete English teaching documents while preserving existing samples when switching', async () => {
  const excel = createDemoBridge('excel', vi.fn(), 'complex');
  const ppt = createDemoBridge('powerpoint', vi.fn());
  setUiLanguage('en');
  const sample = getComplexSamples();
  expect(parseReport(JSON.stringify(sample.report)).sections).toHaveLength(6);
  expect(parseTable(JSON.stringify(sample.table)).rows).toHaveLength(48);
  expect(parseOutline(JSON.stringify({ slides: sample.slides }))).toHaveLength(12);
  expect(Object.values(sample.sources).join(' ')).not.toMatch(/[\u3400-\u9fff]/);
  const before = await excel.readSelection();
  const originalSlides = await ppt.presentation!.read('all');
  expect(originalSlides.slides[0].title).toBe('Quarterly business review');
  setUiLanguage('zh-CN');
  expect((await excel.readSelection()).text).toBe(before.text);
  expect((await ppt.presentation!.read('all')).slides).toEqual(originalSlides.slides);
});
