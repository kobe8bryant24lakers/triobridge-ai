import { afterAll, afterEach, beforeAll, describe, expect, it, vi } from 'vitest';
import { fireEvent, render, screen, waitFor } from '@testing-library/react';
import userEvent from '@testing-library/user-event';
import { Workbench } from '../src/components/Workbench';
import { complete } from '../src/core/model-client';
import { emptyProfile } from '../src/core/profile-store';
import { parseReport } from '../src/core/structured-output';
import type { HostBridge } from '../src/core/types';

const longText = '完整改写内容，保留事实。\n'.repeat(10_000) + '这是结果末尾。';
const profile = {
  ...emptyProfile(),
  name: '长文本模型',
  model: 'local-model',
  baseUrl: 'http://localhost:8001/v1',
  apiFormat: 'anthropic' as const,
};
const response = () =>
  Response.json({
    stop_reason: 'end_turn',
    content: [{ type: 'text', text: longText }],
    usage: { input_tokens: 50, output_tokens: 80_000 },
  });

afterEach(() => vi.unstubAllGlobals());

// jsdom does not implement the native dialog methods used by ConfirmWrite.
const dialogMethods = ['showModal', 'close'] as const;
const originalDialogMethods = dialogMethods.map((name) =>
  Object.getOwnPropertyDescriptor(HTMLDialogElement.prototype, name),
);
beforeAll(() => {
  dialogMethods.forEach((name) =>
    Object.defineProperty(HTMLDialogElement.prototype, name, {
      configurable: true,
      value: function (this: HTMLDialogElement) {
        this.open = name === 'showModal';
      },
    }),
  );
});
afterAll(() => {
  dialogMethods.forEach((name, index) => {
    const descriptor = originalDialogMethods[index];
    if (descriptor) Object.defineProperty(HTMLDialogElement.prototype, name, descriptor);
    else Reflect.deleteProperty(HTMLDialogElement.prototype, name);
  });
});

describe('complete long plain-text results', () => {
  it('returns all text beyond 100,000 characters from a single model response', async () => {
    const fetcher = vi.fn().mockResolvedValue(response());
    expect(longText.length).toBeGreaterThan(100_000);
    await expect(
      complete(profile, [{ role: 'user', content: '改写原文' }], undefined, fetcher),
    ).resolves.toBe(longText);
    expect(fetcher).toHaveBeenCalledTimes(1);
    expect(fetcher.mock.calls[0][0]).toBe('http://localhost:8001/v1/messages');
    expect(JSON.parse(fetcher.mock.calls[0][1].body)).not.toHaveProperty('max_tokens');
  });

  it('previews, edits, confirms and replaces a long Word result without another model call', async () => {
    const fetcher = vi.fn().mockImplementation(async () => response());
    vi.stubGlobal('fetch', fetcher);
    const replace = vi.fn().mockResolvedValue(undefined);
    const source = { text: '原文内容', label: 'Word · 当前选区', replace };
    const bridge: HostBridge = {
      host: 'word',
      mode: 'office',
      capabilities: { read: true, replace: true, insert: true, formula: false, slides: false },
      readSelection: vi.fn().mockResolvedValue(source),
      prepareWrite: vi.fn(async () => ({
        label: source.label,
        before: source.text,
        commit: replace,
      })),
    };
    const user = userEvent.setup();
    render(
      <Workbench
        bridge={bridge}
        profiles={{ profiles: [profile], activeId: profile.id }}
        onChooseModel={vi.fn()}
        onSettings={vi.fn()}
        onBusyChange={vi.fn()}
      />,
    );
    await user.click(screen.getByRole('button', { name: '读取选区' }));
    await user.click(screen.getByRole('button', { name: '生成预览' }));
    await waitFor(() =>
      expect(screen.getByRole('region', { name: '结果内容' }).textContent).toBe(longText),
    );
    expect(screen.queryByRole('alert')).not.toBeInTheDocument();
    await user.click(screen.getByRole('button', { name: '编辑' }));
    const edited = longText + '\n人工核对后的补充。';
    fireEvent.change(screen.getByLabelText('编辑结果'), { target: { value: edited } });
    await user.click(screen.getByRole('button', { name: '完成编辑' }));
    await user.click(screen.getByRole('button', { name: '替换已读取选区' }));
    expect(await screen.findByRole('region', { name: '即将写入的结果' })).toHaveTextContent(
      '人工核对后的补充。',
    );
    expect(screen.getByRole('region', { name: '即将写入的结果' }).textContent).toBe(edited);
    expect(replace).not.toHaveBeenCalled();
    await user.click(screen.getByRole('button', { name: '确认写入' }));
    await waitFor(() => expect(replace).toHaveBeenCalledWith(edited));
    expect(fetcher).toHaveBeenCalledTimes(1);
    expect(fetcher.mock.calls[0][0]).toBe('http://localhost:8001/v1/messages');
    expect(JSON.parse(fetcher.mock.calls[0][1].body)).not.toHaveProperty('max_tokens');
    expect(bridge.prepareWrite).toHaveBeenCalledWith('replace', source, edited);
  });

  it('retains structured report validation independently of plain-text generation', () => {
    expect(() => parseReport(longText)).toThrow('结构化结果超过 100,000 字符');
  });
});
