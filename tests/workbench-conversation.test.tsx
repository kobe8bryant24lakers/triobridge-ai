import { act, fireEvent, render, screen, waitFor } from '@testing-library/react';
import userEvent from '@testing-library/user-event';
import { afterAll, afterEach, beforeAll, beforeEach, describe, expect, it, vi } from 'vitest';
import App from '../src/App';
import { Workbench } from '../src/components/Workbench';
import { complete } from '../src/core/model-client';
import { emptyProfile, saveProfiles } from '../src/core/profile-store';
import type { ChatMessage, HostBridge, HostKind, SelectionSnapshot } from '../src/core/types';
import { setUiLanguage } from '../src/i18n';

vi.mock('../src/core/model-client', async (original) => ({
  ...(await original<typeof import('../src/core/model-client')>()),
  complete: vi.fn(),
}));

const reply = vi.mocked(complete);

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

function modelState() {
  const profile = {
    ...emptyProfile(),
    name: 'Editing test model',
    model: 'test-model',
    baseUrl: 'https://example.test/v1',
  };
  return { profiles: [profile], activeId: profile.id };
}

function officeFixture(host: HostKind = 'word') {
  const snapshot: SelectionSnapshot = {
    text: '原文：项目已完成第一阶段，下一步进行用户测试。',
    label: '测试选区',
    replace: vi.fn(),
    dispose: vi.fn().mockResolvedValue(undefined),
  };
  const commit = vi.fn().mockResolvedValue(undefined);
  const bridge: HostBridge = {
    host,
    mode: 'office',
    capabilities: {
      read: true,
      replace: host === 'word',
      insert: host === 'word',
      formula: host === 'excel',
      slides: host === 'powerpoint',
    },
    readSelection: vi.fn().mockResolvedValue(snapshot),
    prepareWrite: vi.fn().mockResolvedValue({
      label: '测试写入位置',
      before: snapshot.text,
      commit,
    }),
  };
  return { bridge, snapshot, commit };
}

function renderWorkbench(host: HostKind = 'word') {
  const fixture = officeFixture(host);
  const profiles = modelState();
  const rendered = render(
    <Workbench
      bridge={fixture.bridge}
      profiles={profiles}
      onChooseModel={vi.fn()}
      onSettings={vi.fn()}
      onBusyChange={vi.fn()}
      presentationMode={host === 'powerpoint' ? 'text' : undefined}
    />,
  );
  return { ...fixture, ...rendered, profiles };
}

async function generateWord() {
  await userEvent.click(screen.getByRole('button', { name: '读取选区' }));
  fireEvent.change(screen.getByLabelText(/补充要求/), {
    target: { value: '改成正式的项目进度说明。' },
  });
  await userEvent.click(screen.getByRole('button', { name: '生成预览' }));
}

async function refine(value: string) {
  fireEvent.change(screen.getByLabelText('继续调整'), { target: { value } });
  await userEvent.click(screen.getByRole('button', { name: '发送修改要求' }));
}

async function expectPreview(text: string) {
  await waitFor(() =>
    expect(screen.getByRole('region', { name: '结果内容' })).toHaveTextContent(text),
  );
}

function text(messages: ChatMessage[]) {
  return JSON.stringify(messages);
}

beforeEach(() => {
  act(() => setUiLanguage('zh-CN'));
  reply.mockReset();
});

afterEach(() => act(() => setUiLanguage('zh-CN')));

describe('multi-turn document editing', () => {
  it('keeps the original source and every completed turn, writing only after explicit confirmation', async () => {
    reply
      .mockResolvedValueOnce('第一稿：项目第一阶段已经完成。')
      .mockResolvedValueOnce('第二稿：第一阶段完成，准备开展用户测试。')
      .mockResolvedValueOnce('第三稿：第一阶段完成，测试计划将于下周启动。');
    const fixture = renderWorkbench();
    await generateWord();
    await expectPreview('第一稿：项目第一阶段已经完成。');
    await refine('补充下一阶段的测试计划。');
    await expectPreview('第二稿：第一阶段完成，准备开展用户测试。');
    await refine('再说明下周启动，保持简洁。');
    await expectPreview('第三稿：第一阶段完成，测试计划将于下周启动。');

    const messages = reply.mock.calls[2][1];
    expect(messages.map((message) => message.role)).toEqual([
      'system',
      'user',
      'assistant',
      'user',
      'assistant',
      'user',
    ]);
    expect(messages[1]).toEqual(reply.mock.calls[0][1][1]);
    expect(text(messages)).toContain(fixture.snapshot.text);
    expect(messages[2].content).toBe('第一稿：项目第一阶段已经完成。');
    expect(text([messages[3]])).toContain('补充下一阶段的测试计划。');
    expect(messages[4].content).toBe('第二稿：第一阶段完成，准备开展用户测试。');
    expect(text([messages[5]])).toContain('再说明下周启动，保持简洁。');
    expect(fixture.bridge.readSelection).toHaveBeenCalledTimes(1);
    expect(fixture.bridge.prepareWrite).not.toHaveBeenCalled();
    expect(fixture.commit).not.toHaveBeenCalled();

    await userEvent.click(screen.getByRole('button', { name: '替换已读取选区' }));
    expect(fixture.bridge.prepareWrite).toHaveBeenCalledWith(
      'replace',
      fixture.snapshot,
      '第三稿：第一阶段完成，测试计划将于下周启动。',
    );
    expect(fixture.commit).not.toHaveBeenCalled();
    await userEvent.click(screen.getByRole('button', { name: '确认写入' }));
    await waitFor(() =>
      expect(fixture.commit).toHaveBeenCalledWith('第三稿：第一阶段完成，测试计划将于下周启动。'),
    );
  });

  it('uses the manually edited preview as the next assistant turn rather than the old model answer', async () => {
    reply
      .mockResolvedValueOnce('模型生成的初稿。')
      .mockResolvedValueOnce('在手动修改基础上的新稿。');
    renderWorkbench();
    await generateWord();
    await expectPreview('模型生成的初稿。');
    await userEvent.click(screen.getByRole('button', { name: '编辑' }));
    fireEvent.change(screen.getByLabelText('编辑结果'), {
      target: { value: '用户手动补充：测试负责人是小王。' },
    });
    await userEvent.click(screen.getByRole('button', { name: '完成编辑' }));
    await refine('保留负责人，并补充验收安排。');
    await expectPreview('在手动修改基础上的新稿。');
    const messages = reply.mock.calls[1][1];
    expect(messages[messages.length - 2]).toEqual({
      role: 'assistant',
      content: '用户手动补充：测试负责人是小王。',
    });
    expect(text(messages)).not.toContain('模型生成的初稿。');
  });

  it('retains the last preview and follow-up on failure, then retries without adding a failed turn', async () => {
    reply
      .mockResolvedValueOnce('完整的初稿。')
      .mockRejectedValueOnce(new Error('模型暂时不可用。'))
      .mockResolvedValueOnce('成功重试后的第二稿。');
    const fixture = renderWorkbench();
    await generateWord();
    await expectPreview('完整的初稿。');
    await refine('请压缩到一句话。');
    expect(await screen.findByRole('alert')).toHaveTextContent('模型暂时不可用。');
    expect(screen.getByRole('region', { name: '结果内容' })).toHaveTextContent('完整的初稿。');
    expect(screen.getByLabelText('继续调整')).toHaveValue('请压缩到一句话。');
    await userEvent.click(screen.getByRole('button', { name: '发送修改要求' }));
    await expectPreview('成功重试后的第二稿。');
    expect(reply.mock.calls[2][1]).toEqual(reply.mock.calls[1][1]);
    expect(fixture.bridge.prepareWrite).not.toHaveBeenCalled();
  });

  it('ignores a response that arrives after cancellation and keeps the unfinished instruction', async () => {
    let finish!: (value: string) => void;
    reply.mockResolvedValueOnce('停止前保留的初稿。').mockImplementationOnce(
      () =>
        new Promise<string>((resolve) => {
          finish = resolve;
        }),
    );
    renderWorkbench();
    await generateWord();
    await expectPreview('停止前保留的初稿。');
    await refine('为这段内容补充一个标题。');
    const signal = reply.mock.calls[1][2];
    await userEvent.click(screen.getByRole('button', { name: '停止' }));
    expect(signal?.aborted).toBe(true);
    await act(async () => finish('不应替换预览的迟到回复。'));
    expect(screen.getByRole('region', { name: '结果内容' })).toHaveTextContent(
      '停止前保留的初稿。',
    );
    expect(screen.queryByText('不应替换预览的迟到回复。')).not.toBeInTheDocument();
    expect(screen.getByLabelText('继续调整')).toHaveValue('为这段内容补充一个标题。');
    expect(screen.getByRole('button', { name: '发送修改要求' })).toBeEnabled();
  });

  it('starts fresh after changing the source without leaking prior revisions into the new request', async () => {
    reply
      .mockResolvedValueOnce('旧资料初稿。')
      .mockResolvedValueOnce('旧资料第二稿。')
      .mockResolvedValueOnce('新资料初稿。')
      .mockResolvedValueOnce('新资料第二稿。');
    renderWorkbench();
    await generateWord();
    await expectPreview('旧资料初稿。');
    await refine('旧资料专用的要求。');
    await expectPreview('旧资料第二稿。');
    fireEvent.change(screen.getByLabelText('参考内容'), {
      target: { value: '新资料：研发计划已经审批通过。' },
    });
    await userEvent.click(screen.getByRole('button', { name: '生成预览' }));
    await expectPreview('新资料初稿。');
    await refine('继续完善新资料。');
    await expectPreview('新资料第二稿。');
    const request = text(reply.mock.calls[3][1]);
    expect(request).toContain('新资料：研发计划已经审批通过。');
    expect(request).toContain('新资料初稿。');
    expect(request).not.toContain('旧资料');
    expect(request).not.toContain('原文：项目已完成第一阶段');
  });

  it('clears earlier editing turns while keeping the current preview as the new baseline', async () => {
    reply
      .mockResolvedValueOnce('最初预览。')
      .mockResolvedValueOnce('保留的最新预览。')
      .mockResolvedValueOnce('从最新预览继续。');
    renderWorkbench();
    await generateWord();
    await expectPreview('最初预览。');
    await refine('前一轮专用要求。');
    await expectPreview('保留的最新预览。');
    await userEvent.click(screen.getByRole('button', { name: '清空编辑对话' }));
    expect(screen.getByRole('region', { name: '结果内容' })).toHaveTextContent('保留的最新预览。');
    await refine('在当前版本补充一句。');
    await expectPreview('从最新预览继续。');
    const request = text(reply.mock.calls[2][1]);
    expect(request).toContain('保留的最新预览。');
    expect(request).not.toContain('最初预览。');
    expect(request).not.toContain('前一轮专用要求。');
  });

  it('keeps formula constraints and validates a revised Excel formula before exposing write actions', async () => {
    reply.mockResolvedValueOnce('=SUM(B2:B4)').mockResolvedValueOnce('=AVERAGE(B2:B4)');
    const fixture = renderWorkbench('excel');
    fireEvent.change(screen.getByLabelText(/具体要求/), {
      target: { value: '计算 B2:B4 的合计。' },
    });
    await userEvent.click(screen.getByRole('button', { name: '生成预览' }));
    await expectPreview('=SUM(B2:B4)');
    await refine('改为平均值。');
    await expectPreview('=AVERAGE(B2:B4)');
    expect(reply.mock.calls[1][1][0].content).toContain('Excel');
    expect(reply.mock.calls[1][1][2]).toEqual({ role: 'assistant', content: '=SUM(B2:B4)' });
    expect(screen.getByRole('button', { name: '写入选中单元格' })).toBeEnabled();
    expect(fixture.bridge.prepareWrite).not.toHaveBeenCalled();
  });

  it('rejects an invalid structured follow-up without replacing the valid formula or adding it to history', async () => {
    reply
      .mockResolvedValueOnce('=SUM(B2:B4)')
      .mockResolvedValueOnce('这不是一个公式。')
      .mockResolvedValueOnce('=AVERAGE(B2:B4)');
    const fixture = renderWorkbench('excel');
    fireEvent.change(screen.getByLabelText(/具体要求/), {
      target: { value: '计算 B2:B4 的合计。' },
    });
    await userEvent.click(screen.getByRole('button', { name: '生成预览' }));
    await expectPreview('=SUM(B2:B4)');
    await refine('改成平均值公式。');
    expect(await screen.findByRole('alert')).toHaveTextContent('公式');
    expect(screen.getByRole('region', { name: '结果内容' })).toHaveTextContent('=SUM(B2:B4)');
    expect(screen.getByLabelText('继续调整')).toHaveValue('改成平均值公式。');
    await userEvent.click(screen.getByRole('button', { name: '发送修改要求' }));
    await expectPreview('=AVERAGE(B2:B4)');
    expect(reply.mock.calls[2][1]).toEqual(reply.mock.calls[1][1]);
    expect(text(reply.mock.calls[2][1])).not.toContain('这不是一个公式。');
    expect(fixture.bridge.prepareWrite).not.toHaveBeenCalled();
  });

  it('preserves independent editing and chat sessions when moving between modes', async () => {
    saveProfiles(modelState());
    const fixture = officeFixture();
    reply
      .mockResolvedValueOnce('文档模式初稿。')
      .mockResolvedValueOnce('这是对话模式的回答。')
      .mockResolvedValueOnce('文档模式第二稿。');
    render(<App officeBridge={fixture.bridge} />);
    await generateWord();
    await expectPreview('文档模式初稿。');
    fireEvent.change(screen.getByLabelText('继续调整'), {
      target: { value: '保留在文档模式的修改要求。' },
    });
    await userEvent.click(screen.getByRole('button', { name: '对话模式' }));
    fireEvent.change(screen.getByLabelText('你的问题'), {
      target: { value: '插件支持哪些模型？' },
    });
    await userEvent.click(screen.getByRole('button', { name: '发送' }));
    await screen.findByText('这是对话模式的回答。');
    await userEvent.click(screen.getByRole('button', { name: '工作台' }));
    expect(screen.getByLabelText('继续调整')).toHaveValue('保留在文档模式的修改要求。');
    await userEvent.click(screen.getByRole('button', { name: '发送修改要求' }));
    await expectPreview('文档模式第二稿。');
    const request = text(reply.mock.calls[2][1]);
    expect(request).toContain('文档模式初稿。');
    expect(request).not.toContain('插件支持哪些模型');
    expect(request).not.toContain('这是对话模式的回答');
    await userEvent.click(screen.getByRole('button', { name: '对话模式' }));
    expect(screen.getByRole('log', { name: '对话记录' })).toHaveTextContent('这是对话模式的回答。');
  });
});
