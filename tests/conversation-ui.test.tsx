import { act, fireEvent, render, screen, waitFor } from '@testing-library/react';
import userEvent from '@testing-library/user-event';
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import App from '../src/App';
import { Conversation } from '../src/components/Conversation';
import { replyToConversation } from '../src/core/conversation';
import { emptyProfile, saveProfiles } from '../src/core/profile-store';
import { parseSharedState } from '../src/core/shared-model-client';
import type { HostKind } from '../src/core/types';
import { setUiLanguage } from '../src/i18n';
import { createDemoBridge } from '../src/office/demo';

vi.mock('../src/core/conversation', async (importOriginal) => ({
  ...(await importOriginal<typeof import('../src/core/conversation')>()),
  replyToConversation: vi.fn(),
}));

const reply = vi.mocked(replyToConversation);

function modelState() {
  const profile = {
    ...emptyProfile(),
    name: 'Test model',
    model: 'test-model',
    baseUrl: 'http://localhost:8001/v1',
  };
  return { profiles: [profile], activeId: profile.id };
}

function renderConversation(profiles = modelState()) {
  const onBusyChange = vi.fn();
  const onSettings = vi.fn();
  const result = render(
    <Conversation
      host="powerpoint"
      profiles={profiles}
      onChooseModel={vi.fn()}
      onSettings={onSettings}
      onBusyChange={onBusyChange}
    />,
  );
  return { ...result, profiles, onBusyChange, onSettings };
}

async function send(message: string) {
  fireEvent.change(screen.getByLabelText('你的问题'), { target: { value: message } });
  await userEvent.click(screen.getByRole('button', { name: '发送' }));
}

beforeEach(() => {
  act(() => setUiLanguage('zh-CN'));
  reply.mockReset();
});

afterEach(() => {
  act(() => setUiLanguage('zh-CN'));
});

describe('read-only conversation', () => {
  it('sends prior questions and answers for follow-ups, and clears their context on request', async () => {
    reply
      .mockResolvedValueOnce('先确定受众和汇报目标。')
      .mockResolvedValueOnce('可以按背景、进展和下一步展开。')
      .mockResolvedValueOnce('新的回答。');
    const { profiles } = renderConversation();

    await send('帮我梳理季度汇报的思路。');
    expect(await screen.findByText('先确定受众和汇报目标。')).toBeInTheDocument();
    await send('受众是管理层，建议分成哪些部分？');
    expect(await screen.findByText('可以按背景、进展和下一步展开。')).toBeInTheDocument();
    expect(reply.mock.calls[1][0]).toEqual(profiles.profiles[0]);
    expect(reply.mock.calls[1][1]).toEqual([
      { role: 'user', content: '帮我梳理季度汇报的思路。' },
      { role: 'assistant', content: '先确定受众和汇报目标。' },
      { role: 'user', content: '受众是管理层，建议分成哪些部分？' },
    ]);
    expect(reply.mock.calls[1][2]).toBe('powerpoint');

    await userEvent.click(screen.getByRole('button', { name: '清空对话' }));
    expect(screen.queryByText('先确定受众和汇报目标。')).not.toBeInTheDocument();
    await send('换个话题。');
    expect(await screen.findByText('新的回答。')).toBeInTheDocument();
    expect(reply.mock.calls[2][1]).toEqual([{ role: 'user', content: '换个话题。' }]);
  });

  it('retries a failed question without duplicating it in the transcript or model history', async () => {
    reply
      .mockRejectedValueOnce(new Error('模型暂时不可用。'))
      .mockResolvedValueOnce('已恢复回答。');
    renderConversation();
    await send('这个插件怎么配置模型？');
    expect(await screen.findByRole('alert')).toHaveTextContent('模型暂时不可用。');
    await userEvent.click(screen.getByRole('button', { name: '重试' }));
    expect(await screen.findByText('已恢复回答。')).toBeInTheDocument();
    expect(reply.mock.calls[1][1]).toEqual([{ role: 'user', content: '这个插件怎么配置模型？' }]);
    expect(screen.getAllByText('这个插件怎么配置模型？')).toHaveLength(1);
  });

  it('aborts a stopped response and ignores a transport that resolves after cancellation', async () => {
    let resolve!: (text: string) => void;
    reply.mockImplementationOnce(
      () =>
        new Promise<string>((done) => {
          resolve = done;
        }),
    );
    const { onBusyChange } = renderConversation();
    await send('帮我比较两种汇报结构。');
    expect(screen.getByRole('button', { name: '停止回答' })).toBeEnabled();
    expect(onBusyChange).toHaveBeenLastCalledWith(true);
    const signal = reply.mock.calls[0][3];
    await userEvent.click(screen.getByRole('button', { name: '停止回答' }));
    expect(signal?.aborted).toBe(true);
    await act(async () => resolve('不应出现的过期回复。'));
    expect(screen.queryByText('不应出现的过期回复。')).not.toBeInTheDocument();
    expect(screen.queryByRole('button', { name: '停止回答' })).not.toBeInTheDocument();
    expect(onBusyChange).toHaveBeenLastCalledWith(false);
  });

  it('aborts an in-flight response when the conversation is unmounted', async () => {
    reply.mockImplementationOnce(() => new Promise<string>(() => {}));
    const { unmount } = renderConversation();
    await send('解释一下如何组织材料。');
    const signal = reply.mock.calls[0][3];
    unmount();
    expect(signal?.aborted).toBe(true);
  });

  it('requires a configured model instead of silently returning demo answers', async () => {
    renderConversation({ profiles: [], activeId: '' });
    fireEvent.change(screen.getByLabelText('你的问题'), { target: { value: '你好' } });
    expect(screen.getByRole('button', { name: '发送' })).toBeDisabled();
    expect(reply).not.toHaveBeenCalled();
  });

  it('retains an oversized question and prevents sending until it fits the input limit', () => {
    renderConversation();
    const question = '问'.repeat(4_001);
    fireEvent.change(screen.getByLabelText('你的问题'), { target: { value: question } });
    expect(screen.getByLabelText('你的问题')).toHaveValue(question);
    expect(screen.getByLabelText('你的问题')).toHaveAttribute('aria-invalid', 'true');
    expect(screen.getByRole('button', { name: '发送' })).toBeDisabled();
    expect(screen.getByRole('alert')).toHaveTextContent('4,000');
    fireEvent.change(screen.getByLabelText('你的问题'), { target: { value: question.slice(1) } });
    expect(screen.getByRole('button', { name: '发送' })).toBeEnabled();
    expect(screen.queryByRole('alert')).not.toBeInTheDocument();
    expect(reply).not.toHaveBeenCalled();
  });

  it('uses the newest shared model activation and blocks sends when the shared service goes offline', async () => {
    const shared = parseSharedState({
      version: 2,
      revision: 1,
      activeId: 'a',
      storage: 'macos-keychain',
      profiles: ['a', 'b', 'c'].map((id) => ({
        id,
        name: id,
        baseUrl: 'http://localhost:8001/v1',
        model: id,
        timeoutSeconds: 60,
        hasKey: true,
        revision: 1,
      })),
    });
    let finish!: () => void;
    const onChooseModel = vi.fn(
      () =>
        new Promise<void>((resolve) => {
          finish = resolve;
        }),
    );
    const props = {
      host: 'powerpoint' as const,
      profiles: {
        ...shared,
        mode: 'shared' as const,
        connection: 'ready' as const,
        legacyAvailable: false,
      },
      onChooseModel,
      onSettings: vi.fn(),
      onBusyChange: vi.fn(),
    };
    const { rerender } = render(<Conversation {...props} />);
    await userEvent.selectOptions(screen.getByLabelText('对话模型'), 'b');
    const latest = { ...props.profiles, activeId: 'c', revision: 3 };
    rerender(<Conversation {...props} profiles={latest} />);
    await act(async () => finish());
    expect(screen.getByLabelText('对话模型')).toHaveValue('c');
    reply.mockResolvedValueOnce('使用最新模型回答。');
    await send('请解释这个概念。');
    expect(await screen.findByText('使用最新模型回答。')).toBeInTheDocument();
    expect(reply.mock.calls[0][0].id).toBe('c');
    rerender(<Conversation {...props} profiles={{ ...latest, connection: 'offline' as const }} />);
    fireEvent.change(screen.getByLabelText('你的问题'), { target: { value: '保留离线草稿。' } });
    expect(screen.getByRole('button', { name: '发送' })).toBeDisabled();
    expect(screen.getByLabelText('你的问题')).toHaveValue('保留离线草稿。');
    expect(reply).toHaveBeenCalledTimes(1);
  });

  it('keeps user text when the interface language changes and does not submit an IME confirmation', async () => {
    reply.mockResolvedValueOnce('可以先列出你想表达的结论。');
    renderConversation();
    const input = screen.getByLabelText('你的问题');
    fireEvent.change(input, { target: { value: '请帮我整理想法。' } });
    fireEvent.keyDown(input, { key: 'Enter', code: 'Enter', isComposing: true, keyCode: 229 });
    expect(reply).not.toHaveBeenCalled();
    await userEvent.click(screen.getByRole('button', { name: '发送' }));
    expect(await screen.findByText('可以先列出你想表达的结论。')).toBeInTheDocument();
    fireEvent.change(input, { target: { value: '保留这段中文草稿。' } });
    act(() => setUiLanguage('en'));
    expect(screen.getByLabelText('Your message')).toHaveValue('保留这段中文草稿。');
    expect(screen.getByRole('button', { name: 'Send' })).toBeEnabled();
    expect(screen.getByText('可以先列出你想表达的结论。')).toBeInTheDocument();
  });

  it.each<HostKind>(['word', 'excel', 'powerpoint'])(
    'does not read or modify the %s document even when a chat request asks for an edit',
    async (host) => {
      saveProfiles(modelState());
      const bridge = createDemoBridge(host, vi.fn());
      const read = vi.spyOn(bridge, 'readSelection');
      const prepareWrite = vi.spyOn(bridge, 'prepareWrite');
      const presentationOperations = bridge.presentation
        ? [
            vi.spyOn(bridge.presentation, 'read'),
            vi.spyOn(bridge.presentation, 'apply'),
            vi.spyOn(bridge.presentation, 'prepareInsert'),
            vi.spyOn(bridge.presentation, 'insert'),
            vi.spyOn(bridge.presentation, 'create'),
          ]
        : [];
      reply.mockResolvedValueOnce('可以先讨论修改方案。当前对话不会修改文档。');
      render(<App officeBridge={bridge} />);
      await userEvent.click(screen.getByRole('button', { name: '对话模式' }));
      await send('直接删除当前文档并重新生成全部内容。');
      expect(
        await screen.findByText('可以先讨论修改方案。当前对话不会修改文档。'),
      ).toBeInTheDocument();
      expect(read).not.toHaveBeenCalled();
      expect(prepareWrite).not.toHaveBeenCalled();
      for (const operation of presentationOperations) expect(operation).not.toHaveBeenCalled();
      expect(
        screen.queryByRole('button', { name: /确认写入|插入到光标处|替换选区/ }),
      ).not.toBeInTheDocument();
    },
  );

  it('preserves the discussion and unfinished question when visiting settings or the workbench', async () => {
    saveProfiles(modelState());
    reply.mockResolvedValueOnce('建议先明确听众最关心的问题。');
    render(<App />);
    await userEvent.click(screen.getByRole('button', { name: '对话模式' }));
    await send('我的演讲该如何开场？');
    expect(await screen.findByText('建议先明确听众最关心的问题。')).toBeInTheDocument();
    fireEvent.change(screen.getByLabelText('你的问题'), { target: { value: '还没发送的追问。' } });
    await userEvent.click(screen.getByRole('button', { name: '模型设置' }));
    await userEvent.click(screen.getByRole('button', { name: '对话模式' }));
    expect(screen.getByLabelText('你的问题')).toHaveValue('还没发送的追问。');
    expect(screen.getByText('建议先明确听众最关心的问题。')).toBeVisible();
    await userEvent.click(screen.getByRole('button', { name: '工作台' }));
    await userEvent.click(screen.getByRole('button', { name: '对话模式' }));
    expect(screen.getByLabelText('你的问题')).toHaveValue('还没发送的追问。');
    await waitFor(() => expect(reply).toHaveBeenCalledTimes(1));
  });
});
