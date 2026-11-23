import { describe, it, expect, vi } from 'vitest';
import { act, render, screen } from '@testing-library/react';
import { StrictMode } from 'react';
import userEvent from '@testing-library/user-event';
import App from '../src/App';
import { Workbench } from '../src/components/Workbench';
import type { HostBridge, SelectionSnapshot, WritePlan } from '../src/core/types';
import { createDemoBridge } from '../src/office/demo';

describe('workflow boundaries', () => {
  it('does not auto-read or generate when the app is opened', async () => {
    render(<App />);
    expect(screen.getByLabelText('参考内容')).toHaveValue('');
    expect(screen.getByRole('button', { name: /生成演示预览/ })).toBeDisabled();
    await userEvent.click(screen.getByRole('button', { name: '读取示例选区' }));
    expect((screen.getByLabelText('参考内容') as HTMLTextAreaElement).value).toContain('本季度');
    expect(screen.getByRole('button', { name: /生成演示预览/ })).toBeEnabled();
  });
  it('does not silently use a demo model inside Office', () => {
    const bridge: HostBridge = {
      host: 'word',
      mode: 'office',
      capabilities: {
        read: true,
        replace: true,
        insert: true,
        formula: false,
        slides: false,
      },
      readSelection: vi.fn(),
      prepareWrite: vi.fn(),
    };
    render(
      <Workbench
        bridge={bridge}
        profiles={{ profiles: [], activeId: '' }}
        onChooseModel={vi.fn()}
        onSettings={vi.fn()}
        onBusyChange={vi.fn()}
      />,
    );
    expect(screen.queryByText('演示模型 · 无需 API Key')).not.toBeInTheDocument();
    expect(screen.getByRole('button', { name: '生成预览' })).toBeDisabled();
    expect(bridge.readSelection).not.toHaveBeenCalled();
  });
  it('saves a model and exposes validation without leaking keys to export', async () => {
    const user = userEvent.setup();
    render(<App />);
    await user.click(screen.getByRole('button', { name: '模型设置' }));
    await user.selectOptions(screen.getByLabelText('模型服务'), 'custom');
    await user.type(screen.getByLabelText('服务名称'), 'My model');
    await user.type(screen.getByLabelText('模型 ID'), 'my-model');
    await user.type(screen.getByLabelText(/API Key 本地/), 'test-private-key');
    await user.click(screen.getByRole('button', { name: '仅保存配置' }));
    expect(await screen.findByText('模型配置已保存，可返回工作台开始使用。')).toBeInTheDocument();
    expect(localStorage.getItem('triobridge.profiles.v1')).not.toContain('test-private-key');
  });
  it('works under React StrictMode cleanup and remount probes', async () => {
    render(
      <StrictMode>
        <App />
      </StrictMode>,
    );
    await userEvent.click(screen.getByRole('button', { name: '读取示例选区' }));
    await userEvent.click(screen.getByRole('button', { name: '生成演示预览' }));
    expect(await screen.findByText('固定演示样例 · 未调用 AI')).toBeInTheDocument();
  });
  it('disposes a tracked source when reading finishes after unmount', async () => {
    let resolve!: (value: SelectionSnapshot) => void;
    const bridge = createDemoBridge('word', vi.fn());
    bridge.readSelection = () =>
      new Promise((done) => {
        resolve = done;
      });
    const { unmount } = render(
      <Workbench
        bridge={bridge}
        profiles={{ profiles: [], activeId: '' }}
        onChooseModel={vi.fn()}
        onSettings={vi.fn()}
        onBusyChange={vi.fn()}
      />,
    );
    await userEvent.click(screen.getByRole('button', { name: '读取示例选区' }));
    unmount();
    const dispose = vi.fn().mockResolvedValue(undefined);
    await act(async () => resolve({ text: 'late', label: 'late', dispose }));
    expect(dispose).toHaveBeenCalledTimes(1);
  });
  it('disposes a prepared write anchor if the workbench is removed before confirmation', async () => {
    let resolve!: (value: WritePlan) => void;
    const bridge = createDemoBridge('word', vi.fn());
    bridge.prepareWrite = () =>
      new Promise((done) => {
        resolve = done;
      });
    const { unmount } = render(
      <Workbench
        bridge={bridge}
        profiles={{ profiles: [], activeId: '' }}
        onChooseModel={vi.fn()}
        onSettings={vi.fn()}
        onBusyChange={vi.fn()}
      />,
    );
    await userEvent.click(screen.getByRole('button', { name: '读取示例选区' }));
    await userEvent.click(screen.getByRole('button', { name: '生成演示预览' }));
    await screen.findByText('固定演示样例 · 未调用 AI');
    await userEvent.click(screen.getByRole('button', { name: '插入到光标处' }));
    unmount();
    const dispose = vi.fn().mockResolvedValue(undefined);
    const commit = vi.fn();
    await act(async () => resolve({ label: 'late', before: '', commit, dispose }));
    expect(dispose).toHaveBeenCalledTimes(1);
    expect(commit).not.toHaveBeenCalled();
  });
  it('does not overwrite a newer shared activation when an earlier selection finishes', async () => {
    const { parseSharedState } = await import('../src/core/shared-model-client');
    const shared = parseSharedState({
      version: 2,
      revision: 1,
      activeId: 'a',
      storage: 'macos-keychain',
      profiles: ['a', 'b', 'c'].map((id) => ({
        id,
        name: id,
        baseUrl: 'https://example.com/v1',
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
      bridge: createDemoBridge('word', vi.fn()),
      onChooseModel,
      onSettings: vi.fn(),
      onBusyChange: vi.fn(),
      profiles: {
        ...shared,
        mode: 'shared' as const,
        connection: 'ready' as const,
        legacyAvailable: false,
      },
    };
    const { rerender } = render(<Workbench {...props} />);
    await userEvent.selectOptions(screen.getByLabelText('当前模型'), 'b');
    rerender(<Workbench {...props} profiles={{ ...props.profiles, activeId: 'c', revision: 3 }} />);
    expect(screen.getByLabelText('当前模型')).toHaveValue('c');
    await act(async () => finish());
    expect(screen.getByLabelText('当前模型')).toHaveValue('c');
  });
});
