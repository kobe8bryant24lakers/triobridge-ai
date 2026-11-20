import { describe, it, expect, vi } from 'vitest';
import { render, screen } from '@testing-library/react';
import userEvent from '@testing-library/user-event';
import App from '../src/App';
import { Workbench } from '../src/components/Workbench';
import type { HostBridge } from '../src/core/types';

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
    await user.type(screen.getByLabelText('服务名称'), 'My model');
    await user.type(screen.getByLabelText('模型 ID'), 'my-model');
    await user.type(screen.getByLabelText(/API Key 本地/), 'test-private-key');
    await user.click(screen.getByRole('button', { name: '保存配置' }));
    expect(await screen.findByText('模型配置已保存，可返回工作台开始使用。')).toBeInTheDocument();
    expect(localStorage.getItem('triobridge.profiles.v1')).not.toContain('test-private-key');
  });
});
