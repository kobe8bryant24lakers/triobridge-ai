import { afterEach, describe, expect, it, vi } from 'vitest';
import { render, screen, waitFor } from '@testing-library/react';
import userEvent from '@testing-library/user-event';
import App from '../src/App';
import { ModelSettings } from '../src/components/ModelSettings';
import { PROFILE_KEY } from '../src/core/profile-store';

afterEach(() => vi.unstubAllGlobals());

describe('quick model setup', () => {
  it('connects and saves with only a provider and key, without persisting the key by default', async () => {
    const fetcher = vi
      .fn()
      .mockResolvedValue(
        new Response(JSON.stringify({ choices: [{ message: { content: 'OK' } }] })),
      );
    vi.stubGlobal('fetch', fetcher);
    const user = userEvent.setup();
    render(<App />);
    await user.click(screen.getByRole('button', { name: '模型设置' }));
    expect(screen.queryByLabelText('API Base URL')).not.toBeInTheDocument();
    expect(screen.queryByLabelText('已保存的模型')).not.toBeInTheDocument();
    await user.selectOptions(screen.getByLabelText('模型服务'), 'openai');
    await user.type(screen.getByLabelText('API Key'), ' test-private-key ');
    expect(fetcher).not.toHaveBeenCalled();
    await user.click(screen.getByRole('button', { name: '连接并保存' }));
    expect(await screen.findByRole('status')).toHaveTextContent('连接成功，配置已保存');
    expect(fetcher).toHaveBeenCalledOnce();
    const [url, options] = fetcher.mock.calls[0];
    expect(url).toBe('https://api.openai.com/v1/chat/completions');
    expect(JSON.parse(options.body)).toMatchObject({
      model: 'gpt-4.1-mini',
      messages: [{ role: 'user', content: 'Reply with OK.' }],
    });
    const saved = localStorage.getItem(PROFILE_KEY)!;
    expect(JSON.parse(saved).profiles).toHaveLength(1);
    expect(saved).not.toContain('test-private-key');
    await user.click(screen.getByRole('button', { name: '返回工作台' }));
    expect(screen.getByRole('option', { name: 'OpenAI · gpt-4.1-mini' })).toBeInTheDocument();
  });

  it('keeps a failed connection unsaved and clears credentials when changing destinations', async () => {
    const fetcher = vi.fn().mockResolvedValue(new Response('{}', { status: 401 }));
    vi.stubGlobal('fetch', fetcher);
    const save = vi.fn();
    const user = userEvent.setup();
    render(<ModelSettings state={{ profiles: [], activeId: '' }} onSave={save} onBack={vi.fn()} />);
    await user.selectOptions(screen.getByLabelText('模型服务'), 'deepseek');
    await user.click(screen.getByRole('button', { name: '连接并保存' }));
    expect(fetcher).not.toHaveBeenCalled();
    await user.type(screen.getByLabelText('API Key'), 'first-service-key');
    await user.click(screen.getByRole('button', { name: '连接并保存' }));
    expect(await screen.findByRole('alert')).toHaveTextContent('API Key 无效');
    expect(save).not.toHaveBeenCalled();
    await user.click(screen.getByLabelText('在本机记住 API Key'));
    await user.selectOptions(screen.getByLabelText('模型服务'), 'openai');
    expect(screen.getByLabelText('API Key')).toHaveValue('');
    expect(screen.getByLabelText('在本机记住 API Key')).not.toBeChecked();
    await user.type(screen.getByLabelText('API Key'), 'second-service-key');
    await user.click(screen.getByRole('button', { name: /高级设置/ }));
    await user.clear(screen.getByLabelText('API Base URL'));
    expect(screen.getByLabelText(/API Key 本地/)).toHaveValue('');
  });

  it('cancels connection without saving and allows retry', async () => {
    vi.stubGlobal(
      'fetch',
      vi.fn(
        (_url, options) =>
          new Promise((_resolve, reject) => {
            options.signal.addEventListener('abort', () =>
              reject(new DOMException('Aborted', 'AbortError')),
            );
          }),
      ),
    );
    const save = vi.fn();
    const user = userEvent.setup();
    render(<ModelSettings state={{ profiles: [], activeId: '' }} onSave={save} onBack={vi.fn()} />);
    await user.selectOptions(screen.getByLabelText('模型服务'), 'openai');
    await user.type(screen.getByLabelText('API Key'), 'test-key');
    await user.click(screen.getByRole('button', { name: '连接并保存' }));
    expect(screen.getByLabelText('模型服务')).toBeDisabled();
    await user.click(screen.getByRole('button', { name: '取消测试' }));
    await waitFor(() => expect(screen.getByRole('button', { name: '连接并保存' })).toBeEnabled());
    expect(save).not.toHaveBeenCalled();
  });
});
