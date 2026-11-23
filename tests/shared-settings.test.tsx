import { describe, expect, it, vi } from 'vitest';
import { render, screen, waitFor } from '@testing-library/react';
import userEvent from '@testing-library/user-event';
import { ModelSettings } from '../src/components/ModelSettings';
import type { SharedModelActions, SharedProfileState } from '../src/core/profile-store';

const sharedState = (overrides: Partial<SharedProfileState> = {}): SharedProfileState => ({
  mode: 'shared',
  revision: 4,
  activeId: 'shared-openai',
  storage: 'macos-keychain',
  connection: 'ready',
  legacyAvailable: false,
  profiles: [
    {
      kind: 'shared',
      id: 'shared-openai',
      name: '团队 OpenAI',
      baseUrl: 'https://api.openai.com/v1',
      model: 'configured-model',
      timeoutSeconds: 60,
      hasKey: true,
      revision: 4,
    },
  ],
  ...overrides,
});

function setup(state = sharedState(), overrides: Partial<SharedModelActions> = {}) {
  const actions: SharedModelActions = {
    upsert: vi.fn().mockResolvedValue(sharedState({ revision: 5 })),
    remove: vi.fn().mockResolvedValue(sharedState({ revision: 5, profiles: [], activeId: '' })),
    importProfiles: vi.fn().mockResolvedValue(sharedState({ revision: 5 })),
    listLegacyProfiles: vi.fn().mockResolvedValue([
      {
        id: 'legacy-openai',
        name: '原 OpenAI',
        baseUrl: 'https://api.openai.com/v1',
        model: 'configured-model',
        timeoutSeconds: 60,
        hasKey: true,
      },
      {
        id: 'legacy-fixture',
        name: '本地固定样例',
        baseUrl: '/model-proxy/v1',
        model: 'fixture-model',
        timeoutSeconds: 60,
        hasKey: false,
      },
    ]),
    migrateLocal: vi.fn().mockResolvedValue(sharedState({ revision: 5 })),
    refresh: vi.fn().mockResolvedValue(state),
    ...overrides,
  };
  const onSave = vi.fn();
  const props = { state, sharedActions: actions, onSave, onBack: vi.fn() };
  const view = render(<ModelSettings {...props} />);
  return { ...view, actions, onSave, props, user: userEvent.setup() };
}

describe('shared model settings', () => {
  it('shows saved-key status without reading secrets and preserves the key on save', async () => {
    const { user, actions, onSave } = setup();
    expect(screen.getByText('本机共享')).toBeInTheDocument();
    expect(screen.getByText(/其他 Office 应用也会同步切换/)).toBeInTheDocument();
    expect(screen.getByText('已保存密钥')).toBeInTheDocument();
    expect(screen.queryByLabelText('API Key')).not.toBeInTheDocument();
    expect(screen.queryByLabelText('在本机记住 API Key')).not.toBeInTheDocument();
    await user.click(screen.getByRole('button', { name: '连接并保存' }));
    await screen.findByRole('status');
    expect(actions.upsert).toHaveBeenCalledWith(
      {
        expectedRevision: 4,
        profile: {
          id: 'shared-openai',
          name: '团队 OpenAI',
          baseUrl: 'https://api.openai.com/v1',
          model: 'configured-model',
          timeoutSeconds: 60,
        },
        credential: { action: 'keep' },
        activate: true,
        testConnection: true,
      },
      expect.any(AbortSignal),
    );
    expect(onSave).not.toHaveBeenCalled();
  });

  it('requires an explicit replacement or deletion and never exposes the previous key', async () => {
    const { user, actions } = setup();
    await user.selectOptions(screen.getByLabelText('密钥操作'), 'replace');
    const input = screen.getByLabelText('新 API Key');
    expect(input).toHaveValue('');
    expect(input).toHaveAttribute('type', 'password');
    await user.type(input, 'replacement-synthetic-key');
    await user.click(screen.getByRole('button', { name: '仅保存配置' }));
    await screen.findByRole('status');
    expect(actions.upsert).toHaveBeenLastCalledWith(
      expect.objectContaining({
        credential: { action: 'replace', value: 'replacement-synthetic-key' },
        testConnection: false,
      }),
      expect.any(AbortSignal),
    );
    expect(screen.queryByLabelText('新 API Key')).not.toBeInTheDocument();
    await user.selectOptions(screen.getByLabelText('密钥操作'), 'remove');
    expect(screen.getByText(/三个应用都会受到影响/)).toBeInTheDocument();
    await user.click(screen.getByRole('button', { name: '仅保存配置' }));
    await waitFor(() => expect(actions.upsert).toHaveBeenCalledTimes(2));
    expect(actions.upsert).toHaveBeenLastCalledWith(
      expect.objectContaining({ credential: { action: 'remove' } }),
      expect.any(AbortSignal),
    );
  });

  it('cannot carry an existing credential to a changed service URL', async () => {
    const { user, actions } = setup();
    await user.clear(screen.getByLabelText('API Base URL'));
    await user.type(screen.getByLabelText('API Base URL'), 'https://different.example/v1');
    expect(screen.queryByRole('option', { name: '保留已保存的密钥' })).not.toBeInTheDocument();
    expect(screen.getByLabelText('密钥操作')).toHaveValue('replace');
    await user.click(screen.getByRole('button', { name: '仅保存配置' }));
    expect(actions.upsert).not.toHaveBeenCalled();
    expect(await screen.findByRole('alert')).toHaveTextContent('请填写新密钥');
  });

  it('keeps unsaved edits and their revision when another application updates the catalog', async () => {
    const latest = sharedState({
      revision: 5,
      profiles: [{ ...sharedState().profiles[0], name: '远端名称', revision: 5 }],
    });
    const { user, actions, rerender, props } = setup(undefined, {
      upsert: vi.fn().mockRejectedValue(Object.assign(new Error('Conflict'), { status: 409 })),
      refresh: vi.fn().mockResolvedValue(latest),
    });
    await user.clear(screen.getByLabelText('服务名称'));
    await user.type(screen.getByLabelText('服务名称'), '未保存草稿');
    rerender(<ModelSettings {...props} state={latest} />);
    expect(screen.getByLabelText('服务名称')).toHaveValue('未保存草稿');
    await user.click(screen.getByRole('button', { name: '仅保存配置' }));
    await waitFor(() => expect(actions.upsert).toHaveBeenCalled());
    expect(actions.upsert).toHaveBeenCalledWith(
      expect.objectContaining({
        expectedRevision: 4,
        profile: expect.objectContaining({ name: '未保存草稿' }),
      }),
      expect.any(AbortSignal),
    );
    expect(screen.getByLabelText('服务名称')).toHaveValue('未保存草稿');
    await user.click(screen.getByRole('button', { name: '放弃修改并载入最新配置' }));
    await waitFor(() => expect(screen.getByLabelText('服务名称')).toHaveValue('远端名称'));
  });

  it('disables writes during an outage and reconnects without discarding a dirty draft', async () => {
    const { user, actions, rerender, props } = setup();
    await user.clear(screen.getByLabelText('服务名称'));
    await user.type(screen.getByLabelText('服务名称'), '离线前的草稿');
    rerender(<ModelSettings {...props} state={sharedState({ connection: 'offline' })} />);
    expect(screen.getByRole('button', { name: '连接并保存' })).toBeDisabled();
    expect(screen.getByRole('button', { name: '仅保存配置' })).toBeDisabled();
    expect(screen.getByRole('button', { name: '删除此模型' })).toBeDisabled();
    await user.click(screen.getByRole('button', { name: '重新连接' }));
    await waitFor(() => expect(actions.refresh).toHaveBeenCalledOnce());
    expect(screen.getByLabelText('服务名称')).toHaveValue('离线前的草稿');
    expect(actions.upsert).not.toHaveBeenCalled();
  });

  it('reads only migration metadata on demand and migrates only explicitly selected models', async () => {
    const { user, actions } = setup(sharedState({ legacyAvailable: true }));
    expect(actions.listLegacyProfiles).not.toHaveBeenCalled();
    expect(actions.migrateLocal).not.toHaveBeenCalled();
    await user.click(screen.getByRole('button', { name: '迁移本应用配置' }));
    const selected = await screen.findByRole('checkbox', { name: /原 OpenAI/ });
    expect(screen.getByRole('button', { name: '迁移选中的 0 个模型' })).toBeDisabled();
    expect(actions.migrateLocal).not.toHaveBeenCalled();
    await user.click(selected);
    expect(screen.getByRole('checkbox', { name: /本地固定样例/ })).not.toBeChecked();
    await user.click(screen.getByRole('button', { name: '迁移选中的 1 个模型' }));
    expect(actions.migrateLocal).toHaveBeenCalledWith(
      4,
      ['legacy-openai'],
      expect.any(AbortSignal),
    );
    expect(await screen.findByRole('status')).toHaveTextContent('所选配置已迁移');
  });

  it('retains selected migration sources after an atomic migration failure', async () => {
    const migrateLocal = vi.fn().mockRejectedValue(new Error('迁移失败，原配置已保留。'));
    const { user } = setup(sharedState({ legacyAvailable: true }), { migrateLocal });
    expect(migrateLocal).not.toHaveBeenCalled();
    await user.click(screen.getByRole('button', { name: '迁移本应用配置' }));
    await user.click(await screen.findByRole('checkbox', { name: /原 OpenAI/ }));
    await user.click(screen.getByRole('button', { name: '迁移选中的 1 个模型' }));
    expect(migrateLocal).toHaveBeenCalledWith(4, ['legacy-openai'], expect.any(AbortSignal));
    expect(await screen.findByRole('alert')).toHaveTextContent('原配置已保留');
    expect(screen.getByRole('button', { name: '迁移选中的 1 个模型' })).toBeEnabled();
    expect(screen.getByRole('checkbox', { name: /原 OpenAI/ })).toBeChecked();
    expect(screen.getByText('此应用还有旧配置')).toBeInTheDocument();
  });

  it('distinguishes selecting a model for editing from activating it across applications', async () => {
    const state = sharedState();
    state.profiles.push({ ...state.profiles[0], id: 'second', name: '备用服务' });
    const { user, actions } = setup(state);
    await user.selectOptions(screen.getByLabelText('本机共享模型'), 'second');
    expect(screen.getByLabelText('保存后在三个应用中使用此模型')).not.toBeChecked();
    expect(actions.upsert).not.toHaveBeenCalled();
    await user.click(screen.getByRole('button', { name: '仅保存配置' }));
    await waitFor(() => expect(actions.upsert).toHaveBeenCalledOnce());
    expect(actions.upsert).toHaveBeenCalledWith(
      expect.objectContaining({
        activate: false,
        profile: expect.objectContaining({ id: 'second' }),
      }),
      expect.any(AbortSignal),
    );
  });

  it('marks legacy standalone settings as application-only', () => {
    render(
      <ModelSettings state={{ profiles: [], activeId: '' }} onSave={vi.fn()} onBack={vi.fn()} />,
    );
    expect(screen.getByText(/仅此应用 · 配置不会自动同步/)).toBeInTheDocument();
  });
});
