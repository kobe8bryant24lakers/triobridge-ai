import { useEffect, useRef, useState } from 'react';
import {
  ArrowLeft,
  Check,
  Download,
  KeyRound,
  LoaderCircle,
  Plus,
  Trash2,
  Upload,
  Wifi,
} from 'lucide-react';
import {
  emptyProfile,
  exportProfiles,
  importProfiles,
  type ProfileState,
} from '../core/profile-store';
import { testConnection, validateProfile } from '../core/model-client';
import { MODEL_PRESETS, presetFor } from '../core/model-presets';
import { officeError } from '../office/bridge';
import type { ModelProfile } from '../core/types';

interface Props {
  state: ProfileState;
  onSave: (state: ProfileState) => void;
  onBack: () => void;
}

export function ModelSettings({ state, onSave, onBack }: Props) {
  const [draft, setDraft] = useState<ModelProfile>(() => ({
    ...(state.profiles.find((p) => p.id === state.activeId) || emptyProfile()),
  }));
  const [status, setStatus] = useState<{ text: string; error?: boolean }>();
  const [provider, setProvider] = useState(() => presetFor(draft));
  const [advanced, setAdvanced] = useState(() => presetFor(draft) === 'custom');
  const [testing, setTesting] = useState(false);
  const [deleting, setDeleting] = useState(false);
  const controller = useRef<AbortController | null>(null);
  const fileInput = useRef<HTMLInputElement>(null);
  useEffect(
    () => () => {
      controller.current?.abort();
    },
    [],
  );
  const existing = state.profiles.some((p) => p.id === draft.id);
  const preset = MODEL_PRESETS.find((p) => p.id === provider);
  function selectProfile(profile: ModelProfile) {
    setDraft({ ...profile });
    setProvider(presetFor(profile));
    setAdvanced(presetFor(profile) === 'custom');
    setStatus(undefined);
    setDeleting(false);
  }
  function selectProvider(id: string) {
    const next = MODEL_PRESETS.find((p) => p.id === id);
    setProvider(id);
    setAdvanced(id === 'custom');
    // A credential belongs to a service. Never carry it to a different preset.
    if (next)
      setDraft((p) => ({
        ...p,
        name: next.name,
        baseUrl: next.baseUrl,
        model: next.model,
        apiKey: '',
        rememberKey: false,
      }));
    setStatus(undefined);
    setDeleting(false);
  }
  function field<K extends keyof ModelProfile>(key: K, value: ModelProfile[K]) {
    setDraft((p) => ({
      ...p,
      [key]: value,
      ...(key === 'baseUrl' && value !== p.baseUrl ? { apiKey: '', rememberKey: false } : {}),
    }));
    if (key === 'baseUrl' || key === 'model') setProvider('custom');
    setStatus(undefined);
    setDeleting(false);
  }
  function normalizedProfile() {
    if (!provider) throw new Error('请先选择模型服务。');
    if (preset && !draft.apiKey.trim()) throw new Error('请填写所选服务的 API Key。');
    const profile = {
      ...draft,
      name: draft.name.trim(),
      baseUrl: draft.baseUrl.trim(),
      model: draft.model.trim(),
      apiKey: draft.apiKey.trim(),
    };
    validateProfile(profile);
    return profile;
  }
  function persist(profile: ModelProfile) {
    onSave({
      profiles: existing
        ? state.profiles.map((p) => (p.id === profile.id ? profile : p))
        : [...state.profiles, profile],
      activeId: profile.id,
    });
    setDraft(profile);
  }
  function save() {
    try {
      persist(normalizedProfile());
      setStatus({ text: '模型配置已保存，可返回工作台开始使用。' });
    } catch (error) {
      setStatus({ text: officeError(error), error: true });
    }
  }
  async function connectAndSave() {
    if (controller.current) return;
    const abort = new AbortController();
    controller.current = abort;
    setTesting(true);
    setStatus(undefined);
    try {
      const profile = normalizedProfile();
      await testConnection(profile, abort.signal);
      if (abort.signal.aborted) return;
      persist(profile);
      setStatus({ text: '连接成功，配置已保存。返回工作台即可使用。' });
    } catch (error) {
      setStatus({ text: officeError(error), error: true });
    } finally {
      setTesting(false);
      controller.current = null;
    }
  }
  function remove() {
    try {
      const profiles = state.profiles.filter((p) => p.id !== draft.id);
      onSave({
        profiles,
        activeId: state.activeId === draft.id ? profiles[0]?.id || '' : state.activeId,
      });
      selectProfile(profiles[0] || emptyProfile());
      setDeleting(false);
      setStatus({ text: '模型及保存的密钥已删除。' });
    } catch (error) {
      setStatus({ text: officeError(error), error: true });
    }
  }
  async function importFile(file?: File) {
    if (!file) return;
    try {
      if (file.size > 100_000) throw new Error('配置文件过大。');
      const profiles = importProfiles(await file.text());
      onSave({
        profiles: [...state.profiles, ...profiles],
        activeId: state.activeId || profiles[0].id,
      });
      selectProfile(profiles[0]);
      setStatus({
        text: `已导入 ${profiles.length} 个模型配置，请补充 API Key。`,
      });
    } catch (error) {
      setStatus({ text: officeError(error), error: true });
    }
  }
  function download() {
    const url = URL.createObjectURL(
      new Blob([exportProfiles(state.profiles)], { type: 'application/json' }),
    );
    const anchor = document.createElement('a');
    anchor.href = url;
    anchor.download = 'triobridge-models.json';
    anchor.click();
    setTimeout(() => URL.revokeObjectURL(url), 1000);
    setStatus({ text: '已导出配置，不包含 API Key。' });
  }
  return (
    <section className="settings-page" aria-labelledby="settings-title">
      <button className="text-button" onClick={onBack} disabled={testing}>
        <ArrowLeft size={17} />
        返回工作台
      </button>
      <div className="page-heading">
        <h1 id="settings-title">连接你的模型</h1>
        <p>选择服务，填入 API Key，即可连接。</p>
      </div>
      {state.profiles.length > 0 && (
        <div className="settings-toolbar">
          <label className="grow">
            已保存的模型
            <select
              value={existing ? draft.id : ''}
              disabled={testing}
              onChange={(e) => {
                selectProfile(state.profiles.find((p) => p.id === e.target.value)!);
              }}
            >
              <option value="" disabled>
                新模型配置
              </option>
              {state.profiles.map((p) => (
                <option key={p.id} value={p.id}>
                  {p.name}
                </option>
              ))}
            </select>
          </label>
          <button
            className="secondary"
            disabled={testing}
            onClick={() => {
              selectProfile(emptyProfile());
            }}
          >
            <Plus size={17} />
            新增
          </button>
        </div>
      )}
      <form
        noValidate
        onSubmit={(e) => {
          e.preventDefault();
          void connectAndSave();
        }}
      >
        <fieldset disabled={testing}>
          <label>
            模型服务
            <select value={provider} onChange={(e) => selectProvider(e.target.value)}>
              <option value="" disabled>
                请选择你使用的服务
              </option>
              {MODEL_PRESETS.map((p) => (
                <option key={p.id} value={p.id}>
                  {p.name}
                </option>
              ))}
              <option value="custom">自定义 / 本地模型</option>
            </select>
          </label>
          {provider && (
            <p className="field-help connection-summary">
              请求发送至 <span>{draft.baseUrl}</span>
              {draft.model && <> · 模型 {draft.model}</>}
              {preset && (
                <>
                  {' '}
                  ·{' '}
                  <a href={preset.docsUrl} target="_blank" rel="noreferrer">
                    服务说明
                  </a>
                </>
              )}
            </p>
          )}
          <label>
            <span className="inline-label">
              <KeyRound size={15} aria-hidden="true" />
              API Key {provider === 'custom' && <span className="muted">本地无鉴权服务可留空</span>}
            </span>
            <input
              type="password"
              value={draft.apiKey}
              onChange={(e) => field('apiKey', e.target.value)}
              maxLength={4096}
              placeholder="粘贴所选服务的 API Key"
              autoComplete="new-password"
              spellCheck={false}
            />
          </label>
          <label className="checkbox-label">
            <input
              type="checkbox"
              checked={draft.rememberKey}
              onChange={(e) => field('rememberKey', e.target.checked)}
            />
            在本机记住 API Key
          </label>
          <p className="field-help">默认仅当前会话有效。勾选后会明文保存在此 Office / 浏览器中。</p>
          <button
            type="button"
            className="text-button advanced-toggle"
            aria-expanded={advanced}
            aria-controls="advanced-model-settings"
            onClick={() => setAdvanced((open) => !open)}
          >
            {advanced ? '收起高级设置' : '高级设置：地址、模型与超时'}
          </button>
          {advanced && (
            <div id="advanced-model-settings" className="advanced-settings">
              <label>
                服务名称
                <input
                  required
                  maxLength={100}
                  value={draft.name}
                  onChange={(e) => field('name', e.target.value)}
                  placeholder="例如：我的本地模型"
                  autoComplete="off"
                />
              </label>
              <label>
                API Base URL
                <input
                  required
                  value={draft.baseUrl}
                  onChange={(e) => field('baseUrl', e.target.value)}
                  placeholder="https://api.example.com/v1"
                  spellCheck={false}
                  autoComplete="off"
                  aria-describedby="url-help"
                />
              </label>
              <p className="field-help" id="url-help">
                使用服务商提供的完整地址，不会自动补 /v1。更换地址后需重新填写密钥。
              </p>
              <label>
                模型 ID
                <input
                  required
                  maxLength={200}
                  value={draft.model}
                  onChange={(e) => field('model', e.target.value)}
                  placeholder="填写服务提供的模型名称"
                  spellCheck={false}
                  autoComplete="off"
                />
              </label>
              <label>
                请求超时（秒）
                <input
                  type="number"
                  required
                  min={5}
                  max={300}
                  step={1}
                  value={Number.isNaN(draft.timeoutSeconds) ? '' : draft.timeoutSeconds}
                  onChange={(e) => field('timeoutSeconds', e.target.valueAsNumber)}
                />
              </label>
              <button type="button" className="text-button" onClick={save}>
                仅保存配置
              </button>
            </div>
          )}
        </fieldset>
        {status && (
          <div
            className={`notice ${status.error ? 'error' : 'success'}`}
            role={status.error ? 'alert' : 'status'}
          >
            {status.error ? null : <Check size={17} />}
            {status.text}
          </div>
        )}
        <div className="button-row">
          <button type="submit" className="primary" disabled={testing}>
            {testing ? (
              <LoaderCircle size={17} className="spin" aria-hidden="true" />
            ) : (
              <Wifi size={17} aria-hidden="true" />
            )}
            {testing ? '正在连接…' : '连接并保存'}
          </button>
          {testing && (
            <button
              type="button"
              className="text-button"
              onClick={() => controller.current?.abort()}
            >
              取消测试
            </button>
          )}
        </div>
        <p className="field-help">仅发送一句测试文字，成功后保存。可能产生少量调用费用。</p>
      </form>
      <div className="settings-footer">
        <div className="button-row">
          <button
            className="text-button"
            onClick={download}
            disabled={!state.profiles.length || testing}
          >
            <Download size={16} />
            导出配置
          </button>
          <button
            className="text-button"
            disabled={testing}
            onClick={() => fileInput.current?.click()}
          >
            <Upload size={16} />
            导入配置
          </button>
          <input
            aria-label="导入模型配置文件"
            className="visually-hidden"
            type="file"
            accept="application/json,.json"
            ref={fileInput}
            onChange={(e) => {
              void importFile(e.target.files?.[0]);
              e.target.value = '';
            }}
          />
        </div>
        <p className="field-help">
          Office 各应用可能隔离存储。导出后可在另一应用导入，密钥需重新填写。
        </p>
        {existing && (
          <div className="delete-area">
            {deleting ? (
              <>
                <p>删除「{draft.name}」及其保存的密钥？</p>
                <div className="button-row">
                  <button className="danger" onClick={remove}>
                    确认删除模型
                  </button>
                  <button className="secondary" onClick={() => setDeleting(false)}>
                    取消
                  </button>
                </div>
              </>
            ) : (
              <button
                className="text-button danger-text"
                disabled={testing}
                onClick={() => setDeleting(true)}
              >
                <Trash2 size={16} />
                删除此模型
              </button>
            )}
          </div>
        )}
      </div>
    </section>
  );
}
