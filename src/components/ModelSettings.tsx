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
  function field<K extends keyof ModelProfile>(key: K, value: ModelProfile[K]) {
    setDraft((p) => ({ ...p, [key]: value }));
    setStatus(undefined);
    setDeleting(false);
  }
  function save() {
    try {
      const profile = {
        ...draft,
        name: draft.name.trim(),
        baseUrl: draft.baseUrl.trim(),
        model: draft.model.trim(),
        apiKey: draft.apiKey.trim(),
      };
      validateProfile(profile);
      onSave({
        profiles: existing
          ? state.profiles.map((p) => (p.id === profile.id ? profile : p))
          : [...state.profiles, profile],
        activeId: profile.id,
      });
      setDraft(profile);
      setStatus({ text: '模型配置已保存，可返回工作台开始使用。' });
    } catch (error) {
      setStatus({ text: officeError(error), error: true });
    }
  }
  async function test() {
    const abort = new AbortController();
    controller.current = abort;
    setTesting(true);
    setStatus(undefined);
    try {
      await testConnection(draft, abort.signal);
      setStatus({ text: '连接成功，模型已返回响应。请保存配置后使用。' });
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
      setDraft({ ...(profiles[0] || emptyProfile()) });
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
      setDraft(profiles[0]);
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
        <p>选用你熟悉的服务，三件套使用同一套配置方式。</p>
      </div>
      <div className="settings-toolbar">
        <label className="grow">
          已保存的模型
          <select
            value={existing ? draft.id : ''}
            disabled={testing}
            onChange={(e) => {
              setDraft({
                ...state.profiles.find((p) => p.id === e.target.value)!,
              });
              setStatus(undefined);
              setDeleting(false);
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
            setDraft(emptyProfile());
            setStatus(undefined);
            setDeleting(false);
          }}
        >
          <Plus size={17} />
          新增
        </button>
      </div>
      <form
        onSubmit={(e) => {
          e.preventDefault();
          save();
        }}
      >
        <fieldset disabled={testing}>
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
            OpenAI 兼容接口。填写到 /v1；本机代理可填 /model-proxy/v1。
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
            <span className="inline-label">
              <KeyRound size={15} />
              API Key <span className="muted">本地无鉴权服务可留空</span>
            </span>
            <input
              type="password"
              value={draft.apiKey}
              onChange={(e) => field('apiKey', e.target.value)}
              maxLength={4096}
              placeholder="仅用于向你配置的服务鉴权"
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
          <p className="field-help">
            默认只保留到当前会话结束。勾选后会明文存入此 Office /
            浏览器的本地存储，请仅在可信设备上使用。
          </p>
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
            <Check size={17} />
            保存配置
          </button>
          <button type="button" className="secondary" onClick={test} disabled={testing}>
            {testing ? <LoaderCircle size={17} className="spin" /> : <Wifi size={17} />}
            测试连接
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
        <p className="field-help">连接测试只发送一句测试文字，可能产生少量模型调用费用。</p>
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
