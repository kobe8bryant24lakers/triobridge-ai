import { nativeWorkbench } from '../offline/transport';
import { useEffect, useRef, useState } from 'react';
import {
  ArrowLeft,
  Check,
  Download,
  Eye,
  EyeOff,
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
  type LocalProfileState,
  type ProfileState,
  type SharedProfileState,
  type SharedModelActions,
} from '../core/profile-store';
import { completionUrl, testConnection, validateProfile } from '../core/model-client';
import { MODEL_PRESETS, presetFor, customFormatSuffix } from '../core/model-presets';
import { officeError } from '../office/bridge';
import type {
  ModelApiFormat,
  ModelProfile,
  ProfileDetails,
  SharedModelProfile,
} from '../core/types';

function ApiFormatField({
  profile,
  onChange,
}: {
  profile: ModelProfile;
  onChange: (format: ModelApiFormat) => void;
}) {
  let help = '选择本地推理服务提供的接口格式。';
  let endpoint;
  if (profile.baseUrl.trim()) {
    try {
      endpoint = completionUrl(profile.baseUrl, profile.apiFormat);
    } catch (error) {
      help = (error as Error).message;
    }
  }
  return (
    <>
      <label>
        接口格式
        <select
          aria-label="接口格式"
          value={profile.apiFormat ?? 'openai'}
          onChange={(event) => onChange(event.target.value as ModelApiFormat)}
          aria-describedby="model-endpoint"
        >
          <option value="openai">OpenAI（Chat Completions）</option>
          <option value="anthropic">Claude（Anthropic Messages）</option>
        </select>
      </label>
      <p className="field-help connection-summary" id="model-endpoint">
        {endpoint ? (
          <>
            请求发送至 <span>{endpoint}</span>
          </>
        ) : (
          help
        )}
        {endpoint?.endsWith('/v1/messages/chat/completions') &&
          '。如需直接调用 /v1/messages，请选择 Claude 格式。'}
      </p>
    </>
  );
}

function sharedDraft(profile?: SharedModelProfile): ModelProfile {
  const blank = emptyProfile();
  if (!profile) return blank;
  return {
    id: profile.id,
    name: profile.name,
    baseUrl: profile.baseUrl,
    apiFormat: profile.apiFormat ?? 'openai',
    model: profile.model,
    timeoutSeconds: profile.timeoutSeconds,
    apiKey: '',
    rememberKey: false,
  };
}

interface Props {
  state: ProfileState;
  onSave: (state: LocalProfileState) => void;
  onBack: () => void;
  sharedActions?: SharedModelActions;
}

export function ModelSettings({ state, onSave, onBack, sharedActions }: Props) {
  if (state.mode === 'checking')
    return (
      <section className="settings-page" aria-labelledby="settings-title">
        <button className="text-button" onClick={onBack}>
          <ArrowLeft size={17} />
          返回工作台
        </button>
        <h1 id="settings-title">连接你的模型</h1>
        <p role="status">正在检查共享配置…</p>
      </section>
    );
  if (state.mode === 'shared')
    return <SharedModelSettings state={state} actions={sharedActions} onBack={onBack} />;
  return <LocalModelSettings state={state} onSave={onSave} onBack={onBack} />;
}

function LocalModelSettings({
  state,
  onSave,
  onBack,
}: {
  state: LocalProfileState;
  onSave: (state: LocalProfileState) => void;
  onBack: () => void;
}) {
  const [draft, setDraft] = useState<ModelProfile>(() => ({
    ...(state.profiles.find((p) => p.id === state.activeId) || emptyProfile()),
  }));
  const [status, setStatus] = useState<{ text: string; error?: boolean }>();
  const [provider, setProvider] = useState(() => presetFor(draft));
  const [advanced, setAdvanced] = useState(() => presetFor(draft) === 'custom');
  const [testing, setTesting] = useState(false);
  const [showApiKey, setShowApiKey] = useState(false);
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
    setShowApiKey(false);
    setDraft({ ...profile });
    setProvider(presetFor(profile));
    setAdvanced(presetFor(profile) === 'custom');
    setStatus(undefined);
    setDeleting(false);
  }
  function selectProvider(id: string) {
    setShowApiKey(false);
    const next = MODEL_PRESETS.find((p) => p.id === id);
    setProvider(id);
    setAdvanced(id === 'custom');
    // A credential belongs to a service. Never carry it to a different preset.
    if (next)
      setDraft((p) => ({
        ...p,
        name: next.name,
        baseUrl: next.baseUrl,
        apiFormat: 'openai',
        model: next.model,
        apiKey: '',
        rememberKey: false,
      }));
    setStatus(undefined);
    setDeleting(false);
  }
  function field<K extends keyof ModelProfile>(key: K, value: ModelProfile[K]) {
    const destination = key === 'baseUrl' || key === 'apiFormat';
    if (destination || key === 'model') setShowApiKey(false);
    setDraft((p) => ({
      ...p,
      [key]: value,
      ...(destination && value !== p[key] ? { apiKey: '', rememberKey: false } : {}),
    }));
    if (destination || key === 'model') setProvider('custom');
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
      <p className="model-scope-note">仅此应用 · 配置不会自动同步到其他 Office 应用。</p>
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
                  {customFormatSuffix(p)}
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
          {provider === 'custom' ? (
            <ApiFormatField profile={draft} onChange={(value) => field('apiFormat', value)} />
          ) : (
            provider && (
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
            )
          )}
          <div className="api-key-field">
            <label className="inline-label" htmlFor="model-api-key">
              <KeyRound size={15} aria-hidden="true" />
              API Key {provider === 'custom' && <span className="muted">本地无鉴权服务可留空</span>}
            </label>
            <div className="api-key-controls">
              <input
                id="model-api-key"
                type={showApiKey ? 'text' : 'password'}
                value={draft.apiKey}
                onChange={(e) => field('apiKey', e.target.value)}
                maxLength={4096}
                placeholder="粘贴所选服务的 API Key"
                autoComplete="new-password"
                spellCheck={false}
              />
              <button
                type="button"
                className="text-button"
                aria-label={showApiKey ? '隐藏密钥' : '显示密钥'}
                aria-pressed={showApiKey}
                aria-controls="model-api-key"
                onClick={() => setShowApiKey((visible) => !visible)}
              >
                {showApiKey ? (
                  <EyeOff size={17} aria-hidden="true" />
                ) : (
                  <Eye size={17} aria-hidden="true" />
                )}
                {showApiKey ? '隐藏' : '显示'}
              </button>
            </div>
          </div>
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
                  aria-describedby={provider === 'custom' ? 'url-help model-endpoint' : 'url-help'}
                />
              </label>
              <p className="field-help" id="url-help">
                可填写基础地址或完整接口地址。更换地址或接口格式后需重新填写密钥。
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

function sharedProvider(profile: ModelProfile) {
  if (nativeWorkbench()) {
    try {
      if (['localhost', '127.0.0.1', '[::1]'].includes(new URL(profile.baseUrl).hostname))
        return 'device';
    } catch {
      /* A new profile starts without an address. */
    }
  }
  return presetFor(profile);
}

function SharedModelSettings({
  state,
  actions,
  onBack,
}: {
  state: SharedProfileState;
  actions?: SharedModelActions;
  onBack: () => void;
}) {
  const initial =
    state.profiles.find((profile) => profile.id === state.activeId) ?? state.profiles[0];
  const [draft, setDraft] = useState(() => sharedDraft(initial));
  const [baseline, setBaseline] = useState(() => ({
    revision: state.revision,
    baseUrl: initial?.baseUrl ?? '',
    apiFormat: initial?.apiFormat ?? 'openai',
    hasKey: initial?.hasKey ?? false,
  }));
  const [provider, setProvider] = useState(() => sharedProvider(sharedDraft(initial)));
  const [advanced, setAdvanced] = useState(() =>
    ['custom', 'device'].includes(sharedProvider(sharedDraft(initial))),
  );
  const [credential, setCredential] = useState<'keep' | 'replace' | 'remove'>(
    initial?.hasKey
      ? 'keep'
      : sharedProvider(sharedDraft(initial)) === 'device'
        ? 'remove'
        : 'replace',
  );
  const [activate, setActivate] = useState(!initial || initial.id === state.activeId);
  const [dirty, setDirty] = useState(false);
  const [showApiKey, setShowApiKey] = useState(false);
  const [busy, setBusy] = useState<
    'save' | 'connect' | 'delete' | 'import' | 'migrate' | 'legacy' | 'refresh' | null
  >(null);
  const [legacyProfiles, setLegacyProfiles] = useState<
    (ProfileDetails & { hasKey: boolean })[] | null
  >(null);
  const [selectedLegacyIds, setSelectedLegacyIds] = useState<string[]>([]);
  const [status, setStatus] = useState<{ text: string; error?: boolean }>();
  const [conflict, setConflict] = useState(false);
  const [deleting, setDeleting] = useState(false);
  const controller = useRef<AbortController | null>(null);
  const fileInput = useRef<HTMLInputElement>(null);
  const offline = state.connection !== 'ready' || !actions;
  const existing = state.profiles.some((profile) => profile.id === draft.id);
  const changedElsewhere = state.revision > baseline.revision;
  const blocked = offline || !!busy;

  useEffect(() => () => controller.current?.abort(), []);

  function loadDraft(profile: SharedModelProfile | undefined, next: SharedProfileState) {
    const value = sharedDraft(profile);
    setDraft(value);
    setBaseline({
      revision: next.revision,
      baseUrl: profile?.baseUrl ?? '',
      apiFormat: profile?.apiFormat ?? 'openai',
      hasKey: profile?.hasKey ?? false,
    });
    setProvider(sharedProvider(value));
    setAdvanced(['custom', 'device'].includes(sharedProvider(value)));
    setCredential(
      profile?.hasKey ? 'keep' : sharedProvider(value) === 'device' ? 'remove' : 'replace',
    );
    setActivate(!profile || profile.id === next.activeId);
    setDirty(false);
    setShowApiKey(false);
    setConflict(false);
    setDeleting(false);
    setStatus(undefined);
  }

  // A poll may update the catalog, but never replace an unsaved draft or its revision.
  useEffect(() => {
    if (state.revision <= baseline.revision || dirty || busy) return;
    loadDraft(
      state.profiles.find((profile) => profile.id === draft.id) ??
        state.profiles.find((profile) => profile.id === state.activeId),
      state,
    );
  }, [state.revision, baseline.revision, dirty, busy]);

  function field<K extends keyof ModelProfile>(key: K, value: ModelProfile[K]) {
    const destination = key === 'baseUrl' || key === 'apiFormat';
    setDraft((current) => ({
      ...current,
      [key]: value,
      ...(destination && value !== current[key] ? { apiKey: '' } : {}),
    }));
    if (destination) setCredential((current) => (current === 'remove' ? 'remove' : 'replace'));
    if (destination || key === 'model') {
      if (provider !== 'device') setProvider('custom');
      setShowApiKey(false);
    }
    setDirty(true);
    setDeleting(false);
    setStatus(undefined);
  }

  function selectProvider(value: string) {
    const preset = MODEL_PRESETS.find((item) => item.id === value);
    setProvider(value);
    setAdvanced(value === 'custom' || value === 'device');
    setShowApiKey(false);
    setCredential('replace');
    setDirty(true);
    setStatus(undefined);
    setDeleting(false);
    if (value === 'device') {
      setDraft((current) => ({
        ...current,
        name: '本机模型',
        baseUrl: 'http://127.0.0.1:8001/v1',
        apiFormat: 'openai',
        model: '',
        apiKey: '',
      }));
      setCredential('remove');
    }
    if (preset)
      setDraft((current) => ({
        ...current,
        name: preset.name,
        baseUrl: preset.baseUrl,
        apiFormat: 'openai',
        model: preset.model,
        apiKey: '',
      }));
  }

  async function run(
    kind: NonNullable<typeof busy>,
    operation: (signal: AbortSignal) => Promise<SharedProfileState>,
    success: (next: SharedProfileState) => void,
  ) {
    if (controller.current || !actions || (offline && kind !== 'refresh')) return;
    const abort = new AbortController();
    controller.current = abort;
    setBusy(kind);
    setStatus(undefined);
    setShowApiKey(false);
    try {
      const next = await operation(abort.signal);
      if (!abort.signal.aborted) success(next);
    } catch (error) {
      if ((error as { status?: number })?.status === 409) {
        setConflict(true);
        setStatus({
          text: '其他应用已更新配置。你的修改仍在，请重新载入最新配置后核对。',
          error: true,
        });
      } else {
        setStatus({ text: officeError(error), error: true });
      }
    } finally {
      controller.current = null;
      setBusy(null);
    }
  }

  function save(test: boolean) {
    if (!actions || blocked) return;
    try {
      if (!provider) throw new Error('请先选择模型服务。');
      const profile = {
        id: draft.id,
        name: draft.name.trim(),
        baseUrl: draft.baseUrl.trim(),
        apiFormat: draft.apiFormat ?? 'openai',
        model: draft.model.trim(),
        timeoutSeconds: draft.timeoutSeconds,
      };
      validateProfile({ ...profile, apiKey: draft.apiKey.trim(), rememberKey: false });
      if (
        credential === 'keep' &&
        (profile.baseUrl !== baseline.baseUrl || profile.apiFormat !== baseline.apiFormat)
      )
        throw new Error('服务地址或接口格式已更改，请填写新密钥或选择不使用密钥。');
      if (credential === 'replace' && !draft.apiKey.trim())
        throw new Error('请填写新密钥；无需密钥的服务请选择「不使用密钥」。');
      const key =
        credential === 'replace'
          ? { action: 'replace' as const, value: draft.apiKey.trim() }
          : { action: credential };
      void run(
        test ? 'connect' : 'save',
        (signal) =>
          actions.upsert(
            {
              expectedRevision: baseline.revision,
              profile,
              credential: key,
              activate,
              testConnection: test,
            },
            signal,
          ),
        (next) => {
          loadDraft(
            next.profiles.find((item) => item.id === profile.id),
            next,
          );
          setStatus({ text: test ? '连接成功，配置已在三个应用间共享。' : '共享配置已保存。' });
        },
      );
    } catch (error) {
      setStatus({ text: officeError(error), error: true });
    }
  }

  function refreshDraft() {
    if (!actions) return;
    void run(
      'refresh',
      (signal) => actions.refresh(signal),
      (next) => {
        loadDraft(
          next.profiles.find((profile) => profile.id === draft.id) ??
            next.profiles.find((profile) => profile.id === next.activeId),
          next,
        );
        setStatus({ text: '已载入最新共享配置。' });
      },
    );
  }

  function reconnect() {
    if (!actions) return;
    void run(
      'refresh',
      (signal) => actions.refresh(signal),
      (next) => {
        if (!dirty)
          loadDraft(
            next.profiles.find((profile) => profile.id === draft.id),
            next,
          );
        setStatus({ text: '已重新读取共享配置。草稿仍可编辑。' });
      },
    );
  }

  async function chooseLegacyProfiles() {
    if (!actions || blocked || controller.current) return;
    setBusy('legacy');
    setStatus(undefined);
    try {
      const profiles = await actions.listLegacyProfiles();
      setLegacyProfiles(profiles);
      setSelectedLegacyIds([]);
    } catch (error) {
      setStatus({ text: officeError(error), error: true });
    } finally {
      setBusy(null);
    }
  }

  async function importFile(file?: File) {
    if (!file || !actions || blocked) return;
    const revision = state.revision;
    try {
      if (file.size > 100_000) throw new Error('配置文件过大。');
      const profiles = importProfiles(await file.text());
      const items = profiles.map(({ id, name, baseUrl, apiFormat, model, timeoutSeconds }) => ({
        profile: { id, name, baseUrl, apiFormat, model, timeoutSeconds },
        credential: { action: 'remove' as const },
      }));
      void run(
        'import',
        (signal) => actions.importProfiles({ expectedRevision: revision, profiles: items }, signal),
        (next) => {
          loadDraft(
            next.profiles.find((profile) => profile.id === items[0].profile.id),
            next,
          );
          setStatus({ text: `已导入 ${items.length} 个共享配置，请分别补充密钥。` });
        },
      );
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
      <button className="text-button" onClick={onBack} disabled={!!busy}>
        <ArrowLeft size={17} />
        返回工作台
      </button>
      <div className="page-heading">
        <h1 id="settings-title">连接你的模型</h1>
        <p>配置一次，在 Word、Excel 和 PowerPoint 中使用。</p>
      </div>
      <p className="model-scope-note">
        <strong>本机共享</strong> · 切换当前模型后，其他 Office 应用也会同步切换。
      </p>
      {offline && (
        <div className="notice error" role="alert">
          暂时无法读取共享配置。草稿已保留，可重新连接后保存。
          <button
            type="button"
            className="text-button"
            disabled={!!busy || !actions}
            onClick={reconnect}
          >
            重新连接
          </button>
        </div>
      )}
      {(conflict || (changedElsewhere && dirty)) && (
        <div className="notice warning shared-conflict" role="alert">
          <p>其他应用已更新配置。当前草稿已保留，重新载入会放弃未保存的修改。</p>
          <button type="button" className="text-button" disabled={blocked} onClick={refreshDraft}>
            放弃修改并载入最新配置
          </button>
        </div>
      )}
      {state.legacyAvailable && (
        <div className="shared-migration">
          <h2>此应用还有旧配置</h2>
          <p className="field-help">
            选择需要共享的模型。所选配置全部迁移成功后才移除旧副本；未选配置继续留在此应用。
          </p>
          {legacyProfiles ? (
            <>
              <fieldset className="legacy-model-options" disabled={blocked}>
                <legend>选择要迁移的模型</legend>
                {legacyProfiles.map((profile) => (
                  <label className="checkbox-label legacy-model-option" key={profile.id}>
                    <input
                      type="checkbox"
                      checked={selectedLegacyIds.includes(profile.id)}
                      onChange={(event) =>
                        setSelectedLegacyIds((ids) =>
                          event.target.checked
                            ? [...ids, profile.id]
                            : ids.filter((id) => id !== profile.id),
                        )
                      }
                    />
                    <span>
                      <strong>{profile.name}</strong>
                      <span>
                        {profile.model} · {profile.hasKey ? '已保存密钥' : '未保存密钥'}
                      </span>
                      <span>{profile.baseUrl}</span>
                    </span>
                  </label>
                ))}
                {!legacyProfiles.length && <p className="field-help">没有可迁移的配置。</p>}
              </fieldset>
              <div className="button-row">
                <button
                  type="button"
                  className="secondary"
                  disabled={blocked || !selectedLegacyIds.length}
                  onClick={() => {
                    if (!actions) return;
                    void run(
                      'migrate',
                      (signal) => actions.migrateLocal(state.revision, selectedLegacyIds, signal),
                      (next) => {
                        loadDraft(
                          next.profiles.find((profile) => profile.id === next.activeId),
                          next,
                        );
                        setLegacyProfiles(null);
                        setSelectedLegacyIds([]);
                        setStatus({
                          text: next.warning ?? '所选配置已迁移，现在可在三个应用中使用。',
                        });
                      },
                    );
                  }}
                >
                  {busy === 'migrate'
                    ? '正在迁移…'
                    : `迁移选中的 ${selectedLegacyIds.length} 个模型`}
                </button>
                <button
                  type="button"
                  className="text-button"
                  disabled={blocked}
                  onClick={() => {
                    setLegacyProfiles(null);
                    setSelectedLegacyIds([]);
                  }}
                >
                  取消迁移
                </button>
              </div>
            </>
          ) : (
            <button
              type="button"
              className="secondary"
              disabled={blocked}
              onClick={() => {
                void chooseLegacyProfiles();
              }}
            >
              {busy === 'legacy' ? '正在读取…' : '迁移本应用配置'}
            </button>
          )}
        </div>
      )}
      <div className="settings-toolbar">
        <label className="grow">
          本机共享模型
          <select
            aria-label="本机共享模型"
            value={existing ? draft.id : ''}
            disabled={blocked}
            onChange={(event) =>
              loadDraft(
                state.profiles.find((profile) => profile.id === event.target.value),
                state,
              )
            }
          >
            <option value="" disabled>
              新模型配置
            </option>
            {state.profiles.map((profile) => (
              <option key={profile.id} value={profile.id}>
                {profile.name}
                {customFormatSuffix(profile)}
                {profile.id === state.activeId ? ' · 当前使用' : ''}
              </option>
            ))}
          </select>
        </label>
        <button
          type="button"
          className="secondary"
          disabled={blocked}
          onClick={() => loadDraft(undefined, state)}
        >
          <Plus size={17} />
          新增
        </button>
      </div>
      <form
        noValidate
        onSubmit={(event) => {
          event.preventDefault();
          save(true);
        }}
      >
        <fieldset disabled={blocked}>
          <label>
            模型服务
            <select
              aria-label="模型服务"
              value={provider}
              onChange={(event) => selectProvider(event.target.value)}
            >
              <option value="" disabled>
                请选择你使用的服务
              </option>
              {MODEL_PRESETS.map((preset) => (
                <option key={preset.id} value={preset.id}>
                  {preset.name}
                </option>
              ))}
              {nativeWorkbench() && <option value="device">本机 API（127.0.0.1）</option>}
              <option value="custom">自定义 / 本地模型</option>
            </select>
          </label>
          {provider === 'custom' || provider === 'device' ? (
            <ApiFormatField profile={draft} onChange={(value) => field('apiFormat', value)} />
          ) : (
            provider && (
              <p className="field-help connection-summary">
                请求发送至 <span>{draft.baseUrl}</span>
                {draft.model && <> · 模型 {draft.model}</>}
              </p>
            )
          )}
          <div className="shared-key-status">
            <KeyRound size={16} aria-hidden="true" />
            <span>{baseline.hasKey ? '已保存密钥' : '尚未保存密钥'}</span>
          </div>
          <label>
            密钥操作
            <select
              aria-label="密钥操作"
              value={credential}
              onChange={(event) => {
                setCredential(event.target.value as typeof credential);
                setDraft((value) => ({ ...value, apiKey: '' }));
                setShowApiKey(false);
                setDirty(true);
                setStatus(undefined);
              }}
            >
              {baseline.hasKey &&
                draft.baseUrl === baseline.baseUrl &&
                draft.apiFormat === baseline.apiFormat && (
                  <option value="keep">保留已保存的密钥</option>
                )}
              <option value="replace">{baseline.hasKey ? '替换密钥' : '填写密钥'}</option>
              <option value="remove">{baseline.hasKey ? '删除已保存的密钥' : '不使用密钥'}</option>
            </select>
          </label>
          {credential === 'replace' && (
            <div className="api-key-field">
              <label htmlFor="shared-api-key">{baseline.hasKey ? '新 API Key' : 'API Key'}</label>
              <div className="api-key-controls">
                <input
                  id="shared-api-key"
                  type={showApiKey ? 'text' : 'password'}
                  value={draft.apiKey}
                  onChange={(event) => field('apiKey', event.target.value)}
                  maxLength={4096}
                  placeholder="粘贴所选服务的 API Key"
                  autoComplete="new-password"
                  spellCheck={false}
                />
                <button
                  type="button"
                  className="text-button"
                  aria-label={showApiKey ? '隐藏密钥' : '显示密钥'}
                  aria-pressed={showApiKey}
                  aria-controls="shared-api-key"
                  onClick={() => setShowApiKey((value) => !value)}
                >
                  {showApiKey ? (
                    <EyeOff size={17} aria-hidden="true" />
                  ) : (
                    <Eye size={17} aria-hidden="true" />
                  )}
                  {showApiKey ? '隐藏' : '显示'}
                </button>
              </div>
            </div>
          )}
          <p className="field-help">
            {credential === 'remove' && baseline.hasKey
              ? '保存后会删除此模型的共享密钥，三个应用都会受到影响。'
              : '已保存的密钥不会显示或返回到 Office 页面。'}
          </p>
          {state.storage === 'macos-keychain' && credential !== 'remove' && (
            <p className="field-help">
              首次使用或升级后，macOS 可能请求允许三桥 AI
              访问钥匙串。请在系统窗口确认，无需在插件中重新填写已保存的密钥。
            </p>
          )}
          <button
            type="button"
            className="text-button advanced-toggle"
            aria-expanded={advanced}
            aria-controls="shared-advanced-settings"
            onClick={() => setAdvanced((value) => !value)}
          >
            {advanced ? '收起高级设置' : '高级设置：地址、模型与超时'}
          </button>
          {advanced && (
            <div className="advanced-settings" id="shared-advanced-settings">
              <label>
                服务名称
                <input
                  required
                  maxLength={100}
                  value={draft.name}
                  onChange={(event) => field('name', event.target.value)}
                  placeholder="例如：我的模型服务"
                  autoComplete="off"
                />
              </label>
              <label>
                API Base URL
                <input
                  required
                  aria-label="API Base URL"
                  value={draft.baseUrl}
                  onChange={(event) => field('baseUrl', event.target.value)}
                  placeholder="https://api.example.com/v1"
                  autoComplete="off"
                  spellCheck={false}
                  aria-describedby={
                    provider === 'custom' || provider === 'device'
                      ? 'shared-url-help model-endpoint'
                      : 'shared-url-help'
                  }
                />
              </label>
              <p id="shared-url-help" className="field-help">
                可填写基础地址或完整接口地址。更换地址或接口格式后，请重新填写密钥或选择“不使用密钥”。
              </p>
              <label>
                模型 ID
                <input
                  required
                  maxLength={200}
                  aria-label="模型 ID"
                  value={draft.model}
                  onChange={(event) => field('model', event.target.value)}
                  placeholder="填写服务提供的模型名称"
                  autoComplete="off"
                  spellCheck={false}
                />
              </label>
              <label>
                请求超时（秒）
                <input
                  type="number"
                  min={5}
                  max={300}
                  step={1}
                  value={Number.isNaN(draft.timeoutSeconds) ? '' : draft.timeoutSeconds}
                  onChange={(event) => field('timeoutSeconds', event.target.valueAsNumber)}
                />
              </label>
            </div>
          )}
          <label className="checkbox-label">
            <input
              type="checkbox"
              checked={activate}
              onChange={(event) => {
                setActivate(event.target.checked);
                setDirty(true);
              }}
            />
            保存后在三个应用中使用此模型
          </label>
        </fieldset>
        {status && (
          <div
            className={`notice ${status.error ? 'error' : 'success'}`}
            role={status.error ? 'alert' : 'status'}
          >
            {!status.error && <Check size={17} />}
            {status.text}
          </div>
        )}
        <div className="button-row">
          <button type="submit" className="primary" disabled={blocked}>
            {busy === 'connect' ? <LoaderCircle size={17} className="spin" /> : <Wifi size={17} />}
            {busy === 'connect' ? '正在连接…' : '连接并保存'}
          </button>
          <button
            type="button"
            className="secondary"
            disabled={blocked}
            onClick={() => save(false)}
          >
            {busy === 'save' ? '正在保存…' : '仅保存配置'}
          </button>
          {busy === 'connect' && (
            <button
              type="button"
              className="text-button"
              onClick={() => controller.current?.abort()}
            >
              取消测试
            </button>
          )}
        </div>
        <p className="field-help">连接测试仅发送一句测试文字，成功后保存。可能产生少量调用费用。</p>
      </form>
      <div className="settings-footer">
        <div className="button-row">
          <button
            type="button"
            className="text-button"
            disabled={!!busy || !state.profiles.length}
            onClick={download}
          >
            <Download size={16} />
            导出配置
          </button>
          <button
            type="button"
            className="text-button"
            disabled={blocked}
            onClick={() => fileInput.current?.click()}
          >
            <Upload size={16} />
            导入配置
          </button>
          <input
            className="visually-hidden"
            aria-label="导入模型配置文件"
            type="file"
            accept="application/json,.json"
            ref={fileInput}
            onChange={(event) => {
              void importFile(event.target.files?.[0]);
              event.target.value = '';
            }}
          />
        </div>
        <p className="field-help">导出文件只含服务地址等配置，不含密钥。</p>
        {existing && (
          <div className="delete-area">
            {deleting ? (
              <>
                <p>从三个应用中删除「{draft.name}」及其保存的密钥？</p>
                <div className="button-row">
                  <button
                    type="button"
                    className="danger"
                    disabled={blocked}
                    onClick={() => {
                      if (!actions) return;
                      void run(
                        'delete',
                        (signal) =>
                          actions.remove(
                            { expectedRevision: baseline.revision, id: draft.id },
                            signal,
                          ),
                        (next) => {
                          loadDraft(
                            next.profiles.find((profile) => profile.id === next.activeId),
                            next,
                          );
                          setStatus({ text: '模型及保存的密钥已从共享列表删除。' });
                        },
                      );
                    }}
                  >
                    确认删除模型
                  </button>
                  <button
                    type="button"
                    className="secondary"
                    disabled={!!busy}
                    onClick={() => setDeleting(false)}
                  >
                    取消
                  </button>
                </div>
              </>
            ) : (
              <button
                type="button"
                className="text-button danger-text"
                disabled={blocked}
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
