import type {
  ModelProfile,
  ProfileDetails,
  RuntimeModelProfile,
  SharedModelProfile,
} from './types';
import { validateProfile } from './model-client';
import { getLocalCapability } from './local-capability';
import {
  createSharedModelClient,
  SharedModelClientError,
  SHARED_STATE_EVENT,
  parseSharedState,
  type SharedImportRequest,
  type SharedStateSnapshot,
  type SharedUpsertRequest,
} from './shared-model-client';

export const PROFILE_KEY = 'triobridge.profiles.v1';
export const SESSION_KEY = 'triobridge.keys.v1';
export interface LocalProfileState {
  mode?: 'local';
  profiles: ModelProfile[];
  activeId: string;
  warning?: string;
}
export interface SharedProfileState {
  mode: 'shared';
  profiles: SharedModelProfile[];
  activeId: string;
  revision: number;
  storage?: SharedStateSnapshot['storage'];
  connection: 'ready' | 'offline';
  legacyAvailable: boolean;
  warning?: string;
}
export interface CheckingProfileState {
  mode: 'checking';
  profiles: [];
  activeId: '';
  warning?: string;
}
export type ProfileState = LocalProfileState | SharedProfileState | CheckingProfileState;
export interface SharedModelActions {
  upsert: (input: SharedUpsertRequest, signal?: AbortSignal) => Promise<SharedProfileState>;
  remove: (
    input: { expectedRevision: number; id: string },
    signal?: AbortSignal,
  ) => Promise<SharedProfileState>;
  importProfiles: (input: SharedImportRequest, signal?: AbortSignal) => Promise<SharedProfileState>;
  listLegacyProfiles: () => Promise<(ProfileDetails & { hasKey: boolean })[]>;
  migrateLocal: (
    expectedRevision: number,
    ids: string[],
    signal?: AbortSignal,
  ) => Promise<SharedProfileState>;
  refresh: (signal?: AbortSignal) => Promise<SharedProfileState>;
}
export const emptyProfile = (): ModelProfile => ({
  id: crypto.randomUUID(),
  name: '',
  baseUrl: 'https://api.openai.com/v1',
  model: '',
  apiKey: '',
  timeoutSeconds: 60,
  rememberKey: false,
});

function parseProfile(raw: unknown): ModelProfile {
  if (!raw || typeof raw !== 'object') throw new Error('模型配置格式无效。');
  const p = raw as Partial<ModelProfile>;
  if (
    typeof p.id !== 'string' ||
    !p.id ||
    typeof p.name !== 'string' ||
    typeof p.baseUrl !== 'string' ||
    typeof p.model !== 'string' ||
    typeof p.timeoutSeconds !== 'number'
  )
    throw new Error('模型配置缺少必要字段。');
  const profile = {
    id: p.id,
    name: p.name,
    baseUrl: p.baseUrl,
    model: p.model,
    timeoutSeconds: p.timeoutSeconds,
    apiKey: typeof p.apiKey === 'string' ? p.apiKey : '',
    rememberKey: p.rememberKey === true,
  };
  validateProfile(profile);
  return profile;
}

export function loadProfiles(local?: Storage, session?: Storage): LocalProfileState {
  try {
    local ??= localStorage;
    session ??= sessionStorage;
    const saved = local.getItem(PROFILE_KEY);
    if (!saved) return { profiles: [], activeId: '' };
    const data = JSON.parse(saved);
    if (data.version !== 1 || !Array.isArray(data.profiles) || data.profiles.length > 30)
      throw new Error('Invalid profile store');
    const keys = JSON.parse(session.getItem(SESSION_KEY) || '{}') as Record<string, unknown>;
    const profiles = data.profiles.map(parseProfile).map((p: ModelProfile) => ({
      ...p,
      apiKey: p.rememberKey
        ? p.apiKey
        : typeof keys?.[p.id] === 'string'
          ? (keys[p.id] as string)
          : '',
    }));
    if (new Set(profiles.map((p: ModelProfile) => p.id)).size !== profiles.length)
      throw new Error('Duplicate profile IDs');
    return {
      profiles,
      activeId: profiles.some((p: ModelProfile) => p.id === data.activeId)
        ? data.activeId
        : profiles[0]?.id || '',
    };
  } catch {
    return {
      profiles: [],
      activeId: '',
      warning: '本地配置无法读取。原始数据未被覆盖，可在模型设置中重新保存配置。',
    };
  }
}

export function saveProfiles(state: LocalProfileState, local?: Storage, session?: Storage): void {
  state.profiles.forEach(validateProfile);
  if (state.profiles.length > 30) throw new Error('最多保存 30 个模型配置。');
  const sessionKeys = Object.fromEntries(
    state.profiles.filter((p) => !p.rememberKey && p.apiKey).map((p) => [p.id, p.apiKey]),
  );
  try {
    local ??= localStorage;
    session ??= sessionStorage;
    // Save non-persistent secrets first; do not claim success if storage is unavailable.
    session.setItem(SESSION_KEY, JSON.stringify(sessionKeys));
    local.setItem(
      PROFILE_KEY,
      JSON.stringify({
        version: 1,
        activeId: state.activeId,
        profiles: state.profiles.map((p) => ({
          ...p,
          apiKey: p.rememberKey ? p.apiKey : '',
        })),
      }),
    );
  } catch {
    throw new Error('无法保存配置，浏览器存储可能被禁用或已满。请允许本地存储后重试。');
  }
}

export function exportProfiles(profiles: RuntimeModelProfile[]): string {
  return JSON.stringify(
    {
      version: 1,
      profiles: profiles.map((p) => ({
        id: p.id,
        name: p.name,
        baseUrl: p.baseUrl,
        model: p.model,
        timeoutSeconds: p.timeoutSeconds,
        apiKey: '',
        rememberKey: false,
      })),
    },
    null,
    2,
  );
}

export function hasLegacyProfiles(local: Storage = localStorage): boolean {
  try {
    // Inspect key names only. Reading legacy credentials requires explicit migration.
    for (let i = 0; i < local.length; i += 1) if (local.key(i) === PROFILE_KEY) return true;
  } catch {
    // A blocked local store cannot be migrated automatically.
  }
  return false;
}

interface ControllerOptions {
  fetcher?: typeof fetch;
  capability?: () => string | undefined;
  local?: Storage;
  session?: Storage;
  pollMilliseconds?: number;
}

function availableStorage(name: 'localStorage' | 'sessionStorage'): Storage {
  try {
    return globalThis[name];
  } catch {
    const blocked = (): never => {
      throw new Error('浏览器存储不可用。');
    };
    // Shared metadata remains usable even when the WebView blocks browser stores.
    // Local saves still fail honestly; there is no pretend-persistent fallback.
    return {
      get length() {
        return blocked();
      },
      clear: blocked,
      getItem: blocked,
      key: blocked,
      removeItem: blocked,
      setItem: blocked,
    };
  }
}

export function createProfileController(options: ControllerOptions = {}) {
  const local = options.local ?? availableStorage('localStorage');
  const session = options.session ?? availableStorage('sessionStorage');
  const capability = options.capability ?? getLocalCapability;
  const client = createSharedModelClient(options.fetcher ?? fetch, capability);
  let state: ProfileState = capability()
    ? { mode: 'checking', profiles: [], activeId: '' }
    : loadProfiles(local, session);
  const subscribers = new Set<() => void>();
  let lifetime: AbortController | undefined;
  let interval: ReturnType<typeof setInterval> | undefined;
  let refreshing: Promise<SharedProfileState> | undefined;
  let mutating = false;
  let localWarning: string | undefined;
  const publish = (next: ProfileState) => {
    state = next;
    subscribers.forEach((listener) => listener());
  };
  function apply(snapshot: SharedStateSnapshot): SharedProfileState {
    if (state.mode === 'shared' && snapshot.revision < state.revision) return state;
    const next: SharedProfileState = {
      mode: 'shared',
      profiles: snapshot.profiles,
      activeId: snapshot.activeId,
      revision: snapshot.revision,
      storage: snapshot.storage,
      connection: 'ready',
      legacyAvailable: hasLegacyProfiles(local),
      warning: [snapshot.warning, localWarning].filter(Boolean).join(' ') || undefined,
    };
    publish(next);
    return next;
  }
  function offline(error: unknown) {
    if (error instanceof DOMException && error.name === 'AbortError') return;
    const previous = state.mode === 'shared' ? state : undefined;
    publish({
      mode: 'shared',
      profiles: previous?.profiles ?? [],
      activeId: previous?.activeId ?? '',
      revision: previous?.revision ?? 0,
      storage: previous?.storage,
      connection: 'offline',
      legacyAvailable: hasLegacyProfiles(local),
      warning: '本机共享服务暂时未连接。已保留共享配置，请恢复本机服务后重试；未切换到旧本地密钥。',
    });
  }
  async function refresh(signal?: AbortSignal): Promise<SharedProfileState> {
    try {
      return apply(await client.state(signal ?? lifetime?.signal));
    } catch (error) {
      offline(error);
      throw error;
    }
  }
  function backgroundRefresh() {
    if (!capability() || refreshing) return;
    const pending = refresh().finally(() => {
      if (refreshing === pending) refreshing = undefined;
    });
    refreshing = pending;
    void pending.catch(() => {});
  }
  async function mutate(
    operation: () => Promise<SharedStateSnapshot>,
  ): Promise<SharedProfileState> {
    if (mutating) throw new Error('正在保存模型配置，请等待当前操作完成。');
    if (state.mode !== 'shared' || state.connection !== 'ready')
      throw new Error('请先恢复本机共享服务连接后再修改配置。');
    mutating = true;
    try {
      return apply(await operation());
    } catch (error) {
      if (error instanceof SharedModelClientError && error.state) apply(error.state);
      else if (
        !(error instanceof SharedModelClientError) ||
        ['storage', 'unavailable', 'local_auth', 'local_timeout'].includes(error.code)
      )
        offline(error);
      throw error;
    } finally {
      mutating = false;
    }
  }
  const actions: SharedModelActions = {
    refresh,
    upsert: (input, signal) => mutate(() => client.upsert(input, signal ?? lifetime?.signal)),
    remove: (input, signal) => mutate(() => client.remove(input, signal ?? lifetime?.signal)),
    importProfiles: (input, signal) =>
      mutate(() => client.importProfiles(input, signal ?? lifetime?.signal)),
    async listLegacyProfiles() {
      const source = loadProfiles(local, session);
      if (source.warning) throw new Error(source.warning);
      return source.profiles.map(({ id, name, baseUrl, model, timeoutSeconds, apiKey }) => ({
        id,
        name,
        baseUrl,
        model,
        timeoutSeconds,
        hasKey: Boolean(apiKey),
      }));
    },
    async migrateLocal(expectedRevision, ids, signal) {
      // This is the only shared-mode operation that reads old browser credentials.
      const originalLocal = local.getItem(PROFILE_KEY);
      const originalSession = session.getItem(SESSION_KEY);
      const source = loadProfiles(local, session);
      if (source.warning) throw new Error(source.warning);
      const selected = source.profiles.filter((p) => ids.includes(p.id));
      if (!ids.length || new Set(ids).size !== ids.length || selected.length !== ids.length)
        throw new Error('请选择仍存在的旧模型配置后再迁移。');
      const sharedActive = state.mode === 'shared' ? state.activeId : '';
      const migrated = await actions.importProfiles(
        {
          expectedRevision,
          profiles: selected.map(({ id, name, baseUrl, model, timeoutSeconds, apiKey }) => ({
            profile: { id, name, baseUrl, model, timeoutSeconds },
            credential: apiKey
              ? { action: 'replace' as const, value: apiKey }
              : { action: 'remove' as const },
          })),
          activeId: selected.some((p) => p.id === source.activeId)
            ? source.activeId
            : sharedActive || selected[0].id,
        },
        signal,
      );
      try {
        if (
          local.getItem(PROFILE_KEY) !== originalLocal ||
          session.getItem(SESSION_KEY) !== originalSession
        )
          throw new Error('旧配置在迁移期间发生变化，已保留本地副本，请核对后手动清理。');
        const remaining = source.profiles.filter((p) => !ids.includes(p.id));
        if (remaining.length) {
          saveProfiles(
            {
              profiles: remaining,
              activeId: remaining.some((p) => p.id === source.activeId)
                ? source.activeId
                : remaining[0].id,
            },
            local,
            session,
          );
        } else {
          local.removeItem(PROFILE_KEY);
          session.removeItem(SESSION_KEY);
        }
      } catch (error) {
        localWarning =
          error instanceof Error && error.message.includes('发生变化')
            ? error.message
            : '共享迁移已成功，但旧浏览器副本未能清理，请核对后清理该应用的旧配置。';
      }
      const next = {
        ...migrated,
        legacyAvailable: hasLegacyProfiles(local),
        warning: localWarning ?? migrated.warning,
      };
      publish(next);
      return next;
    },
  };
  const onVisible = () => {
    if (document.visibilityState === 'visible') backgroundRefresh();
  };
  const onSharedState = (event: Event) => {
    try {
      apply(parseSharedState((event as CustomEvent<unknown>).detail));
    } catch {
      /* Ignore malformed internal notifications; never replace good state. */
    }
  };
  return {
    getSnapshot: () => state,
    subscribe(listener: () => void) {
      subscribers.add(listener);
      return () => {
        subscribers.delete(listener);
      };
    },
    actions,
    start() {
      if (lifetime) return;
      lifetime = new AbortController();
      if (!capability()) return;
      backgroundRefresh();
      interval = setInterval(backgroundRefresh, options.pollMilliseconds ?? 5000);
      document.addEventListener('visibilitychange', onVisible);
      window.addEventListener('focus', backgroundRefresh);
      window.addEventListener(SHARED_STATE_EVENT, onSharedState);
    },
    stop() {
      lifetime?.abort();
      lifetime = undefined;
      refreshing = undefined;
      if (interval) clearInterval(interval);
      interval = undefined;
      document.removeEventListener('visibilitychange', onVisible);
      window.removeEventListener('focus', backgroundRefresh);
      window.removeEventListener(SHARED_STATE_EVENT, onSharedState);
    },
    saveLocal(next: LocalProfileState) {
      if (state.mode === 'shared' || state.mode === 'checking')
        throw new Error('共享模式不能写入旧浏览器配置。');
      saveProfiles(next, local, session);
      publish(next);
    },
    async activate(id: string) {
      if (state.mode === 'checking') throw new Error('正在连接共享配置，请稍候。');
      if (state.mode === 'shared') {
        const revision = state.revision;
        await mutate(() => client.activate({ expectedRevision: revision, id }, lifetime?.signal));
      } else {
        const next = { ...state, activeId: id };
        saveProfiles(next, local, session);
        publish(next);
      }
    },
  };
}

export function importProfiles(raw: string): ModelProfile[] {
  if (raw.length > 100_000) throw new Error('配置文件过大。');
  let data;
  try {
    data = JSON.parse(raw);
  } catch {
    throw new Error('请选择有效的 JSON 配置文件。');
  }
  if (
    !data ||
    data.version !== 1 ||
    !Array.isArray(data.profiles) ||
    !data.profiles.length ||
    data.profiles.length > 30
  )
    throw new Error('配置文件需包含 1–30 个模型，格式版本为 1。');
  return data.profiles.map((item: unknown) => ({
    ...parseProfile(item),
    id: crypto.randomUUID(),
    apiKey: '',
    rememberKey: false,
  }));
}
