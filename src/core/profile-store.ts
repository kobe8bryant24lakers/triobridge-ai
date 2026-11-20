import type { ModelProfile } from './types';
import { validateProfile } from './model-client';

export const PROFILE_KEY = 'triobridge.profiles.v1';
export const SESSION_KEY = 'triobridge.keys.v1';
export interface ProfileState {
  profiles: ModelProfile[];
  activeId: string;
  warning?: string;
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

export function loadProfiles(local?: Storage, session?: Storage): ProfileState {
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

export function saveProfiles(state: ProfileState, local?: Storage, session?: Storage): void {
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

export function exportProfiles(profiles: ModelProfile[]): string {
  return JSON.stringify(
    {
      version: 1,
      profiles: profiles.map((p) => ({ ...p, apiKey: '', rememberKey: false })),
    },
    null,
    2,
  );
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
