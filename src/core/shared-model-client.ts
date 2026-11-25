import { nativeWorkbench } from '../offline/transport';
import type { ProfileDetails, SharedModelProfile } from './types';
import { getLocalCapability } from './local-capability';
import { startModelAudit } from './model-audit';

export type SharedCredential = { action: 'keep' | 'remove' } | { action: 'replace'; value: string };
export interface SharedUpsertRequest {
  expectedRevision: number;
  profile: ProfileDetails;
  credential: SharedCredential;
  activate?: boolean;
  testConnection?: boolean;
}
export interface SharedImportRequest {
  expectedRevision: number;
  profiles: { profile: ProfileDetails; credential: SharedCredential }[];
  activeId?: string;
}
export interface SharedStateSnapshot {
  version: 2;
  revision: number;
  activeId: string;
  profiles: SharedModelProfile[];
  storage: 'macos-keychain' | 'windows-credential-manager';
  warning?: string;
}

export class SharedModelClientError extends Error {
  constructor(
    message: string,
    public readonly code: string,
    public readonly status?: number,
    public readonly state?: SharedStateSnapshot,
  ) {
    super(message);
    this.name = 'SharedModelClientError';
  }
}

function object(value: unknown): value is Record<string, unknown> {
  return Boolean(value) && typeof value === 'object' && !Array.isArray(value);
}

export function parseSharedState(value: unknown): SharedStateSnapshot {
  const invalid = () => {
    throw new SharedModelClientError('读取的共享模型配置无效，请重新加载。', 'response');
  };
  if (
    !object(value) ||
    value.version !== 2 ||
    !Number.isSafeInteger(value.revision) ||
    (value.revision as number) < 0 ||
    typeof value.activeId !== 'string' ||
    !Array.isArray(value.profiles) ||
    value.profiles.length > 30 ||
    !['macos-keychain', 'windows-credential-manager'].includes(String(value.storage))
  )
    return invalid();
  const profiles = value.profiles.map((profile): SharedModelProfile => {
    if (
      !object(profile) ||
      'apiKey' in profile ||
      'credential' in profile ||
      typeof profile.id !== 'string' ||
      !/^[a-zA-Z0-9_-]{1,100}$/.test(profile.id) ||
      typeof profile.name !== 'string' ||
      !profile.name.trim() ||
      profile.name.length > 100 ||
      typeof profile.baseUrl !== 'string' ||
      profile.baseUrl.length > 2048 ||
      typeof profile.model !== 'string' ||
      !profile.model.trim() ||
      profile.model.length > 200 ||
      typeof profile.hasKey !== 'boolean' ||
      !Number.isSafeInteger(profile.timeoutSeconds) ||
      (profile.timeoutSeconds as number) < 5 ||
      (profile.timeoutSeconds as number) > 300 ||
      !Number.isSafeInteger(profile.revision) ||
      (profile.revision as number) < 0 ||
      (profile.revision as number) > (value.revision as number)
    )
      return invalid();
    return {
      kind: 'shared',
      id: profile.id,
      name: profile.name,
      baseUrl: profile.baseUrl,
      model: profile.model,
      timeoutSeconds: profile.timeoutSeconds as number,
      hasKey: profile.hasKey,
      revision: profile.revision as number,
    };
  });
  if (
    new Set(profiles.map((p) => p.id)).size !== profiles.length ||
    (value.activeId ? !profiles.some((p) => p.id === value.activeId) : profiles.length > 0)
  )
    return invalid();
  return {
    version: 2,
    revision: value.revision as number,
    activeId: value.activeId,
    profiles,
    storage: value.storage as SharedStateSnapshot['storage'],
    ...(typeof value.warning === 'string' ? { warning: value.warning } : {}),
  };
}

export function localModelHeaders(capability = getLocalCapability()): Record<string, string> {
  if (nativeWorkbench()) return { 'Content-Type': 'application/json' };
  if (!capability)
    throw new SharedModelClientError(
      '请从本机安装的共享加载项打开。当前入口没有本机共享凭据。',
      'capability',
    );
  return { 'Content-Type': 'application/json', 'X-TrioBridge-Local-Token': capability };
}

export const SHARED_STATE_EVENT = 'triobridge:shared-model-state';
export function publishSharedState(state: SharedStateSnapshot): void {
  window.dispatchEvent(new CustomEvent(SHARED_STATE_EVENT, { detail: state }));
}

export async function sharedResponseError(response: Response): Promise<SharedModelClientError> {
  let value: unknown;
  try {
    value = await response.json();
  } catch {
    return new SharedModelClientError(
      '暂时无法读取共享配置，请重新打开模型设置后重试。',
      'unavailable',
      response.status,
    );
  }
  const error = object(value) && object(value.error) ? value.error : undefined;
  let state: SharedStateSnapshot | undefined;
  if (response.status === 409 && object(value) && value.state) {
    state = parseSharedState(value.state);
    publishSharedState(state);
  }
  return new SharedModelClientError(
    typeof error?.message === 'string' ? error.message : '共享模型操作未完成，请重试。',
    typeof error?.code === 'string' ? error.code : 'response',
    response.status,
    state,
  );
}

export function createSharedModelClient(
  fetcher: typeof fetch = fetch,
  capability: () => string | undefined = getLocalCapability,
) {
  async function request(
    action: 'state' | 'upsert' | 'delete' | 'activate' | 'import',
    body: unknown,
    signal?: AbortSignal,
  ): Promise<SharedStateSnapshot> {
    if (signal?.aborted) throw new DOMException('已取消操作。', 'AbortError');
    const headers = localModelHeaders(capability());
    const connection =
      action === 'upsert' && (body as SharedUpsertRequest).testConnection
        ? (body as SharedUpsertRequest)
        : undefined;
    const audit = connection
      ? startModelAudit(
          `${connection.profile.baseUrl.replace(/\/+$/, '').replace(/\/chat\/completions$/, '')}/chat/completions`,
          connection.profile.model,
          [{ role: 'user', content: 'Reply with OK.' }],
          connection.credential.action === 'replace' ? connection.credential.value : '',
        )
      : undefined;
    let auditOutcome = 'network';
    const controller = new AbortController();
    let timedOut = false;
    const abort = () => controller.abort();
    signal?.addEventListener('abort', abort, { once: true });
    const timer = setTimeout(
      () => {
        timedOut = true;
        controller.abort();
      },
      connection
        ? connection.profile.timeoutSeconds * 1000 + 5000
        : action === 'state'
          ? 10000
          : action === 'import'
            ? 120000
            : 30000,
    );
    try {
      const response = await fetcher(`/local-models/${action}`, {
        method: 'POST',
        headers,
        body: JSON.stringify(body),
        signal: controller.signal,
        credentials: 'omit',
        redirect: 'error',
        cache: 'no-store',
      });
      audit?.response(response);
      if (!response.ok) throw await sharedResponseError(response);
      let data: unknown;
      try {
        data = await response.json();
      } catch {
        throw new SharedModelClientError('无法从此入口读取共享配置。', 'unavailable');
      }
      const state = parseSharedState(data);
      if (object(data) && data.connectionTest) audit?.payload(data.connectionTest);
      auditOutcome = 'success';
      return state;
    } catch (error) {
      auditOutcome = timedOut
        ? 'timeout'
        : signal?.aborted
          ? 'cancelled'
          : error instanceof SharedModelClientError
            ? error.status === 401 || error.status === 403
              ? 'auth'
              : error.status === 429
                ? 'rate'
                : error.code === 'conflict' || error.code === 'config'
                  ? 'config'
                  : error.code === 'timeout' || error.code === 'cancelled'
                    ? error.code
                    : 'response'
            : 'network';
      if (timedOut)
        throw new SharedModelClientError(
          connection
            ? '连接超时。请查看系统是否在请求密钥访问，并检查模型 API 是否可用。重新连接后可核对配置。'
            : '读取或保存共享模型配置超时，请重试。',
          'local_timeout',
        );
      if (signal?.aborted) throw new DOMException('已取消操作。', 'AbortError');
      if (error instanceof SharedModelClientError) throw error;
      throw new SharedModelClientError(
        nativeWorkbench()
          ? '工作台未能读取或保存共享配置。请重新打开模型设置，核对后重试。'
          : '无法连接 TrioBridge 本机共享服务。请双击本机共享安装包中的 Restart，重启后重试。本次修改尚未确认保存，请重新连接后核对配置。',
        'local_network',
      );
    } finally {
      clearTimeout(timer);
      signal?.removeEventListener('abort', abort);
      await audit?.finish(auditOutcome);
    }
  }
  return {
    state: (signal?: AbortSignal) => request('state', {}, signal),
    upsert: (input: SharedUpsertRequest, signal?: AbortSignal) => request('upsert', input, signal),
    remove: (input: { expectedRevision: number; id: string }, signal?: AbortSignal) =>
      request('delete', input, signal),
    activate: (input: { expectedRevision: number; id: string }, signal?: AbortSignal) =>
      request('activate', input, signal),
    importProfiles: (input: SharedImportRequest, signal?: AbortSignal) =>
      request('import', input, signal),
  };
}
