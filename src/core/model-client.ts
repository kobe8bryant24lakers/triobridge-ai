import { modelTransport } from '../offline/transport';
import type { ChatMessage, RuntimeModelProfile, SharedModelProfile } from './types';
import { MAX_RESULT_LENGTH } from './tasks';
import { startModelAudit } from './model-audit';
import {
  localModelHeaders,
  sharedResponseError,
  SharedModelClientError,
} from './shared-model-client';

export function isSharedProfile(profile: RuntimeModelProfile): profile is SharedModelProfile {
  return 'kind' in profile && profile.kind === 'shared';
}

export class ModelError extends Error {
  constructor(
    message: string,
    public readonly code:
      'config' | 'auth' | 'rate' | 'network' | 'timeout' | 'cancelled' | 'response',
  ) {
    super(message);
    this.name = 'ModelError';
  }
}

export function completionUrl(baseUrl: string): string {
  const base = baseUrl.trim().replace(/\/+$/, '');
  if (/^\/model-proxy(?:\/[\w-]+)*$/.test(base))
    return `${base.replace(/\/chat\/completions$/, '')}/chat/completions`;
  let url: URL;
  try {
    url = new URL(base);
  } catch {
    throw new ModelError('请填写完整的 API Base URL，例如 https://api.example.com/v1。', 'config');
  }
  if (
    !['http:', 'https:'].includes(url.protocol) ||
    url.username ||
    url.password ||
    url.search ||
    url.hash
  )
    throw new ModelError('服务地址只支持 HTTP(S)，不能包含账号、密码、查询参数或片段。', 'config');
  const local = ['localhost', '127.0.0.1', '[::1]'].includes(url.hostname);
  if (url.protocol === 'http:' && !local)
    throw new ModelError('远程模型服务需要 HTTPS；本地 HTTP 服务请使用本机代理。', 'config');
  return `${base.replace(/\/chat\/completions$/, '')}/chat/completions`;
}

export function validateProfile(profile: RuntimeModelProfile): void {
  if (!profile.name.trim() || !profile.model.trim())
    throw new ModelError('请填写服务名称和模型 ID。', 'config');
  if (
    profile.name.length > 100 ||
    profile.model.length > 200 ||
    (!isSharedProfile(profile) && profile.apiKey.length > 4096)
  )
    throw new ModelError('模型配置字段过长，请检查输入。', 'config');
  completionUrl(profile.baseUrl);
  if (isSharedProfile(profile) && (!Number.isSafeInteger(profile.revision) || profile.revision < 0))
    throw new ModelError('共享模型版本无效，请刷新配置后重试。', 'config');
  if (
    !Number.isFinite(profile.timeoutSeconds) ||
    profile.timeoutSeconds < 5 ||
    profile.timeoutSeconds > 300
  )
    throw new ModelError('超时应为 5–300 秒。', 'config');
}

const HTTP_ERRORS: Record<number, string> = {
  400: '服务拒绝了请求，请检查模型 ID 和接口参数。',
  401: 'API Key 无效或已过期，请在模型设置中更新。',
  403: '当前 API Key 没有访问权限，请检查模型权限。',
  404: '找不到接口或模型，请检查 Base URL（通常包含 /v1）和模型 ID。',
  429: '请求过于频繁或额度不足，请稍后重试并检查服务额度。',
};

export async function complete(
  profile: RuntimeModelProfile,
  messages: ChatMessage[],
  signal?: AbortSignal,
  fetcher: typeof fetch = modelTransport,
): Promise<string> {
  validateProfile(profile);
  if (signal?.aborted) throw new ModelError('已取消生成。', 'cancelled');
  const endpoint = completionUrl(profile.baseUrl);
  const secret = isSharedProfile(profile) ? '' : profile.apiKey.trim();
  const audit = startModelAudit(endpoint, profile.model.trim(), messages, secret);
  let auditOutcome = 'network';
  const controller = new AbortController();
  let timedOut = false;
  const abort = () => controller.abort();
  signal?.addEventListener('abort', abort, { once: true });
  const timer = setTimeout(() => {
    timedOut = true;
    controller.abort();
  }, profile.timeoutSeconds * 1000);
  try {
    const shared = isSharedProfile(profile);
    const response = await fetcher(shared ? '/local-models/complete' : endpoint, {
      method: 'POST',
      signal: controller.signal,
      redirect: 'error',
      credentials: 'omit',
      cache: 'no-store',
      headers: shared
        ? localModelHeaders()
        : {
            'Content-Type': 'application/json',
            ...(secret ? { Authorization: `Bearer ${secret}` } : {}),
          },
      body: JSON.stringify(
        shared
          ? {
              profileId: profile.id,
              profileRevision: profile.revision,
              messages,
            }
          : {
              model: profile.model.trim(),
              messages,
              stream: false,
            },
      ),
    });
    audit?.response(response);
    if (!response.ok && shared) {
      const error = await sharedResponseError(response);
      throw new ModelError(
        error.message,
        response.status === 401 || response.status === 403
          ? 'auth'
          : response.status === 429
            ? 'rate'
            : error.code === 'timeout'
              ? 'timeout'
              : error.code === 'cancelled'
                ? 'cancelled'
                : error.code === 'network'
                  ? 'network'
                  : response.status === 409 || error.code === 'config'
                    ? 'config'
                    : 'response',
      );
    }
    if (!response.ok)
      throw new ModelError(
        HTTP_ERRORS[response.status] ?? `模型服务返回 HTTP ${response.status}，请稍后重试。`,
        response.status === 401 || response.status === 403
          ? 'auth'
          : response.status === 429
            ? 'rate'
            : 'response',
      );
    const data = (await response.json()) as {
      choices?: {
        message?: { content?: unknown; refusal?: unknown };
        finish_reason?: string;
      }[];
    };
    audit?.payload(data);
    if (!data || typeof data !== 'object')
      throw new ModelError('服务返回的数据格式无效，请确认兼容 Chat Completions。', 'response');
    const choice = data.choices?.[0];
    if (choice?.finish_reason === 'length')
      throw new ModelError(
        '模型输出被长度限制截断，请缩短任务或调整服务端输出上限后重试。',
        'response',
      );
    if (choice?.message?.refusal || choice?.finish_reason === 'content_filter')
      throw new ModelError('模型未完成此请求，请调整输入后重试。', 'response');
    const content = choice?.message?.content;
    if (typeof content !== 'string' || !content.trim())
      throw new ModelError('服务未返回可用文本，请确认接口兼容 Chat Completions。', 'response');
    if (content.length > MAX_RESULT_LENGTH)
      throw new ModelError('模型返回内容过长，请缩小任务范围。', 'response');
    auditOutcome = 'success';
    return content.trim();
  } catch (error) {
    auditOutcome = timedOut
      ? 'timeout'
      : signal?.aborted
        ? 'cancelled'
        : error instanceof ModelError
          ? error.code
          : error instanceof SyntaxError
            ? 'response'
            : 'network';
    if (timedOut)
      throw new ModelError(
        `请求超过 ${profile.timeoutSeconds} 秒，请重试或增加超时时间。`,
        'timeout',
      );
    if (signal?.aborted) throw new ModelError('已取消生成。', 'cancelled');
    if (error instanceof ModelError) throw error;
    if (error instanceof SharedModelClientError)
      throw new ModelError(error.message, error.code === 'capability' ? 'config' : 'response');
    if (error instanceof SyntaxError)
      throw new ModelError('服务返回了无效 JSON，请检查接口地址。', 'response');
    if (isSharedProfile(profile))
      throw new ModelError(
        '无法连接 TrioBridge 本机共享服务。请双击本机共享安装包中的 Restart，重启服务后重新连接。',
        'network',
      );
    if (/^https?:\/\/(?:localhost|127\.0\.0\.1|\[::1\])(?=[:/])/i.test(endpoint))
      throw new ModelError(
        '当前加载项正在直连本地 API，连接失败。请打开「三桥 AI（本机共享）」并确认模型设置显示「本机共享」；共享版通过本机服务调用，不要求 API 开放浏览器 CORS。',
        'network',
      );
    throw new ModelError(
      '无法连接模型服务。请检查地址、网络和 CORS；本地 HTTP 模型可配置本机代理。',
      'network',
    );
  } finally {
    clearTimeout(timer);
    signal?.removeEventListener('abort', abort);
    await audit?.finish(auditOutcome);
  }
}

export async function testConnection(
  profile: RuntimeModelProfile,
  signal?: AbortSignal,
): Promise<void> {
  await complete(profile, [{ role: 'user', content: 'Reply with OK.' }], signal);
}
