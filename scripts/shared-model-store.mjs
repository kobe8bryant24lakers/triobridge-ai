import { randomUUID } from 'node:crypto';
import { lstat, open, readFile, rename, unlink } from 'node:fs/promises';
import { homedir } from 'node:os';
import { join } from 'node:path';
import { createNativeModelVault } from './shared-model-vault.mjs';
import { ensurePrivateModelDirectory } from './local-capability.mjs';
import { enforcePrivatePath } from './private-local-permissions.mjs';
import { modelConnectionError } from './model-connection-error.mjs';

export class SharedModelError extends Error {
  constructor(code, message, status = 400, state) {
    super(message);
    this.name = 'SharedModelError';
    this.code = code;
    this.status = status;
    this.state = state;
  }
}

const fail = (code, message, status = 400) => {
  throw new SharedModelError(code, message, status);
};
const isObject = (value) => value !== null && typeof value === 'object' && !Array.isArray(value);
const idPattern = /^[a-zA-Z0-9_-]{1,100}$/;
const refPattern = /^[a-zA-Z0-9_-]{1,100}\.[a-f0-9-]{36}$/;
const integer = (value) => Number.isSafeInteger(value) && value >= 0;
const empty = () => ({
  version: 2,
  revision: 0,
  activeId: '',
  profiles: [],
  retiredCredentials: [],
});

export function sharedModelsDirectory(platform = process.platform) {
  if (platform === 'darwin')
    return join(homedir(), 'Library', 'Application Support', 'TrioBridge AI', 'models');
  if (platform === 'win32' && process.env.LOCALAPPDATA)
    return join(process.env.LOCALAPPDATA, 'TrioBridgeAI', 'models');
  throw new Error('Shared models require a supported per-user application data directory.');
}

function validateBaseUrl(baseUrl, allowProxy = true) {
  if (typeof baseUrl !== 'string' || baseUrl.length > 2048) fail('config', '模型服务地址无效。');
  if (allowProxy && /^\/model-proxy(?:\/[\w-]+)*$/.test(baseUrl.trim())) return;
  let url;
  try {
    url = new URL(baseUrl.trim());
  } catch {
    fail('config', '共享模型需要完整服务地址或已配置的本机代理路径。');
  }
  const loopback = ['localhost', '127.0.0.1', '[::1]'].includes(url.hostname);
  if (
    !['http:', 'https:'].includes(url.protocol) ||
    (url.protocol === 'http:' && !loopback) ||
    url.username ||
    url.password ||
    url.search ||
    url.hash
  )
    fail('config', '远程模型需要 HTTPS，HTTP 仅限本机回环地址；服务地址不能含凭据、参数或片段。');
}

export function sharedCompletionUrl(baseUrl, proxyTarget) {
  validateBaseUrl(baseUrl);
  const base = baseUrl
    .trim()
    .replace(/\/+$/, '')
    .replace(/\/chat\/completions$/, '');
  if (base.startsWith('/model-proxy')) {
    if (!proxyTarget) fail('config', '本机代理尚未配置 MODEL_PROXY_TARGET。', 503);
    validateBaseUrl(proxyTarget, false);
    const target = new URL(proxyTarget);
    if (target.pathname !== '/') fail('config', '本机代理目标必须是服务根地址。');
    return new URL(`${base.slice('/model-proxy'.length)}/chat/completions`, target).href;
  }
  return `${base}/chat/completions`;
}

function profileFields(value) {
  if (
    !isObject(value) ||
    Object.keys(value).some(
      (key) => !['id', 'name', 'baseUrl', 'model', 'timeoutSeconds'].includes(key),
    )
  )
    fail('config', '模型配置包含不支持的字段。');
  if (
    typeof value.id !== 'string' ||
    !idPattern.test(value.id) ||
    typeof value.name !== 'string' ||
    !value.name.trim() ||
    value.name.length > 100 ||
    typeof value.model !== 'string' ||
    !value.model.trim() ||
    value.model.length > 200 ||
    !Number.isInteger(value.timeoutSeconds) ||
    value.timeoutSeconds < 5 ||
    value.timeoutSeconds > 300
  )
    fail('config', '请检查模型名称、ID 和 5–300 秒的超时设置。');
  validateBaseUrl(value.baseUrl);
  return {
    id: value.id,
    name: value.name.trim(),
    baseUrl: value.baseUrl.trim().replace(/\/+$/, ''),
    model: value.model.trim(),
    timeoutSeconds: value.timeoutSeconds,
  };
}

function credentialFields(value) {
  if (
    !isObject(value) ||
    !['keep', 'replace', 'remove'].includes(value.action) ||
    Object.keys(value).some((key) => !['action', 'value'].includes(key))
  )
    fail('config', '请选择保留、替换或删除密钥。');
  if (value.action === 'replace') {
    if (typeof value.value !== 'string' || !value.value.trim() || value.value.length > 4096)
      fail('config', '请输入有效的新 API Key。');
    return { action: 'replace', value: value.value.trim() };
  }
  if ('value' in value) fail('config', '只有替换密钥时可以提交 API Key。');
  return { action: value.action };
}

export function validateSharedMessages(messages) {
  if (!Array.isArray(messages) || !messages.length || messages.length > 16)
    fail('request', '模型请求需要 1–16 条消息。');
  let length = 0,
    imageLength = 0,
    imageCount = 0;
  return messages.map((message) => {
    if (
      !isObject(message) ||
      Object.keys(message).some((key) => !['role', 'content'].includes(key)) ||
      !['system', 'user'].includes(message.role)
    )
      fail('request', '模型消息格式无效。');
    let content;
    if (typeof message.content === 'string') {
      length += message.content.length;
      content = message.content;
    } else if (
      message.role === 'user' &&
      Array.isArray(message.content) &&
      message.content.length > 0 &&
      message.content.length <= 12
    ) {
      content = message.content.map((part) => {
        if (!isObject(part)) fail('request', '模型消息内容无效。');
        if (
          part.type === 'text' &&
          typeof part.text === 'string' &&
          Object.keys(part).every((key) => ['type', 'text'].includes(key))
        ) {
          length += part.text.length;
          return { type: 'text', text: part.text };
        }
        const image = part.image_url;
        if (
          part.type !== 'image_url' ||
          Object.keys(part).some((key) => !['type', 'image_url'].includes(key)) ||
          !isObject(image) ||
          Object.keys(image).some((key) => !['url', 'detail'].includes(key)) ||
          typeof image.url !== 'string' ||
          !/^data:image\/(?:jpeg|png);base64,[A-Za-z0-9+/]+={0,2}$/.test(image.url) ||
          !['low', 'high', 'auto'].includes(image.detail)
        )
          fail('request', '只支持本次选择的 PNG / JPEG 图片数据，不支持远程图片地址。');
        imageLength += image.url.length;
        if (++imageCount > 4 || imageLength > 200_000 || image.url.length > 50_000)
          fail('request', '图片过多或过大，请减少图片数量。', 413);
        return { type: 'image_url', image_url: { url: image.url, detail: image.detail } };
      });
    } else fail('request', '模型消息格式无效。');
    if (length > 150_000) fail('request', '模型请求内容过长。', 413);
    return { role: message.role, content };
  });
}

function usageFields(value) {
  if (!isObject(value)) return undefined;
  const result = {};
  for (const name of ['prompt_tokens', 'completion_tokens', 'total_tokens'])
    if (integer(value[name])) result[name] = value[name];
  for (const name of ['prompt_tokens_details', 'completion_tokens_details']) {
    if (!isObject(value[name])) continue;
    const details = {};
    for (const key of [
      'cached_tokens',
      'audio_tokens',
      'text_tokens',
      'reasoning_tokens',
      'accepted_prediction_tokens',
      'rejected_prediction_tokens',
    ])
      if (integer(value[name][key])) details[key] = value[name][key];
    if (Object.keys(details).length) result[name] = details;
  }
  return Object.keys(result).length ? result : undefined;
}

function redact(value, secret) {
  if (typeof value === 'string') return secret ? value.split(secret).join('[REDACTED]') : value;
  if (Array.isArray(value)) return value.map((item) => redact(item, secret));
  if (isObject(value))
    return Object.fromEntries(
      Object.entries(value).map(([key, item]) => [key, redact(item, secret)]),
    );
  return value;
}

async function readResponse(response) {
  const reader = response.body?.getReader();
  const chunks = [];
  let bytes = 0;
  if (reader)
    for (;;) {
      const { done, value } = await reader.read();
      if (done) break;
      bytes += value.byteLength;
      if (bytes > 1_000_000) {
        await reader.cancel();
        fail('response', '模型响应过大，请缩小任务。', 502);
      }
      chunks.push(Buffer.from(value));
    }
  try {
    return JSON.parse(Buffer.concat(chunks).toString('utf8'));
  } catch {
    fail('response', '模型服务返回了无效 JSON。', 502);
  }
}

async function requestCompletion(profile, secret, messages, signal, fetcher) {
  const controller = new AbortController();
  let timeout = false;
  const abort = () => controller.abort();
  signal?.addEventListener('abort', abort, { once: true });
  if (signal?.aborted) controller.abort();
  const timer = setTimeout(() => {
    timeout = true;
    controller.abort();
  }, profile.timeoutSeconds * 1000);
  const endpoint = profile.endpoint || sharedCompletionUrl(profile.baseUrl);
  try {
    const response = await fetcher(endpoint, {
      method: 'POST',
      redirect: 'error',
      credentials: 'omit',
      cache: 'no-store',
      signal: controller.signal,
      headers: {
        'Content-Type': 'application/json',
        ...(secret ? { Authorization: `Bearer ${secret}` } : {}),
      },
      body: JSON.stringify({ model: profile.model, messages, stream: false }),
    });
    if (!response.ok) {
      await response.body?.cancel();
      const errors = {
        400: '模型服务拒绝了请求，请检查模型 ID。',
        401: 'API Key 无效或已过期。',
        403: 'API Key 没有此模型的访问权限。',
        404: '找不到模型接口或模型。',
        429: '请求过于频繁或额度不足，请稍后重试。',
      };
      fail(
        'upstream',
        errors[response.status] || `模型服务返回 HTTP ${response.status}。`,
        response.status >= 400 ? response.status : 502,
      );
    }
    const data = await readResponse(response);
    const choice = isObject(data) && Array.isArray(data.choices) ? data.choices[0] : undefined;
    if (
      !isObject(choice) ||
      !isObject(choice.message) ||
      (typeof choice.message.content !== 'string' && typeof choice.message.refusal !== 'string')
    )
      fail('response', '模型服务未返回兼容的文本响应。', 502);
    if (typeof choice.message.content === 'string' && choice.message.content.length > 100_000)
      fail('response', '模型输出超过 100,000 字符。', 502);
    const body = {
      ...(typeof data.id === 'string' ? { id: data.id.slice(0, 300) } : {}),
      ...(typeof data.model === 'string' ? { model: data.model.slice(0, 200) } : {}),
      choices: [
        {
          message: {
            ...(typeof choice.message.content === 'string'
              ? { content: choice.message.content }
              : {}),
            ...(typeof choice.message.refusal === 'string'
              ? { refusal: choice.message.refusal.slice(0, 1000) }
              : {}),
          },
          ...(typeof choice.finish_reason === 'string'
            ? { finish_reason: choice.finish_reason.slice(0, 80) }
            : {}),
        },
      ],
      ...(usageFields(data.usage) ? { usage: usageFields(data.usage) } : {}),
      localModel: {
        profileId: profile.id,
        profileRevision: profile.revision,
        endpoint,
        requestedModel: profile.model,
      },
    };
    return {
      body: redact(body, secret),
      requestId: redact(response.headers.get('x-request-id')?.slice(0, 300) || null, secret),
    };
  } catch (error) {
    if (timeout)
      fail('timeout', `请求超过 ${profile.timeoutSeconds} 秒，请重试或增加超时时间。`, 504);
    if (signal?.aborted) fail('cancelled', '已取消生成。', 499);
    if (error instanceof SharedModelError) throw error;
    fail('network', redact(modelConnectionError(error, endpoint), secret), 502);
  } finally {
    clearTimeout(timer);
    signal?.removeEventListener('abort', abort);
  }
}

/** Single metadata catalog, optimistic revisions, and cross-process serialization. */
export function createSharedModelStore({
  directory = sharedModelsDirectory(),
  vault = createNativeModelVault(),
  fetcher = fetch,
  proxyTarget = process.env.MODEL_PROXY_TARGET,
} = {}) {
  const file = join(directory, 'profiles.json');
  const lockFile = join(directory, 'profiles.lock');

  async function ensureDirectory() {
    try {
      await ensurePrivateModelDirectory(directory);
    } catch {
      fail('storage', '共享配置目录不安全，请检查目录权限。', 503);
    }
  }
  async function load() {
    await ensureDirectory();
    let value;
    try {
      const info = await lstat(file);
      if (
        !info.isFile() ||
        info.isSymbolicLink() ||
        (process.getuid && info.uid !== process.getuid())
      )
        fail('storage', '共享配置文件不安全，原始文件未修改。', 503);
      if (info.size > 200_000) fail('storage', '共享配置文件过大，原始文件未修改。', 503);
      await enforcePrivatePath(file, 'file');
      value = JSON.parse(await readFile(file, 'utf8'));
    } catch (error) {
      if (error.code === 'ENOENT') return empty();
      if (error instanceof SharedModelError) throw error;
      fail('storage', '共享配置无法读取，原始文件未被覆盖。', 503);
    }
    try {
      if (
        !isObject(value) ||
        value.version !== 2 ||
        !integer(value.revision) ||
        !Array.isArray(value.profiles) ||
        value.profiles.length > 30 ||
        !Array.isArray(value.retiredCredentials) ||
        value.retiredCredentials.length > 300 ||
        !value.retiredCredentials.every((ref) => typeof ref === 'string' && refPattern.test(ref))
      )
        throw new Error();
      const profiles = value.profiles.map((item) => {
        const { revision, credentialRef, endpoint, ...fields } = item;
        if (
          !integer(revision) ||
          revision > value.revision ||
          (credentialRef !== null &&
            (typeof credentialRef !== 'string' || !refPattern.test(credentialRef)))
        )
          throw new Error();
        validateBaseUrl(endpoint, false);
        if (sharedCompletionUrl(endpoint) !== endpoint) throw new Error();
        return { ...profileFields(fields), revision, credentialRef, endpoint };
      });
      if (
        new Set(profiles.map((p) => p.id)).size !== profiles.length ||
        (value.activeId !== '' && !profiles.some((p) => p.id === value.activeId)) ||
        profiles.some((p) => p.credentialRef && value.retiredCredentials.includes(p.credentialRef))
      )
        throw new Error();
      return {
        version: 2,
        revision: value.revision,
        activeId: value.activeId,
        profiles,
        retiredCredentials: value.retiredCredentials,
      };
    } catch {
      fail('storage', '共享配置格式损坏，原始文件未被覆盖。', 503);
    }
  }
  function publicState(state) {
    return {
      version: 2,
      revision: state.revision,
      activeId: state.activeId,
      storage: vault.kind,
      profiles: state.profiles.map(({ credentialRef, endpoint: _endpoint, ...profile }) => ({
        ...profile,
        hasKey: Boolean(credentialRef),
      })),
      ...(state.retiredCredentials.length
        ? { warning: '配置已更新，部分旧密钥尚待从系统安全存储清理；请重试清理。' }
        : {}),
    };
  }
  async function write(state) {
    const temporary = join(directory, `.profiles-${randomUUID()}.tmp`);
    let handle;
    try {
      handle = await open(temporary, 'wx', 0o600);
      await handle.writeFile(JSON.stringify(state));
      await handle.sync();
      await handle.close();
      handle = undefined;
      await enforcePrivatePath(temporary, 'file', { created: true });
      await rename(temporary, file);
    } catch {
      fail('storage', '无法保存共享配置，请检查本机目录权限和可用空间。', 503);
    } finally {
      await handle?.close().catch(() => {});
      await unlink(temporary).catch(() => {});
    }
  }
  async function locked(operation) {
    await ensureDirectory();
    let handle;
    for (let attempt = 0; attempt < 2 && !handle; attempt++) {
      try {
        handle = await open(lockFile, 'wx', 0o600);
      } catch (error) {
        if (error.code !== 'EEXIST') fail('storage', '无法锁定共享配置，请检查目录权限。', 503);
        let stale = false;
        let info;
        try {
          info = await lstat(lockFile);
          if (!info.isFile() || info.isSymbolicLink())
            fail('storage', '共享配置锁文件不安全。', 503);
          try {
            const lock = JSON.parse(await readFile(lockFile, 'utf8'));
            if (Number.isSafeInteger(lock.pid) && lock.pid > 0) {
              try {
                process.kill(lock.pid, 0);
              } catch (failure) {
                stale = failure.code === 'ESRCH';
              }
            } else stale = Date.now() - info.mtimeMs > 60_000;
          } catch {
            stale = Date.now() - info.mtimeMs > 60_000;
          }
          if (stale && attempt === 0) {
            const current = await lstat(lockFile);
            if (current.ino === info.ino && current.mtimeMs === info.mtimeMs) {
              await unlink(lockFile);
              continue;
            }
          }
        } catch (failure) {
          if (failure instanceof SharedModelError) throw failure;
          if (failure.code === 'ENOENT' && attempt === 0) continue;
        }
        fail('busy', '其他 Office 正在保存模型配置，请刷新后重试。', 409);
      }
    }
    if (!handle) fail('busy', '共享配置正在更新，请重试。', 409);
    try {
      await handle.writeFile(
        JSON.stringify({ pid: process.pid, createdAt: new Date().toISOString() }),
      );
      return await operation();
    } finally {
      await handle.close().catch(() => {});
      await unlink(lockFile).catch(() => {});
    }
  }
  function revision(state, expected) {
    if (!integer(expected)) fail('request', '缺少有效的配置版本。');
    if (state.revision !== expected)
      throw new SharedModelError(
        'conflict',
        '模型配置已在另一 Office 中更改。请刷新后重新确认本次修改。',
        409,
        publicState(state),
      );
  }
  async function secret(profile, signal) {
    if (!profile.credentialRef) return '';
    let value;
    try {
      value = await vault.get(profile.credentialRef, signal);
    } catch {
      if (signal?.aborted) fail('cancelled', '已取消生成。', 499);
      fail('credential', '系统安全存储不可用或未解锁，密钥没有降级保存到文件。', 503);
    }
    if (typeof value !== 'string' || !value)
      fail('credential', '此模型的系统密钥已缺失，请重新填写 API Key。', 503);
    return value;
  }
  async function cleanup(state) {
    const retained = [];
    for (const reference of state.retiredCredentials) {
      try {
        await vault.delete(reference);
      } catch {
        retained.push(reference);
      }
    }
    if (retained.length === state.retiredCredentials.length) return state;
    const cleaned = { ...state, retiredCredentials: retained };
    try {
      await write(cleaned);
      return cleaned;
    } catch {
      return state;
    }
  }
  function candidate(profile, credential, old) {
    const endpoint = sharedCompletionUrl(profile.baseUrl, proxyTarget);
    if (credential.action === 'keep' && old && endpoint !== old.endpoint)
      fail('credential', '更换服务地址后需要重新填写或明确移除密钥，不能沿用旧服务的密钥。');
    return {
      ...profile,
      endpoint,
      credentialRef:
        credential.action === 'replace'
          ? `${profile.id}.${randomUUID()}`
          : credential.action === 'keep'
            ? old?.credentialRef || null
            : null,
    };
  }
  async function commit(state, next, newCredentials, signal) {
    if (signal?.aborted) fail('cancelled', '已取消保存。', 499);
    if (newCredentials.length) {
      const journal = {
        ...state,
        retiredCredentials: [
          ...state.retiredCredentials,
          ...newCredentials.map((item) => item.reference),
        ],
      };
      await write(journal);
      try {
        for (const item of newCredentials) {
          if (signal?.aborted) fail('cancelled', '已取消保存。', 499);
          await vault.set(item.reference, item.value, signal);
        }
        if (signal?.aborted) fail('cancelled', '已取消保存。', 499);
        await write(next);
      } catch (error) {
        await cleanup(journal);
        if (error instanceof SharedModelError) throw error;
        fail('credential', '无法保存到系统安全存储，本次配置未生效。', 503);
      }
    } else {
      if (signal?.aborted) fail('cancelled', '已取消保存。', 499);
      await write(next);
    }
    return publicState(await cleanup(next));
  }

  return {
    async state() {
      return publicState(await load());
    },
    async upsert(input, signal) {
      const profile = profileFields(input.profile);
      const credential = credentialFields(input.credential);
      let connectionTest;
      if (input.testConnection === true) {
        const before = await load();
        revision(before, input.expectedRevision);
        const old = before.profiles.find((p) => p.id === profile.id);
        const proposed = candidate(profile, credential, old);
        const key =
          credential.action === 'replace'
            ? credential.value
            : credential.action === 'keep' && old
              ? await secret(old, signal)
              : '';
        connectionTest = await requestCompletion(
          { ...proposed, revision: old?.revision || 0 },
          key,
          [{ role: 'user', content: 'Reply with OK.' }],
          signal,
          fetcher,
        );
        const choice = connectionTest.body.choices[0];
        if (
          !choice.message.content?.trim() ||
          choice.message.refusal ||
          ['length', 'content_filter'].includes(choice.finish_reason)
        )
          fail('response', '连接测试未返回完整可用的文本，本次配置未保存。', 502);
      }
      const result = await locked(async () => {
        const current = await load();
        revision(current, input.expectedRevision);
        const old = current.profiles.find((p) => p.id === profile.id);
        if (!old && current.profiles.length >= 30) fail('config', '最多保存 30 个模型配置。');
        const proposed = { ...candidate(profile, credential, old), revision: current.revision + 1 };
        const retired =
          old?.credentialRef && old.credentialRef !== proposed.credentialRef
            ? [old.credentialRef]
            : [];
        const next = {
          ...current,
          revision: current.revision + 1,
          activeId: input.activate === true || !current.activeId ? profile.id : current.activeId,
          profiles: old
            ? current.profiles.map((p) => (p.id === profile.id ? proposed : p))
            : [...current.profiles, proposed],
          retiredCredentials: [...current.retiredCredentials, ...retired],
        };
        return commit(
          current,
          next,
          credential.action === 'replace'
            ? [{ reference: proposed.credentialRef, value: credential.value }]
            : [],
          signal,
        );
      });
      return { ...result, ...(connectionTest ? { connectionTest: connectionTest.body } : {}) };
    },
    async import(input, signal) {
      if (!Array.isArray(input.profiles) || !input.profiles.length || input.profiles.length > 30)
        fail('config', '一次可导入 1–30 个模型。');
      const items = input.profiles.map((item) => ({
        profile: profileFields(item.profile),
        credential: credentialFields(item.credential),
      }));
      if (new Set(items.map((item) => item.profile.id)).size !== items.length)
        fail('config', '导入内容包含重复模型 ID。');
      return locked(async () => {
        const current = await load();
        revision(current, input.expectedRevision);
        if (current.profiles.length + items.length > 30)
          fail('config', '共享模型总数不能超过 30 个。');
        if (items.some((item) => current.profiles.some((p) => p.id === item.profile.id)))
          throw new SharedModelError(
            'conflict',
            '导入包含已有模型 ID，不会覆盖已有配置。',
            409,
            publicState(current),
          );
        const profiles = items.map((item) => ({
          ...candidate(item.profile, item.credential),
          revision: current.revision + 1,
        }));
        const activeId = input.activeId ?? (current.activeId || profiles[0].id);
        if (![...current.profiles, ...profiles].some((p) => p.id === activeId))
          fail('config', '导入后的默认模型不存在。');
        const next = {
          ...current,
          revision: current.revision + 1,
          activeId,
          profiles: [...current.profiles, ...profiles],
        };
        const credentials = items.flatMap((item, index) =>
          item.credential.action === 'replace'
            ? [{ reference: profiles[index].credentialRef, value: item.credential.value }]
            : [],
        );
        return commit(current, next, credentials, signal);
      });
    },
    async delete(input, signal) {
      return locked(async () => {
        const current = await load();
        revision(current, input.expectedRevision);
        const profile = current.profiles.find((p) => p.id === input.id);
        if (!profile) fail('missing', '模型已被删除或不存在。', 404);
        const profiles = current.profiles.filter((p) => p.id !== input.id);
        const next = {
          ...current,
          revision: current.revision + 1,
          profiles,
          activeId: current.activeId === input.id ? profiles[0]?.id || '' : current.activeId,
          retiredCredentials: [
            ...current.retiredCredentials,
            ...(profile.credentialRef ? [profile.credentialRef] : []),
          ],
        };
        return commit(current, next, [], signal);
      });
    },
    async activate(input, signal) {
      return locked(async () => {
        const current = await load();
        revision(current, input.expectedRevision);
        if (!current.profiles.some((p) => p.id === input.id))
          fail('missing', '模型已被删除或不存在。', 404);
        if (current.activeId === input.id) return publicState(current);
        return commit(
          current,
          { ...current, revision: current.revision + 1, activeId: input.id },
          [],
          signal,
        );
      });
    },
    async complete(input, signal) {
      const messages = validateSharedMessages(input.messages);
      const state = await load();
      const profile = state.profiles.find((p) => p.id === input.profileId);
      if (!profile) fail('missing', '模型已被删除或不存在，请刷新模型列表。', 404);
      if (input.profileRevision !== profile.revision)
        throw new SharedModelError(
          'conflict',
          '此模型已在另一 Office 中更新，请刷新后再生成。',
          409,
          publicState(state),
        );
      return requestCompletion(profile, await secret(profile, signal), messages, signal, fetcher);
    },
    async cleanup() {
      return locked(async () => publicState(await cleanup(await load())));
    },
  };
}
