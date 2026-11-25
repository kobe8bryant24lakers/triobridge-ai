import { storageDetailMessage } from './shared/model-storage-diagnostic.mjs';
import { SharedModelError } from './shared/shared-model-store.mjs';

const fields = new Map(
  Object.entries({
    state: [],
    language: ['uiLanguage'],
    complete: ['profileId', 'profileRevision', 'messages'],
    activate: ['expectedRevision', 'id'],
    delete: ['expectedRevision', 'id'],
    import: ['expectedRevision', 'profiles', 'activeId'],
    upsert: ['expectedRevision', 'profile', 'credential', 'activate', 'testConnection'],
  }),
);

export async function modelAction(store, action, payload, signal) {
  try {
    if (!fields.has(action) || typeof payload !== 'string' || Buffer.byteLength(payload) > 512000)
      throw new SharedModelError('request', '请求过大或操作无效。', 400);
    let data;
    try {
      data = JSON.parse(payload);
    } catch {
      throw new SharedModelError('request', '无效的模型请求。', 400);
    }
    if (
      !data ||
      typeof data !== 'object' ||
      Array.isArray(data) ||
      Object.keys(data).some((k) => !fields.get(action).includes(k))
    )
      throw new SharedModelError('request', '模型操作包含不支持的字段。', 400);
    const result = await store[action](data, signal);
    return action === 'complete'
      ? { status: 200, body: result.body, requestId: result.requestId }
      : { status: 200, body: result };
  } catch (error) {
    if (error instanceof SharedModelError)
      return {
        status: error.status,
        body: {
          error: {
            code: error.code,
            message: error.message,
            ...(storageDetailMessage(error.storageDetail) ? { detail: error.storageDetail } : {}),
          },
          ...(error.state ? { state: error.state } : {}),
        },
      };
    return {
      status: 503,
      body: {
        error: { code: 'storage', message: '无法读取当前账户的共享配置，请检查配置目录权限。' },
      },
    };
  }
}
