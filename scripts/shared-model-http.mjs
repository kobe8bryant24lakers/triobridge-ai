import { acceptsLocalOrigin } from './local-request-guard.mjs';
import { SharedModelError } from './shared-model-store.mjs';
import { matchesLocalCapability } from './local-capability.mjs';
import { storageDetailHeader, storageDetailMessage } from './model-storage-diagnostic.mjs';

const fields = {
  state: [],
  upsert: ['expectedRevision', 'profile', 'credential', 'activate', 'testConnection'],
  import: ['expectedRevision', 'profiles', 'activeId'],
  delete: ['expectedRevision', 'id'],
  activate: ['expectedRevision', 'id'],
  complete: ['profileId', 'profileRevision', 'messages'],
};

export function createSharedModelMiddleware({ store, allowedOrigin, authToken }) {
  const origin = new URL(allowedOrigin);
  if (
    origin.protocol !== 'https:' ||
    !['localhost', '127.0.0.1'].includes(origin.hostname) ||
    origin.origin !== allowedOrigin
  )
    throw new Error('Shared models require an exact HTTPS loopback origin.');
  if (typeof authToken !== 'string' || !/^[a-f0-9]{64}$/.test(authToken))
    throw new Error('Shared models require a per-install local capability.');
  return async (req, res, next) => {
    if (!req.url?.startsWith('/local-models/')) return next();
    const send = (status, value, requestId) => {
      if (res.destroyed) return;
      res.statusCode = status;
      res.setHeader('Content-Type', 'application/json');
      res.setHeader('Cache-Control', 'no-store');
      res.setHeader('X-Content-Type-Options', 'nosniff');
      if (requestId) res.setHeader('x-request-id', requestId);
      res.end(JSON.stringify(value));
    };
    const error = (status, code, message) => send(status, { error: { code, message } });
    if (!acceptsLocalOrigin(req, allowedOrigin))
      return error(403, 'origin', '只允许同源的本机加载项访问共享模型。');
    if (!matchesLocalCapability(req.headers['x-triobridge-local-token'], authToken))
      return error(401, 'local_auth', '此加载项尚未连接本机共享服务，请使用本机安装入口重新打开。');
    if (req.method !== 'POST' || !req.headers['content-type']?.startsWith('application/json'))
      return error(405, 'method', '共享模型仅接受 JSON POST。');
    const action = req.url.slice('/local-models/'.length);
    if (!Object.hasOwn(fields, action)) return error(404, 'path', '共享模型接口不存在。');
    const controller = new AbortController();
    const disconnect = () => {
      if (!res.writableEnded) controller.abort();
    };
    res.on('close', disconnect);
    try {
      const chunks = [];
      let size = 0;
      for await (const chunk of req) {
        size += chunk.length;
        if (size > 512_000) return error(413, 'size', '共享模型请求过大。');
        chunks.push(Buffer.from(chunk));
      }
      let input;
      try {
        input = JSON.parse(Buffer.concat(chunks).toString('utf8'));
      } catch {
        return error(400, 'request', '请求需要有效的 JSON。');
      }
      if (
        !input ||
        typeof input !== 'object' ||
        Array.isArray(input) ||
        Object.keys(input).some((key) => !fields[action].includes(key))
      )
        return error(400, 'request', '请求包含不支持的字段。');
      const result = await store[action](input, controller.signal);
      if (action === 'complete') send(200, result.body, result.requestId);
      else send(200, result);
    } catch (failure) {
      if (action === 'state' && storageDetailMessage(failure?.storageDetail))
        res.setHeader(storageDetailHeader, failure.storageDetail);
      if (failure instanceof SharedModelError)
        send(failure.status, {
          error: { code: failure.code, message: failure.message },
          ...(failure.state ? { state: failure.state } : {}),
        });
      else error(503, 'storage', '本机共享模型服务暂时不可用，未降级保存密钥。');
    } finally {
      res.removeListener('close', disconnect);
    }
  };
}
