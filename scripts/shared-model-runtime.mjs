import { getOrCreateLocalCapability } from './local-capability.mjs';
import { acceptsLocalOrigin } from './local-request-guard.mjs';
import { storageDetailHeader, storageFailureDetail } from './model-storage-diagnostic.mjs';
import { createSharedModelStore, sharedModelsDirectory } from './shared-model-store.mjs';
import { createSharedModelMiddleware } from './shared-model-http.mjs';

/** Only the installed HTTPS loopback app can use the per-user shared catalog. */
export async function createSharedModelRuntime(allowedOrigin, env = process.env, options = {}) {
  if (env.TRIOBRIDGE_SHARED_MODELS === '0') return (_req, _res, next) => next();
  let ready;
  let retryAt = 0;
  let lastDetail;
  async function initialize() {
    let fallback = 'CAPABILITY_ACCESS';
    try {
      const directory = options.directory ?? sharedModelsDirectory();
      const authToken = await getOrCreateLocalCapability(directory);
      fallback = 'STORE_INIT';
      const store = createSharedModelStore({ directory, proxyTarget: env.MODEL_PROXY_TARGET });
      const middleware = createSharedModelMiddleware({ store, allowedOrigin, authToken });
      // Each authenticated state request reads the catalog again. A failed first
      // read must never replace the middleware for the lifetime of the process.
      await store.cleanup().catch(() => {
        console.warn('Shared model maintenance will be retried on a later start.');
      });
      return middleware;
    } catch (error) {
      lastDetail = storageFailureDetail(error, fallback);
      console.warn('Local shared model initialization failed [' + lastDetail + '].');
      retryAt = Date.now() + 1000;
      ready = undefined;
      return null;
    }
  }
  function getReady() {
    if (!ready && Date.now() >= retryAt) ready = Promise.resolve().then(initialize);
    return ready;
  }
  getReady();
  return async (req, res, next) => {
    if (!req.url?.startsWith('/local-models/')) return next();
    if (!acceptsLocalOrigin(req, allowedOrigin)) {
      res.writeHead(403, { 'Content-Type': 'application/json', 'Cache-Control': 'no-store' });
      return res.end(
        JSON.stringify({
          error: { code: 'origin', message: '只允许同源的本机加载项访问共享模型。' },
        }),
      );
    }
    const middleware = await getReady();
    if (middleware) return middleware(req, res, next);
    const headers = { 'Content-Type': 'application/json', 'Cache-Control': 'no-store' };
    if (
      req.url === '/local-models/state' &&
      req.method === 'POST' &&
      /^[a-f0-9]{64}$/.test(req.headers['x-triobridge-local-token'] ?? '')
    )
      headers[storageDetailHeader] = lastDetail;
    res.writeHead(503, headers);
    return res.end(
      JSON.stringify({
        error: { code: 'storage', message: '本机共享配置暂时不可用，请检查本机服务与安装。' },
      }),
    );
  };
}
