import { getOrCreateLocalCapability } from './local-capability.mjs';
import { createSharedModelStore, sharedModelsDirectory } from './shared-model-store.mjs';
import { createSharedModelMiddleware } from './shared-model-http.mjs';

/** Only the installed HTTPS loopback app can use the per-user shared catalog. */
export async function createSharedModelRuntime(allowedOrigin, env = process.env) {
  if (env.TRIOBRIDGE_SHARED_MODELS === '0') return (_req, _res, next) => next();
  try {
    const directory = sharedModelsDirectory();
    const authToken = await getOrCreateLocalCapability(directory);
    const store = createSharedModelStore({ directory, proxyTarget: env.MODEL_PROXY_TARGET });
    await store.state();
    await store.cleanup().catch(() => {
      console.warn('Retired local model credentials will be cleaned up on a later start.');
    });
    return createSharedModelMiddleware({ store, allowedOrigin, authToken });
  } catch {
    // Keep the UI available to explain recovery; never substitute an empty catalog.
    console.warn(
      'Local shared model storage is unavailable. Existing configuration was not replaced.',
    );
    return (req, res, next) => {
      if (!req.url?.startsWith('/local-models/')) return next();
      res.writeHead(503, { 'Content-Type': 'application/json', 'Cache-Control': 'no-store' });
      res.end(
        JSON.stringify({
          error: {
            code: 'storage',
            message: '本机共享配置暂时不可用，请检查本机服务与安装。原配置未被覆盖。',
          },
        }),
      );
    };
  }
}
