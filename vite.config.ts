import { defineConfig } from 'vitest/config';
import { loadEnv } from 'vite';
import react from '@vitejs/plugin-react';
import { createProxy, tlsOptions } from './scripts/server-utils.mjs';
import { resolve } from 'node:path';
import { createModelAuditMiddleware } from './scripts/model-audit';
import { createSharedModelRuntime } from './scripts/shared-model-runtime.mjs';

export default defineConfig(({ mode, command }) => {
  const env = command === 'serve' ? { ...loadEnv(mode, process.cwd(), ''), ...process.env } : {};
  const office = mode === 'office';
  return {
    plugins: [
      react(),
      {
        name: 'local-model-proxy',
        async configureServer(server) {
          if (office && ['darwin', 'win32'].includes(process.platform))
            server.middlewares.use(await createSharedModelRuntime('https://localhost:3000', env));
          if (env.VITE_REAL_MODEL_AUDIT === '1')
            server.middlewares.use(
              createModelAuditMiddleware(
                `${office ? 'https' : 'http'}://localhost:3000`,
                resolve(
                  env.REAL_MODEL_AUDIT_PATH ||
                    '.cache/real-model-acceptance-2026-11-21/requests.jsonl',
                ),
              ),
            );
          server.middlewares.use(
            createProxy(env.MODEL_PROXY_TARGET, `${office ? 'https' : 'http'}://localhost:3000`),
          );
        },
      },
    ],
    server: {
      port: 3000,
      strictPort: true,
      ...(office ? { https: tlsOptions(env) } : {}),
    },
    preview: { port: 4173, strictPort: true },
    build: { target: ['es2020', 'safari15'], sourcemap: false },
    test: {
      environment: 'jsdom',
      setupFiles: ['./tests/setup.ts'],
      include: ['tests/**/*.test.{ts,tsx}'],
      restoreMocks: true,
    },
  };
});
