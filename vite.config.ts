import { defineConfig } from 'vitest/config';
import { loadEnv } from 'vite';
import react from '@vitejs/plugin-react';
import { createProxy, tlsOptions } from './scripts/server-utils.mjs';

export default defineConfig(({ mode, command }) => {
  const env = command === 'serve' ? { ...loadEnv(mode, process.cwd(), ''), ...process.env } : {};
  const office = mode === 'office';
  return {
    plugins: [
      react(),
      {
        name: 'local-model-proxy',
        configureServer(server) {
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
