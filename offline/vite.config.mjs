import { defineConfig } from 'vite';
import react from '@vitejs/plugin-react';
export default defineConfig({
  base: './',
  publicDir: false,
  plugins: [react()],
  build: {
    outDir: 'build/native-ui',
    target: 'es2022',
    license: { fileName: 'THIRD-PARTY-NOTICES.md' },
    rollupOptions: { input: 'offline.html' },
  },
});
