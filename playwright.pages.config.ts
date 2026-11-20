import { defineConfig, devices } from '@playwright/test';

export default defineConfig({
  testDir: './tests/pages-e2e',
  outputDir: 'test-results/pages',
  fullyParallel: true,
  reporter: 'list',
  use: { baseURL: 'http://127.0.0.1:34156/triobridge-ai/', trace: 'retain-on-failure' },
  projects: [
    { name: 'chromium', use: { ...devices['Desktop Chrome'] } },
    { name: 'webkit', use: { ...devices['Desktop Safari'] } },
  ],
  webServer: {
    command: 'node scripts/preview-pages.mjs',
    url: 'http://127.0.0.1:34156/triobridge-ai/app/',
    reuseExistingServer: false,
  },
});
