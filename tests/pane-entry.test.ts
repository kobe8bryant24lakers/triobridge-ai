import { afterEach, expect, it, vi } from 'vitest';
import { nativeWorkbench } from '../src/offline/transport';

afterEach(() => vi.unstubAllGlobals());

it.each([
  ['https://triobridge.local/offline.html', true, true],
  ['https://triobridge.local/offline.html#main', true, true],
  ['file:///applications/triobridge/ui/offline.html', true, true],
  ['https://triobridge.local/offline.html', false, false],
  ['https://triobridge.local/other.html', true, false],
  ['https://triobridge.local/offline.html?remote=1', true, false],
  ['https://example.test/offline.html', true, false],
] as const)('checks the native workbench entry for %s, top frame %s', (url, topFrame, allowed) => {
  const bridge = { host: 'word' };
  const surface = { location: new URL(url), trioWorkbench: bridge, top: {} };
  if (topFrame) surface.top = surface;
  vi.stubGlobal('window', surface);
  expect(nativeWorkbench()).toBe(allowed ? bridge : undefined);
});
