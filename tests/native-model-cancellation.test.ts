import { afterEach, expect, it, vi } from 'vitest';
import { modelTransport } from '../src/offline/transport';

afterEach(() => vi.unstubAllGlobals());

it('releases a cancelled native completion while a credential dialog still owns the request', async () => {
  let finish!: (reply: { status: number; body: string }) => void;
  const native = {
    model: vi.fn(
      () =>
        new Promise<{ status: number; body: string }>((resolve) => {
          finish = resolve;
        }),
    ),
    cancel: vi.fn(),
  };
  const workbenchWindow = {
    get top() {
      return this;
    },
    location: { protocol: 'file:', href: 'file:///workbench/offline.html' },
    trioWorkbench: native,
  };
  vi.stubGlobal('window', workbenchWindow);
  const controller = new AbortController();
  const pending = modelTransport('/local-models/complete', {
    method: 'POST',
    body: '{}',
    signal: controller.signal,
  });
  controller.abort();
  await expect(pending).rejects.toMatchObject({ name: 'AbortError' });
  const [id] = native.cancel.mock.calls[0];
  expect(native.model).toHaveBeenCalledWith(id, 'complete', '{}');
  finish({ status: 200, body: 'late completion' });
});
