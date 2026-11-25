import type { HostKind } from '../core/types';

interface ModelReply {
  status: number;
  body: unknown;
  requestId?: string;
}
export interface NativeWorkbench {
  host: HostKind;
  model(id: string, action: string, payload: string): Promise<ModelReply>;
  cancel(id: string): void;
  office(
    action: string,
    data?: unknown,
  ): Promise<{ ok: boolean; value?: unknown; message?: string }>;
}
declare global {
  interface Window {
    trioWorkbench?: NativeWorkbench;
  }
}
export function nativeWorkbench() {
  return typeof window !== 'undefined' && window.location.protocol === 'file:'
    ? window.trioWorkbench
    : undefined;
}
export const modelTransport: typeof fetch = async (url, options) => {
  const native = nativeWorkbench();
  if (!native) return fetch(url, options);
  const action =
    typeof url === 'string' &&
    /^\/local-models\/(state|upsert|import|delete|activate|complete)$/.exec(url)?.[1];
  if (!action || options?.method !== 'POST' || typeof options.body !== 'string')
    throw new Error('本地工作台无法识别这项模型操作。');
  const id = crypto.randomUUID();
  const signal = options.signal;
  signal?.throwIfAborted();
  const cancel = () => native.cancel(id);
  signal?.addEventListener('abort', cancel, { once: true });
  try {
    const result = await native.model(id, action, options.body);
    signal?.throwIfAborted();
    return new Response(JSON.stringify(result.body), {
      status: result.status,
      headers: {
        'Content-Type': 'application/json',
        ...(result.requestId ? { 'x-request-id': result.requestId } : {}),
      },
    });
  } finally {
    signal?.removeEventListener('abort', cancel);
  }
};
