import type { IncomingMessage, ServerResponse } from 'node:http';
export function tlsOptions(env?: NodeJS.ProcessEnv): {
  key: Buffer;
  cert: Buffer;
};
export function createProxy(
  target: string | undefined,
  allowedOrigin: string,
): (req: IncomingMessage, res: ServerResponse, next: () => void) => Promise<void>;
