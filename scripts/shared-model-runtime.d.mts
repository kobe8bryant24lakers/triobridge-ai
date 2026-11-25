import type { IncomingMessage, ServerResponse } from 'node:http';
export function createSharedModelRuntime(
  allowedOrigin: string,
  env?: NodeJS.ProcessEnv,
  options?: { directory?: string },
): Promise<(req: IncomingMessage, res: ServerResponse, next: () => void) => Promise<void> | void>;
