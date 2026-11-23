import type { IncomingMessage, ServerResponse } from 'node:http';
import type { ServerOptions } from 'node:https';
export function tlsOptions(env?: NodeJS.ProcessEnv): ServerOptions;
export function createProxy(
  target: string | undefined,
  allowedOrigin: string,
): (req: IncomingMessage, res: ServerResponse, next: () => void) => Promise<void>;
