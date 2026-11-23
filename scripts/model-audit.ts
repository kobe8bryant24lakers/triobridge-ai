import { appendFile, mkdir } from 'node:fs/promises';
import type { IncomingMessage, ServerResponse } from 'node:http';
import type { Http2ServerRequest, Http2ServerResponse } from 'node:http2';
import { dirname } from 'node:path';
import { MODEL_AUDIT_PATH, normalizeAuditRecord } from '../src/core/model-audit';

/** Local dev server only. No target/file path is accepted from the client. */
export function createModelAuditMiddleware(allowedOrigin: string, outputFile: string) {
  const allowed = new URL(allowedOrigin);
  if (!['localhost', '127.0.0.1'].includes(allowed.hostname))
    throw new Error('Model audit requires a loopback origin.');
  return async (
    req: IncomingMessage | Http2ServerRequest,
    res: ServerResponse | Http2ServerResponse,
    next: () => void,
  ) => {
    if (!req.url?.startsWith(MODEL_AUDIT_PATH)) return next();
    const end = (status: number) => {
      res.statusCode = status;
      res.setHeader('Cache-Control', 'no-store');
      res.end();
    };
    if (req.url !== MODEL_AUDIT_PATH) return end(400);
    const authority = req.headers[':authority'];
    const host = authority ?? req.headers.host;
    if (
      (authority && req.headers.host && authority !== req.headers.host) ||
      ![allowed.host, `127.0.0.1:${allowed.port}`].includes(String(host)) ||
      ![allowedOrigin, `${allowed.protocol}//127.0.0.1:${allowed.port}`].includes(
        String(req.headers.origin),
      )
    )
      return end(403);
    if (req.method !== 'POST' || !req.headers['content-type']?.startsWith('application/json'))
      return end(405);
    try {
      const chunks: Buffer[] = [];
      let size = 0;
      for await (const chunk of req) {
        size += chunk.length;
        if (size > 512_000) return end(413);
        chunks.push(Buffer.from(chunk));
      }
      let record;
      try {
        record = normalizeAuditRecord(JSON.parse(Buffer.concat(chunks).toString('utf8')));
      } catch {
        return end(400);
      }
      if (!record) return end(400);
      await mkdir(dirname(outputFile), { recursive: true, mode: 0o700 });
      await appendFile(outputFile, `${JSON.stringify(record)}\n`, { mode: 0o600 });
      end(204);
    } catch {
      end(500);
    }
  };
}
