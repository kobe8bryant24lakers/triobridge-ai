import { readFileSync } from 'node:fs';
import { homedir } from 'node:os';
import { join } from 'node:path';
import { Readable } from 'node:stream';
import { pipeline } from 'node:stream/promises';
import { modelHeaders } from './model-protocol.mjs';

export function tlsOptions(env = process.env) {
  const certDir = join(homedir(), '.office-addin-dev-certs');
  try {
    if (env.TLS_PFX_PATH)
      return {
        pfx: readFileSync(env.TLS_PFX_PATH),
        ...(env.TLS_PFX_PASSWORD_PATH
          ? { passphrase: readFileSync(env.TLS_PFX_PASSWORD_PATH, 'utf8').trim() }
          : {}),
      };
    return {
      key: readFileSync(env.TLS_KEY_PATH || join(certDir, 'localhost.key')),
      cert: readFileSync(env.TLS_CERT_PATH || join(certDir, 'localhost.crt')),
    };
  } catch {
    throw new Error(
      '未找到本地 HTTPS 证书。请先运行 npm run certs，或配置 TLS_KEY_PATH / TLS_CERT_PATH。',
    );
  }
}

/** Only a locally configured upstream is reachable; clients cannot supply a target URL. */
export function createProxy(target, allowedOrigin) {
  let upstream;
  if (target) {
    upstream = new URL(target);
    if (
      !['http:', 'https:'].includes(upstream.protocol) ||
      upstream.username ||
      upstream.password ||
      upstream.search ||
      upstream.hash ||
      upstream.pathname !== '/'
    )
      throw new Error('MODEL_PROXY_TARGET 必须是 HTTP(S) 服务根地址，不含路径、凭据或参数。');
    if (
      upstream.protocol === 'http:' &&
      !['localhost', '127.0.0.1', '[::1]'].includes(upstream.hostname)
    )
      throw new Error('远程代理目标必须使用 HTTPS；HTTP 仅限本机回环地址。');
  }
  return async (req, res, next) => {
    const requestPath = (req.url || '').split('?')[0];
    if (!requestPath.startsWith('/model-proxy')) {
      next();
      return;
    }
    const sendError = (status, message) => {
      res.writeHead(status, {
        'Content-Type': 'application/json',
        'Cache-Control': 'no-store',
      });
      res.end(JSON.stringify({ error: { message } }));
    };
    // Reject cross-site requests, non-JSON simple requests and DNS rebinding hosts.
    const allowed = new URL(allowedOrigin);
    // Office on Mac negotiates HTTP/2 with the HTTPS dev server. In that
    // transport the destination is :authority, and Host is normally absent.
    const authority = req.headers[':authority'];
    const requestHost = authority ?? req.headers.host;
    if (
      (authority && req.headers.host && authority !== req.headers.host) ||
      ![allowed.host, `127.0.0.1:${allowed.port}`].includes(requestHost) ||
      ![allowedOrigin, `${allowed.protocol}//127.0.0.1:${allowed.port}`].includes(
        req.headers.origin,
      )
    ) {
      sendError(403, 'Only same-origin local requests are allowed.');
      return;
    }
    if (req.method !== 'POST' || !req.headers['content-type']?.startsWith('application/json')) {
      sendError(405, 'POST application/json required.');
      return;
    }
    if (!upstream) {
      sendError(503, 'MODEL_PROXY_TARGET is not configured.');
      return;
    }
    const path = requestPath.slice('/model-proxy'.length);
    const modelId = /^\/(?:[\w-]+\/)*models\/([\w.%!~'()*-]+)$/.exec(path)?.[1];
    if (
      req.url.includes('?') ||
      (!modelId && !/^\/(?:[\w-]+\/)*(?:chat\/completions|messages)$/.test(path))
    ) {
      sendError(400, 'Unsupported model API path.');
      return;
    }
    if (modelId) {
      try {
        const id = decodeURIComponent(modelId);
        if (
          !id ||
          id.length > 200 ||
          /[\\%\u0000-\u001f]/.test(id) ||
          id.split('/').some((s) => !s || s === '.' || s === '..')
        )
          throw new Error();
      } catch {
        sendError(400, 'Invalid model identifier.');
        return;
      }
    }
    const controller = new AbortController();
    const timeout = setTimeout(() => controller.abort(), 300_000);
    const disconnected = () => {
      if (!res.writableEnded) controller.abort();
    };
    res.on('close', disconnected);
    try {
      const chunks = [];
      let size = 0;
      for await (const chunk of req) {
        size += chunk.length;
        if (size > 512_000) {
          sendError(413, 'Request too large.');
          return;
        }
        chunks.push(chunk);
      }
      const payload = Buffer.concat(chunks).toString('utf8');
      if (modelId) {
        // Keep the browser-facing POST origin check. The upstream lookup is a bodyless GET.
        if (payload) {
          sendError(400, 'Model lookup does not accept a body.');
          return;
        }
      } else {
        try {
          const parsed = JSON.parse(payload);
          if (parsed.stream !== false || !Array.isArray(parsed.messages)) throw new Error();
        } catch {
          sendError(400, 'Invalid chat request.');
          return;
        }
      }
      const response = await fetch(new URL(path, upstream), {
        method: modelId ? 'GET' : 'POST',
        signal: controller.signal,
        redirect: 'error',
        headers:
          modelId || path.endsWith('/messages')
            ? modelHeaders('anthropic', req.headers['x-api-key'])
            : {
                'Content-Type': 'application/json',
                ...(req.headers.authorization ? { Authorization: req.headers.authorization } : {}),
              },
        body: modelId ? undefined : payload,
      });
      if (!res.destroyed) {
        res.writeHead(response.status, {
          'Content-Type': 'application/json',
          'Cache-Control': 'no-store',
        });
        if (response.body) await pipeline(Readable.fromWeb(response.body), res);
        else res.end();
      }
    } catch {
      if (res.headersSent) res.destroy();
      else if (!res.destroyed)
        sendError(controller.signal.aborted ? 504 : 502, 'Model upstream unavailable.');
    } finally {
      clearTimeout(timeout);
      res.removeListener('close', disconnected);
    }
  };
}
