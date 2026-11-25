import { request } from 'node:https';
import { createServer } from 'node:net';
import { setTimeout as delay } from 'node:timers/promises';
import { storageDetailHeader, storageDetailMessage } from './model-storage-diagnostic.mjs';

const messages = {
  PORT_BUSY: '3000 端口仍被占用。请检查占用程序，释放端口后再双击 Install。',
  PORT_DENIED: '系统拒绝监听本机 3000 端口。请让管理员检查该端口的保留范围及本机策略。',
  NO_LISTENER: '等待本机 HTTPS 服务监听超时。请检查后台服务是否已运行。',
  TIMEOUT: '本机服务未在限定时间内完成响应。',
  CONNECTION_RESET: '本机服务在响应完成前断开连接。',
  TLS_CERTIFICATE: '本机 HTTPS 证书校验失败。请重新运行 Install 检查证书及端口归属。',
  SERVICE_IDENTITY: '端口上的服务与本次安装的标识或版本不匹配。',
  LOCAL_AUTH: '共享配置接口拒绝了本次安装的凭据。请使用同一 Windows 或 Mac 账号运行安装。',
  MODEL_STORAGE: '本机服务已启动，但共享配置读取失败。',
  INVALID_RESPONSE: '本机服务返回了无法识别的响应。',
  TRANSPORT: '无法通过本机回环地址建立 HTTPS 连接。',
};

export class ReadinessError extends Error {
  constructor(code, detail) {
    const message = code === 'MODEL_STORAGE' && storageDetailMessage(detail);
    super(message ? messages[code] + ' [' + detail + '] ' + message : messages[code]);
    this.code = code;
    if (message) this.storageDetail = detail;
  }
}

// A reserved Windows port can refuse a bind even when a TCP connection sees no listener.
export async function checkLoopbackPort(port = 3000) {
  const server = createServer((socket) => socket.destroy());
  return new Promise((done, reject) => {
    server.once('error', (error) => {
      const code = { EADDRINUSE: 'PORT_BUSY', EACCES: 'PORT_DENIED', EPERM: 'PORT_DENIED' };
      reject(new ReadinessError(code[error.code] || 'TRANSPORT'));
    });
    server.listen({ host: '127.0.0.1', port, exclusive: true }, () => server.close(done));
  });
}

function transportError(error) {
  if (error instanceof ReadinessError) return error;
  if (error.code === 'ECONNREFUSED') return new ReadinessError('NO_LISTENER');
  if (error.code === 'ECONNRESET' || error.code === 'EPIPE')
    return new ReadinessError('CONNECTION_RESET');
  if (error.code === 'ETIMEDOUT') return new ReadinessError('TIMEOUT');
  if (/CERT|TLS|SSL|SELF_SIGNED|ISSUER|VERIFY_LEAF/.test(error.code || ''))
    return new ReadinessError('TLS_CERTIFICATE');
  return new ReadinessError('TRANSPORT');
}

function probe({ ca, port, token }, timeout) {
  let timer;
  return new Promise((done, reject) => {
    const req = request(
      {
        hostname: '127.0.0.1',
        servername: 'localhost',
        port,
        ca,
        rejectUnauthorized: true,
        // The private installation credential must only travel over a direct loopback connection.
        agent: false,
        path: token ? '/local-models/state' : '/local-health',
        method: token ? 'POST' : 'GET',
        headers: {
          Host: 'localhost:' + port,
          ...(token && {
            Origin: 'https://localhost:' + port,
            'Content-Type': 'application/json',
            'X-TrioBridge-Local-Token': token,
          }),
        },
      },
      (res) => {
        res.on('error', reject);
        if (res.statusCode !== 200) {
          let code = 'INVALID_RESPONSE';
          if (token && [401, 403].includes(res.statusCode)) code = 'LOCAL_AUTH';
          if (token && res.statusCode === 503) code = 'MODEL_STORAGE';
          reject(new ReadinessError(code, res.headers[storageDetailHeader]));
          res.destroy();
          return;
        }
        const chunks = [];
        let size = 0;
        res.on('data', (chunk) => {
          size += chunk.length;
          if (size > 512_000) req.destroy(new ReadinessError('INVALID_RESPONSE'));
          else chunks.push(chunk);
        });
        res.on('end', () => {
          try {
            done(JSON.parse(Buffer.concat(chunks).toString('utf8')));
          } catch {
            reject(new ReadinessError('INVALID_RESPONSE'));
          }
        });
      },
    );
    req.on('error', reject);
    // Bound the entire TLS exchange, including a peer that keeps sending partial data.
    timer = setTimeout(() => req.destroy(new ReadinessError('TIMEOUT')), timeout);
    req.end(token ? '{}' : undefined);
  }).finally(() => clearTimeout(timer));
}

export async function waitForLocalEndpoint({
  ca,
  installationId,
  version,
  token,
  port = 3000,
  timeout = process.platform === 'win32' ? 120_000 : 30_000,
}) {
  const deadline = Date.now() + timeout;
  let retryDelay = 300;
  while (true) {
    try {
      const remaining = Math.max(1, deadline - Date.now());
      const data = await probe({ ca, port, token }, token ? remaining : Math.min(remaining, 5000));
      if (token) {
        if (
          data?.version !== 2 ||
          !Number.isSafeInteger(data.revision) ||
          !Array.isArray(data.profiles)
        )
          throw new ReadinessError('INVALID_RESPONSE');
      } else if (
        data?.product !== 'triobridge-ai-local' ||
        data.installationId !== installationId ||
        (version && data.version !== version)
      ) {
        throw new ReadinessError('SERVICE_IDENTITY');
      }
      return;
    } catch (failure) {
      const error = transportError(failure);
      const remaining = deadline - Date.now();
      const retry =
        ['NO_LISTENER', 'CONNECTION_RESET', 'TIMEOUT'].includes(error.code) ||
        error.storageDetail === 'CATALOG_BUSY';
      if (!retry || remaining <= 0) throw error;
      await delay(Math.min(retryDelay, remaining));
      if (Date.now() >= deadline) throw error;
      retryDelay = Math.min(2000, retryDelay * 2);
    }
  }
}
