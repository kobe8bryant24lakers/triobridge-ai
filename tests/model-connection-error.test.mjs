import { test } from 'node:test';
import assert from 'node:assert/strict';
import { modelConnectionError } from '../scripts/model-connection-error.mjs';

test('transport diagnostics explain known causes without returning raw error contents', () => {
  const endpoint = 'http://127.0.0.1:8001/private-path/chat/completions';
  for (const [code, expected] of [
    ['ECONNREFUSED', '拒绝连接'],
    ['ENOTFOUND', '解析'],
    ['EAI_AGAIN', 'DNS'],
    ['ETIMEDOUT', '超时'],
    ['UND_ERR_CONNECT_TIMEOUT', '超时'],
    ['ECONNRESET', '中断'],
    ['UND_ERR_SOCKET', '中断'],
    ['DEPTH_ZERO_SELF_SIGNED_CERT', '证书'],
    ['ERR_TLS_CERT_ALTNAME_INVALID', '证书'],
    ['ERR_SSL_WRONG_VERSION_NUMBER', 'HTTP 响应'],
    ['EHOSTUNREACH', '网络'],
    ['unknown-private-code', '请检查 API'],
  ]) {
    const error = new TypeError('private-message', {
      cause: new AggregateError([Object.assign(new Error('private-cause'), { code })]),
    });
    const message = modelConnectionError(error, endpoint);
    assert.match(message, new RegExp(expected));
    assert.match(message, /127\.0\.0\.1:8001/);
    assert.doesNotMatch(message, /private|CORS/);
  }
});

test('cyclic error causes are bounded and remote errors do not suggest loopback port mapping', () => {
  const error = { code: 'ECONNREFUSED' };
  error.cause = error;
  assert.doesNotMatch(modelConnectionError(error, 'https://api.example.com/v1'), /容器|回环/);
});
