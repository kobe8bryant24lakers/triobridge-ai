import assert from 'node:assert/strict';
import test from 'node:test';
import { sendModelRequest } from '../scripts/model-protocol.mjs';

const endpoint = 'http://127.0.0.1:11439/v1/messages';
const messages = [{ role: 'user', content: 'Summarize the synthetic sample.' }];
const required = { error: { message: 'max_tokens is required and must be positive' } };

test('explicit Ollama missing-and-positive requirement negotiates only the advertised output limit', async () => {
  const calls = [];
  const fetcher = async (url, options) => {
    calls.push({ url, options });
    if (calls.length === 1) return Response.json(required, { status: 400 });
    if (calls.length === 2) return Response.json({ id: 'local', max_tokens: 8192 });
    return Response.json({
      content: [{ type: 'text', text: 'complete' }],
      stop_reason: 'end_turn',
    });
  };
  const response = await sendModelRequest(fetcher, endpoint, 'anthropic', 'local', messages, '');
  assert.equal(response.status, 200);
  assert.equal(calls.length, 3);
  assert.equal(Object.hasOwn(JSON.parse(calls[0].options.body), 'max_tokens'), false);
  assert.equal(calls[1].url, 'http://127.0.0.1:11439/v1/models/local');
  assert.equal(calls[1].options.method, 'GET');
  assert.equal(JSON.parse(calls[2].options.body).max_tokens, 8192);
});

test('Ollama metadata without an output limit gives a manual OpenAI option without replaying generation', async () => {
  const calls = [];
  const fetcher = async (url, options) => {
    calls.push({ url, options });
    return calls.length === 1
      ? Response.json(required, { status: 400 })
      : Response.json({ id: 'qwen2.5:0.5b', object: 'model', created: 1, owned_by: 'library' });
  };
  await assert.rejects(
    sendModelRequest(fetcher, endpoint, 'anthropic', 'qwen2.5:0.5b', messages, ''),
    (error) => {
      assert.match(error.message, /模型信息未提供有效的 max_tokens/);
      assert.match(error.message, /同一服务提供 OpenAI 兼容接口/);
      assert.match(error.message, /手动选择/);
      return true;
    },
  );
  assert.equal(calls.length, 2);
  assert.equal(calls.filter(({ options }) => options.method === 'POST').length, 1);
  assert.equal(calls[1].url, 'http://127.0.0.1:11439/v1/models/qwen2.5%3A0.5b');
});

test('invalid-value errors do not become missing-field negotiation', async () => {
  for (const message of [
    'max_tokens is required to be positive',
    'max_tokens must be greater than 0',
    'max_tokens is required to be at most 4096',
    'max_tokens is required and must be positive; model is invalid',
  ]) {
    let calls = 0;
    const fetcher = async () => {
      calls++;
      return Response.json({ error: { message } }, { status: 400 });
    };
    const response = await sendModelRequest(fetcher, endpoint, 'anthropic', 'local', messages, '');
    assert.equal(response.status, 400, message);
    assert.equal(calls, 1, message);
  }
});
