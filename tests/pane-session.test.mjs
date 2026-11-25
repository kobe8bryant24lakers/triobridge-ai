import assert from 'node:assert/strict';
import { test } from 'node:test';
import { PassThrough } from 'node:stream';
import { createInterface } from 'node:readline';
import { readFile } from 'node:fs/promises';
import { servePane } from '../offline/app/pane-session.mjs';

function pane(host, model) {
  const input = new PassThrough();
  const output = new PassThrough();
  const messages = createInterface({ input: output })[Symbol.asyncIterator]();
  const closed = servePane({ host, root: '.', model, input, output });
  return {
    send: (message) => input.write(JSON.stringify(message) + '\n'),
    read: async () => JSON.parse((await messages.next()).value),
    close: () => {
      input.end();
      return closed;
    },
  };
}
test('each task pane owns its Office context and write tokens', async () => {
  const first = pane('word');
  const second = pane('word');
  try {
    first.send({ id: '1', method: 'office', action: 'selection' });
    const { call, op, request } = await first.read();
    assert.equal(op, 'native');
    assert.deepEqual(request, { host: 'word', op: 'selection', kind: 'replace' });
    const guard = { document: 'A', start: 2, end: 5 };
    first.send({ reply: call, ok: true, value: { guard, text: '原文', writable: true } });
    const { token } = (await first.read()).value.value;
    const write = { id: '2', method: 'office', action: 'write', data: { token, value: '## 标题' } };
    second.send(write);
    assert.equal((await second.read()).value.ok, false);
    first.send(write);
    const writing = await first.read();
    assert.deepEqual(writing.request.guard, guard);
    assert.match(await readFile(writing.request.file, 'utf8'), /<h2\b/);
    first.send({ reply: writing.call, ok: false, message: '目标窗口已切换' });
    assert.equal((await first.read()).value.message, '目标窗口已切换');
    await assert.rejects(readFile(writing.request.file), { code: 'ENOENT' });
  } finally {
    await Promise.all([first.close(), second.close()]);
  }
});
test('pane model completion and cancellation use the shared model transport', async () => {
  let started;
  const ready = new Promise((resolve) => {
    started = resolve;
  });
  const rpc = pane('excel', async (action, payload, signal) => {
    if (action === 'state') return { status: 200, body: { profiles: [] } };
    assert.equal(payload, 'model input');
    started();
    return new Promise((_resolve, reject) =>
      signal.addEventListener('abort', () => reject(Error('取消'))),
    );
  });
  try {
    rpc.send({ id: 5, method: 'model', requestId: 'a', action: 'state', payload: '{}' });
    assert.deepEqual(await rpc.read(), {
      id: 5,
      ok: true,
      value: { status: 200, body: { profiles: [] } },
    });
    rpc.send({
      id: 6,
      method: 'model',
      requestId: 'b',
      action: 'complete',
      payload: 'model input',
    });
    await ready;
    rpc.send({ method: 'cancel', requestId: 'b' });
    assert.deepEqual(await rpc.read(), { id: 6, ok: false, message: '取消' });
    rpc.send({ id: 7, method: 'office', action: 'selection' });
    assert.equal((await rpc.read()).op, 'native');
  } finally {
    await rpc.close();
  }
});
