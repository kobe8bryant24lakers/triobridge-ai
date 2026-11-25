import assert from 'node:assert/strict';
import { spawn } from 'node:child_process';
import { once } from 'node:events';
import { createInterface } from 'node:readline';

// Exercise the executable and files extracted from the ZIP. The simulated Office
// endpoint only returns synthetic text. No model configuration or key is opened.
export async function verifyPaneWorker(executable, worker) {
  const child = spawn(executable, [worker, 'word'], {
    stdio: ['pipe', 'pipe', 'inherit'],
    windowsHide: true,
    env: { ...process.env, ELECTRON_RUN_AS_NODE: '1' },
  });
  const exit = once(child, 'exit');
  const lines = createInterface({ input: child.stdout });
  const responses = lines[Symbol.asyncIterator]();
  const next = async () => {
    const { done, value } = await responses.next();
    assert.equal(done, false, 'Packaged pane worker exited without a reply.');
    return JSON.parse(value);
  };
  const send = (data) => child.stdin.write(JSON.stringify(data) + '\n');
  const timer = setTimeout(() => child.kill(), 30000);
  try {
    send({ id: 'read', method: 'office', action: 'selection', data: {} });
    const call = await next();
    assert.deepEqual(call.request, { host: 'word', op: 'selection', kind: 'replace' });
    send({
      reply: call.call,
      ok: true,
      value: { text: '原文', guard: { position: 4 }, writable: true },
    });
    const response = await next();
    assert.equal(response.id, 'read');
    assert.equal(response.value.value.text, '原文');
    assert.equal(typeof response.value.value.token, 'string');
    child.stdin.end();
    assert.equal((await exit)[0], 0);
  } finally {
    clearTimeout(timer);
    child.kill();
    lines.close();
  }
}
