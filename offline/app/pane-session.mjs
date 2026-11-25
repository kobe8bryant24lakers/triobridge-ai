import { randomUUID } from 'node:crypto';
import { createInterface } from 'node:readline';
import { createOfficeSession } from './office-session.mjs';

// The installed Office control owns both ends of this worker's redirected stdio.
// Each control has a dedicated process, cancellation registry and write-target ledger.
export async function servePane({ host, root, input, output, model }) {
  if (!['word', 'excel', 'powerpoint'].includes(host)) throw Error('无效的 Office 应用。');
  const lines = createInterface({ input, crlfDelay: Infinity, terminal: false });
  const calls = new Map();
  const controllers = new Map();
  const pending = new Set();
  let closed = false;
  function send(message) {
    if (closed) throw Error('侧边栏已关闭。');
    output.write(JSON.stringify(message) + '\n');
  }
  function callHost(op, request) {
    const call = randomUUID();
    return new Promise((resolve, reject) => {
      send({ call, op, request });
      calls.set(call, { resolve, reject });
    });
  }
  const office = createOfficeSession(host, root, {
    run: (_root, request) => callHost('native', request),
    saveDeck: () => callHost('saveDeck'),
  });
  async function dispatch(msg) {
    if (msg.method === 'office') return office.invoke(msg.action, msg.data);
    if (msg.method !== 'model' || typeof msg.requestId !== 'string')
      throw Error('未知的侧边栏请求。');
    if (controllers.has(msg.requestId) || controllers.size >= 12) throw Error('请求繁忙，请重试。');
    const controller = new AbortController();
    controllers.set(msg.requestId, controller);
    try {
      return await model(msg.action, msg.payload, controller.signal);
    } finally {
      controllers.delete(msg.requestId);
    }
  }
  async function respond(msg) {
    try {
      const value = await dispatch(msg);
      if (!closed) send({ id: msg.id, ok: true, value });
    } catch (error) {
      if (!closed) send({ id: msg.id, ok: false, message: String(error.message) });
    }
  }
  function disconnect() {
    if (closed) return;
    closed = true;
    office.close();
    for (const controller of controllers.values()) controller.abort();
    for (const call of calls.values()) call.reject(Error('侧边栏已关闭。'));
    calls.clear();
    lines.close();
  }
  input.on('error', disconnect);
  output.on('error', disconnect);
  try {
    for await (const line of lines) {
      if (closed) break;
      const msg = JSON.parse(line);
      if (typeof msg.reply === 'string') {
        const call = calls.get(msg.reply);
        if (!call) continue;
        calls.delete(msg.reply);
        if (msg.ok === true) call.resolve(msg.value);
        else call.reject(Error(msg.message || 'Office 未能完成该操作。'));
      } else if (msg.method === 'cancel') {
        controllers.get(msg.requestId)?.abort();
      } else if (typeof msg.id === 'string' || Number.isSafeInteger(msg.id)) {
        const task = respond(msg);
        pending.add(task);
        task.then(() => pending.delete(task), disconnect);
      }
    }
  } finally {
    disconnect();
    await Promise.allSettled(pending);
    input.off('error', disconnect);
    output.off('error', disconnect);
  }
}
