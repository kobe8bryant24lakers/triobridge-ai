import { test } from 'node:test';
import assert from 'node:assert/strict';
import { mkdtemp, rm } from 'node:fs/promises';
import { join } from 'node:path';
import { tmpdir } from 'node:os';
import { createOfficeSession } from '../offline/app/office-session.mjs';
import { createSharedModelStore } from '../scripts/shared-model-store.mjs';

test('shared model generation preserves a complete response above one megabyte without retrying', async (t) => {
  const directory = await mkdtemp(join(tmpdir(), 'triobridge-long-model-'));
  t.after(() => rm(directory, { recursive: true, force: true }));
  const content = '中'.repeat(400_000) + '完整结果末尾';
  let calls = 0;
  const store = createSharedModelStore({
    directory,
    vault: {
      kind: 'macos-keychain',
      get: async () => null,
      set: async () => {},
      delete: async () => {},
    },
    fetcher: async () => {
      calls++;
      return Response.json({ choices: [{ finish_reason: 'stop', message: { content } }] });
    },
  });
  const state = await store.upsert({
    expectedRevision: 0,
    profile: {
      id: 'long-text',
      name: '长文本',
      model: 'local-model',
      baseUrl: 'http://localhost:8001/v1',
      timeoutSeconds: 5,
    },
    credential: { action: 'remove' },
  });
  const result = await store.complete({
    profileId: state.profiles[0].id,
    profileRevision: state.profiles[0].revision,
    messages: [{ role: 'user', content: '改写原文' }],
  });
  assert.equal(result.body.choices[0].message.content, content);
  assert.equal(calls, 1);
});

for (const kind of ['replace', 'insert'])
  test(`native Word ${kind} passes complete plain text beyond the old JSON write limit`, async () => {
    const value = '中文𠮷<纯文本>\n'.repeat(220_000) + '末尾';
    assert.ok(JSON.stringify(value).length > 2_000_000);
    const guard = { document: 'original.docx', start: 0, end: 4, text: '原文内容' };
    let writes = 0;
    const session = createOfficeSession('word', '.', {
      run: async (_root, q) => {
        if (q.op === 'selection' || q.op === 'target')
          return { guard, text: '原文内容', label: 'Word', writable: true };
        writes++;
        assert.equal(q.kind, kind);
        assert.equal(q.value, value);
        assert.deepEqual(q.guard, guard);
        assert.equal(q.file, undefined);
        assert.equal(q.textFile, undefined);
      },
    });
    const prepared = await session.invoke(kind === 'replace' ? 'selection' : 'target', { kind });
    assert.equal(prepared.ok, true);
    const token = prepared.value.token;
    assert.equal((await session.invoke('write', { token, value })).ok, true);
    assert.equal(writes, 1);
    assert.equal((await session.invoke('write', { token, value })).ok, false);
    session.close();
  });

test('native Word plain-text writes still reject non-string payloads before calling Office', async () => {
  let writes = 0;
  const session = createOfficeSession('word', '.', {
    run: async (_root, q) => {
      if (q.op === 'target') return { guard: {}, text: '', writable: true };
      writes++;
    },
  });
  const { value: target } = await session.invoke('target', { kind: 'insert' });
  for (const value of [null, {}, ['text'], 42, true])
    assert.equal((await session.invoke('write', { token: target.token, value })).ok, false);
  assert.equal(writes, 0);
  session.close();
});
