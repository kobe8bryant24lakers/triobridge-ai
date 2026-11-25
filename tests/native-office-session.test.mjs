import { test } from 'node:test';
import assert from 'node:assert/strict';
import { existsSync, readFileSync } from 'node:fs';
import { runInNewContext } from 'node:vm';
import { createOfficeSession } from '../offline/app/office-session.mjs';
import { reportRtf } from '../offline/app/report.mjs';

test('Mac Word inserts into the new collapsed range and never into the original selection', () => {
  const script = readFileSync(new URL('../offline/mac/office-mac.js', import.meta.url), 'utf8');
  const word = runInNewContext(script + '\nword', { ObjC: { import() {} } });
  const originalRange = { content: '原有段落' };
  const collapsed = { startOfContent: () => 12, endOfContent: () => 12, content: '' };
  const doc = { name: () => '合成文档', fullName: () => '合成文档.docx' };
  const selection = {
    textObject: originalRange,
    content: () => originalRange.content,
    selectionStart: () => 8,
    selectionEnd: () => 12,
    storyType: () => 'main story',
  };
  const app = {
    activeDocument: doc,
    selection,
    collapseRange: (range) => {
      assert.equal(range, originalRange);
      return collapsed;
    },
  };
  const { guard } = word(app, { op: 'target', kind: 'insert' });
  word(app, { op: 'write', kind: 'insert', guard, value: '插入内容' });
  assert.equal(collapsed.content, '插入内容');
  assert.equal(originalRange.content, '原有段落');
  collapsed.startOfContent = () => 8;
  assert.throws(
    () => word(app, { op: 'write', kind: 'insert', guard, value: '错误目标' }),
    /插入位置/,
  );
  assert.equal(collapsed.content, '插入内容');
});

test('native writes are host-bound, single-use, and use the original target guard', async () => {
  const calls = [];
  const run = async (_root, q) => {
    calls.push(q);
    if (q.op === 'target')
      return {
        guard: { document: 'a', position: 8 },
        text: 'before',
        label: 'Word',
        writable: true,
      };
  };
  const session = createOfficeSession('word', '.', { run });
  assert.equal((await session.invoke('target', { kind: 'formula' })).ok, false);
  const { value } = await session.invoke('target', { kind: 'insert' });
  assert.equal(value.guard, undefined);
  assert.equal((await session.invoke('write', { token: value.token, value: 'approved' })).ok, true);
  assert.deepEqual(calls[1], {
    host: 'word',
    op: 'write',
    kind: 'insert',
    guard: { document: 'a', position: 8 },
    value: 'approved',
  });
  assert.equal((await session.invoke('write', { token: value.token, value: 'again' })).ok, false);
  session.close();
  assert.equal((await session.invoke('selection')).ok, false);
});
test('reading a structured range grants no replacement capability', async () => {
  const session = createOfficeSession('word', '.', {
    run: async () => ({ guard: {}, text: 'table', writable: false }),
  });
  const { value } = await session.invoke('selection');
  assert.equal((await session.invoke('write', { token: value.token, value: 'replace' })).ok, false);
});

for (const kind of ['replace', 'insert']) {
  test(`native Word ${kind} imports only generated Word markup at the confirmed target`, async () => {
    const guard = { document: 'synthetic.docx', start: 8, end: 12 };
    let path;
    const session = createOfficeSession('word', '.', {
      run: async (_root, request) => {
        if (request.op === 'target' || request.op === 'selection')
          return { guard, text: '旧文', writable: true };
        assert.equal(request.op, 'write');
        assert.equal(request.kind, kind);
        assert.equal(request.block, true);
        assert.deepEqual(request.guard, guard);
        assert.equal(request.value, undefined);
        path = request.file;
        const html = readFileSync(path, 'utf8');
        assert.equal(html.charCodeAt(0), 0xfeff);
        assert.match(html, /charset=["']?utf-8/i);
        assert.match(html, /<h2\b[^>]*>季度进展<\/h2>/);
        assert.match(html, /<strong\b[^>]*>已完成<\/strong>/);
        assert.doesNotMatch(html, /<img\b|## 季度进展|\*\*已完成\*\*/i);
      },
    });
    const { value: target } = await session.invoke(kind === 'replace' ? 'selection' : 'target', {
      kind,
    });
    const reply = await session.invoke('write', {
      token: target.token,
      value: '## 季度进展\n\n- **已完成**\n- ![示意图](https://example.invalid/img)',
    });
    assert.equal(reply.ok, true, reply.message);
    assert.equal(existsSync(path), false);
    session.close();
  });
}

test('native Word removes the temporary Markdown import after a rejected write', async () => {
  let path;
  const session = createOfficeSession('word', '.', {
    run: async (_root, request) => {
      if (request.op === 'selection') return { guard: {}, text: '原文', writable: true };
      path = request.file;
      assert.ok(existsSync(path));
      throw Error('选区已变化');
    },
  });
  const { value: target } = await session.invoke('selection');
  const reply = await session.invoke('write', { token: target.token, value: '# 新标题' });
  assert.deepEqual(reply, { ok: false, message: '选区已变化' });
  assert.equal(existsSync(path), false);
  session.close();
});

test('native Word sends an inline replacement without adding a paragraph', async () => {
  const session = createOfficeSession('word', '.', {
    run: async (_root, request) => {
      if (request.op === 'selection') return { guard: {}, text: '原词', writable: true };
      assert.equal(request.block, false);
      const html = readFileSync(request.file, 'utf8');
      assert.match(html, /<body><strong>新词<\/strong><\/body>/);
    },
  });
  const { value: target } = await session.invoke('selection');
  const reply = await session.invoke('write', { token: target.token, value: '**新词**' });
  assert.equal(reply.ok, true, reply.message);
  session.close();
});

function deckBytes() {
  const bytes = Buffer.alloc(22);
  bytes.writeUInt32LE(0x04034b50);
  return bytes;
}

test('native slide insertion uses a separate document target without reading a selection', async () => {
  const calls = [];
  const guard = { name: '合成演示稿.pptx', window: '123' };
  const target = { slideIds: ['256', '257'], size: { width: 960, height: 540 } };
  let importedFile;
  const session = createOfficeSession('powerpoint', '.', {
    run: async (_root, q) => {
      calls.push(q);
      if (q.op === 'presentationTarget') return { guard, target };
      assert.equal(q.op, 'append');
      assert.deepEqual(q.guard, guard);
      assert.deepEqual(q.target, target);
      assert.equal(q.snapshot, undefined);
      assert.equal(q.count, 3);
      importedFile = q.file;
      assert.deepEqual(readFileSync(q.file), deckBytes());
    },
  });
  const prepared = await session.invoke('presentationTarget');
  assert.equal(prepared.ok, true);
  assert.deepEqual(prepared.value.target, target);
  assert.equal(prepared.value.guard, undefined);
  assert.equal(
    (await session.invoke('edit', { token: prepared.value.token, proposal: {} })).ok,
    false,
  );
  const append = { token: prepared.value.token, base64: deckBytes().toString('base64'), count: 3 };
  assert.equal((await session.invoke('append', append)).ok, true);
  assert.deepEqual(
    calls.map((q) => q.op),
    ['presentationTarget', 'append'],
  );
  assert.equal(existsSync(importedFile), false);
  assert.equal((await session.invoke('append', append)).ok, false);
});

test('native slide insertion rejects edit snapshots, other hosts and invalid counts before writing', async () => {
  const calls = [];
  const run = async (_root, q) => {
    calls.push(q);
    if (q.op === 'presentation') return { guard: {}, snapshot: {} };
    if (q.op === 'presentationTarget')
      return { guard: {}, target: { slideIds: [], size: { width: 720, height: 540 } } };
    assert.fail('invalid insertion must not reach Office');
  };
  const session = createOfficeSession('powerpoint', '.', { run });
  const snapshot = await session.invoke('presentation', { scope: 'selected' });
  assert.equal(
    (
      await session.invoke('append', {
        token: snapshot.value.token,
        base64: deckBytes().toString('base64'),
        count: 1,
      })
    ).ok,
    false,
  );
  const target = await session.invoke('presentationTarget');
  for (const count of [0, -1, 1.5, undefined])
    assert.equal(
      (
        await session.invoke('append', {
          token: target.value.token,
          base64: deckBytes().toString('base64'),
          count,
        })
      ).ok,
      false,
    );
  for (const host of ['word', 'excel']) {
    const other = createOfficeSession(host, '.', { run });
    assert.equal((await other.invoke('presentationTarget')).ok, false);
  }
  assert.deepEqual(
    calls.map((q) => q.op),
    ['presentation', 'presentationTarget'],
  );
});
test('native Word reports escape markup, preserve Chinese and contain real RTF table rows', () => {
  const output = reportRtf({
    title: '指标{A}',
    sections: [
      {
        heading: '阶段',
        paragraphs: ['a\\b'],
        bullets: ['检验'],
        table: { headers: ['编号'], rows: [['42']] },
      },
    ],
  });
  assert.match(output, /\\trowd/);
  assert.ok(output.includes('\\{A\\}'));
  assert.ok(output.includes('a\\\\b'));
  assert.match(output, /\\u-?\d+\?/);
});
