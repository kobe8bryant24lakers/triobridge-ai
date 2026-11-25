import { test } from 'node:test';
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
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
