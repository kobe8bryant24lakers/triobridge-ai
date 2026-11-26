import { afterEach, describe, expect, it } from 'vitest';
import {
  appendEditingTurn,
  buildEditingFollowUp,
  type EditingTurn,
} from '../src/core/editing-conversation';
import type { ChatMessage } from '../src/core/types';
import { setUiLanguage } from '../src/i18n';

const base: ChatMessage[] = [
  { role: 'system', content: 'Preserve the supplied facts. Return only JSON: {"title":"..."}.' },
  { role: 'user', content: '{"requirements":"Generate a report","documentContent":"Revenue: 42"}' },
];
const history: EditingTurn[] = [
  { instruction: 'Generate a report', result: '{"title":"First draft"}' },
  { instruction: 'Make the title shorter', result: '{"title":"Short title"}' },
  { instruction: 'Use Chinese', result: '{"title":"中文标题"}' },
];

function requestTextLength(messages: ChatMessage[]) {
  return messages.reduce(
    (length, message) =>
      length +
      (typeof message.content === 'string'
        ? message.content.length
        : message.content.reduce(
            (total, part) => total + (part.type === 'text' ? part.text.length : 0),
            0,
          )),
    0,
  );
}

afterEach(() => setUiLanguage('zh-CN'));

describe('multi-turn document editing context', () => {
  it('keeps manual edits in earlier round baselines when saving subsequent successful revisions', () => {
    const before = JSON.stringify(history);
    const updated = appendEditingTurn(
      history,
      'Manually edited third draft',
      '  More concise  ',
      'Fourth draft',
    );
    expect(updated).toEqual([
      history[0],
      history[1],
      { instruction: history[2].instruction, result: 'Manually edited third draft' },
      { instruction: 'More concise', result: 'Fourth draft' },
    ]);
    expect(JSON.stringify(history)).toBe(before);
    expect(
      buildEditingFollowUp(base, updated, 'Fourth draft', 'Next').map((message) => message.content),
    ).toContain('Manually edited third draft');
    expect(appendEditingTurn([], 'Kept after clear', 'Next', 'New result')).toEqual([
      { instruction: '', result: 'Kept after clear' },
      { instruction: 'Next', result: 'New result' },
    ]);
  });

  it('preserves the original task, sources, and successive successful requests and results', () => {
    const messages = buildEditingFollowUp(
      base,
      history,
      history[2].result,
      'Make it more specific',
    );
    expect(messages[0].content).toContain(base[0].content);
    expect(messages[0].content).toContain('完整新版结果');
    expect(messages[0].content).toContain('不得声称已执行写入');
    expect(messages.slice(1)).toEqual([
      base[1],
      { role: 'assistant', content: history[0].result },
      { role: 'user', content: history[1].instruction },
      { role: 'assistant', content: history[1].result },
      { role: 'user', content: history[2].instruction },
      { role: 'assistant', content: history[2].result },
      { role: 'user', content: 'Make it more specific' },
    ]);
  });

  it('uses the manually edited current preview instead of the last model response', () => {
    const manual = '{"title":"用户手动保留的标题"}';
    const before = JSON.stringify({ base, history });
    const messages = buildEditingFollowUp(base, history, manual, '保留标题，继续完善');
    expect(messages[messages.length - 2]).toEqual({ role: 'assistant', content: manual });
    expect(messages.some((message) => message.content === history[2].result)).toBe(false);
    expect(JSON.stringify({ base, history })).toBe(before);
  });

  it('can start from a current preview after clearing history or making a manual proposal', () => {
    const messages = buildEditingFollowUp(base, [], 'Current preview', '  Keep the facts  ');
    expect(messages.slice(1)).toEqual([
      base[1],
      { role: 'assistant', content: 'Current preview' },
      { role: 'user', content: 'Keep the facts' },
    ]);
  });

  it('retains selected images once without counting image data as text or mutating the source', () => {
    const image = {
      type: 'image_url' as const,
      image_url: { url: 'data:image/png;base64,' + 'A'.repeat(49_000), detail: 'low' as const },
    };
    const multimodal: ChatMessage[] = [
      base[0],
      { role: 'user', content: [{ type: 'text', text: 'Use only this source image.' }, image] },
    ];
    const messages = buildEditingFollowUp(
      multimodal,
      history,
      'A'.repeat(110_000),
      'Make it shorter',
    );
    expect(messages[1]).toEqual(multimodal[1]);
    expect(messages[1]).not.toBe(multimodal[1]);
    expect(messages[1].content).not.toBe(multimodal[1].content);
    expect(JSON.stringify(messages).match(/data:image\/png/g)).toHaveLength(1);
    expect(requestTextLength(messages)).toBeLessThanOrEqual(150_000);
  });

  it('drops oldest whole rounds to fit the shared service limit while retaining original source', () => {
    const many = Array.from({ length: 20 }, (_, index) => ({
      instruction: `Request ${index}`,
      result: `Result ${index}`,
    }));
    const messages = buildEditingFollowUp(base, many, 'Current manual draft', 'Next request');
    expect(messages).toHaveLength(16);
    expect(messages[1]).toEqual(base[1]);
    expect(messages[2]).toEqual({ role: 'assistant', content: 'Result 13' });
    expect(messages[3]).toEqual({ role: 'user', content: 'Request 14' });
    expect(messages[messages.length - 2]).toEqual({
      role: 'assistant',
      content: 'Current manual draft',
    });
    expect(messages[messages.length - 1]).toEqual({ role: 'user', content: 'Next request' });
    messages
      .slice(1)
      .forEach((message, index) => expect(message.role).toBe(index % 2 ? 'assistant' : 'user'));
  });

  it('keeps JSON results whole when dropping earlier history to meet the text budget', () => {
    const long = [
      {
        instruction: 'Initial',
        result: JSON.stringify({ title: 'Old', body: 'A'.repeat(90_000) }),
      },
      {
        instruction: 'Refine',
        result: JSON.stringify({ title: 'Recent', body: 'B'.repeat(70_000) }),
      },
      {
        instruction: 'Refine again',
        result: JSON.stringify({ title: 'Newest', body: 'C'.repeat(70_000) }),
      },
    ];
    const messages = buildEditingFollowUp(base, long, long[2].result, 'Keep it complete');
    expect(messages).toHaveLength(6);
    expect(messages[2]).toEqual({ role: 'assistant', content: long[1].result });
    expect(messages[4]).toEqual({ role: 'assistant', content: long[2].result });
    expect(requestTextLength(messages)).toBeLessThanOrEqual(150_000);
    for (const message of messages.filter((entry) => entry.role === 'assistant'))
      expect(() => JSON.parse(message.content as string)).not.toThrow();
  });

  it('rejects oversized required context instead of silently truncating sources or current output', () => {
    const source = { ...base[1], content: 'S'.repeat(60_000) };
    const draft = 'D'.repeat(100_000);
    expect(() => buildEditingFollowUp([base[0], source], history, draft, 'Shorten')).toThrow(
      '内容已完整保留',
    );
    expect(source.content).toHaveLength(60_000);
    expect(draft).toHaveLength(100_000);
  });

  it('keeps document text beyond the 100,000-character output convention when the request fits', () => {
    const draft = 'Text'.repeat(27_000);
    const messages = buildEditingFollowUp(base, [], draft, 'Improve the wording');
    expect(messages[2].content).toBe(draft);
    expect(requestTextLength(messages)).toBeLessThanOrEqual(150_000);
  });

  it.each([
    { instruction: '', expected: '请填写继续调整' },
    { instruction: ' '.repeat(12), expected: '请填写继续调整' },
    { instruction: 'x'.repeat(4_001), expected: '超过 4,000 字符' },
  ])(
    'rejects an invalid follow-up before constructing the request ($expected)',
    ({ instruction, expected }) => {
      expect(() => buildEditingFollowUp(base, history, 'Current', instruction)).toThrow(expected);
    },
  );

  it('rejects missing previews and malformed context instead of adding action messages', () => {
    expect(() => buildEditingFollowUp(base, history, ' ', 'Next')).toThrow('当前预览为空');
    expect(() =>
      buildEditingFollowUp([{ role: 'assistant', content: 'Bad' }], history, 'Preview', 'Next'),
    ).toThrow('编辑上下文无效');
    expect(() =>
      buildEditingFollowUp(base, [{ instruction: 'Bad', result: '' }], 'Preview', 'Next'),
    ).toThrow('编辑记录无效');
    const messages = buildEditingFollowUp(base, history, '{"tool_calls":["write"]}', '继续');
    expect(messages[messages.length - 2]?.content).toBe('{"tool_calls":["write"]}');
    expect(
      messages.every((message) => Object.keys(message).sort().join(',') === 'content,role'),
    ).toBe(true);
  });

  it('uses English revision instructions and errors without translating user or document content', () => {
    setUiLanguage('en');
    const messages = buildEditingFollowUp(base, history, '用户手动修改', '保留原来的语言');
    expect(messages[0].content).toContain('complete revised result');
    expect(messages[messages.length - 2]?.content).toBe('用户手动修改');
    expect(messages[messages.length - 1]?.content).toBe('保留原来的语言');
    expect(() => buildEditingFollowUp(base, [], '', 'Next')).toThrow('current preview is empty');
  });
});
