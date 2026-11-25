import { describe, expect, it, vi } from 'vitest';
import { parseReport, parseTable } from '../src/core/structured-output';
import { translateModelMessage } from '../src/core/model-ui-messages';

function diagnostic(run: () => unknown) {
  try {
    run();
  } catch (error) {
    return translateModelMessage((error as Error).message);
  }
  throw new Error('The fixture should have failed validation');
}

describe('model diagnostic translations', () => {
  it('retains the configured host, port and loopback guidance from the actual transport error', async () => {
    const { modelConnectionError } = await vi.importActual<{
      modelConnectionError: (error: Error & { code: string }, endpoint: string) => string;
    }>('../scripts/model-connection-error.mjs');
    const message = modelConnectionError(
      Object.assign(new Error('Connection refused'), { code: 'ECONNREFUSED' }),
      'http://[::1]:8041/v1/messages',
    );
    const translated = translateModelMessage(message);
    expect(translated).toContain('[::1]:8041 refused');
    expect(translated).toContain('ECONNREFUSED');
    expect(translated).toContain('Loopback');
  });

  it('retains report validation field names and Excel cell positions', () => {
    const report = { title: 'R', sections: [{ heading: 'S', 自定义字段: 'value' }] };
    expect(diagnostic(() => parseReport(JSON.stringify(report)))).toContain(
      'Section 1 contains unsupported fields: 自定义字段',
    );
    const table = { name: 'A', headers: ['B'], rows: [[{ formula: '=SUM(' }]] };
    expect(diagnostic(() => parseTable(JSON.stringify(table)))).toMatch(
      /Row 1, column 1: .*parenthesis/,
    );
  });

  it('does not translate unrelated user content or inherited object properties', () => {
    for (const input of [
      '模型服务 用户文字',
      '这是我们的报告，需要 1–10 字符。',
      'toString',
      '__proto__',
    ]) {
      expect(translateModelMessage(input)).toBeUndefined();
    }
  });
});
