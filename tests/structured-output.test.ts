import { describe, expect, it } from 'vitest';
import {
  columnName,
  hasExplicitListNumbers,
  parseReport,
  parseTable,
  reportToHtml,
  reportToText,
  tableMatrix,
  tableToText,
} from '../src/core/structured-output';
import {
  COMPLEX_REPORT,
  COMPLEX_TABLE,
  COMPLEX_SLIDES,
  COMPLEX_SOURCES,
} from '../src/office/complex-samples';
import { parseOutline } from '../src/core/output';
import { buildMessages, TASKS } from '../src/core/tasks';

describe('structured document validation and conversion', () => {
  it('keeps a short numbered action list with its introduction and prevents split item lines', () => {
    const bullets = [
      '1. 补齐资料，负责人待补充。',
      '2. 完成复核，日期待补充。',
      '3. 核对签到，工期待补充。',
    ];
    const report = parseReport(
      JSON.stringify({
        title: '行动',
        sections: [{ heading: '后续行动', paragraphs: ['按以下顺序处理。'], bullets }],
      }),
    );
    const doc = new DOMParser().parseFromString(reportToHtml(report), 'text/html');
    const paragraphs = [...doc.querySelectorAll('p')];
    expect(paragraphs.map((p) => p.textContent)).toEqual(['按以下顺序处理。', ...bullets]);
    for (const p of paragraphs.slice(0, -1)) expect(p.style.pageBreakAfter).toBe('avoid');
    expect(paragraphs[paragraphs.length - 1].style.pageBreakAfter).toBe('');
    for (const p of paragraphs.slice(1)) expect(p.style.pageBreakInside).toBe('avoid');
  });

  it.each([
    { bullets: Array.from({ length: 7 }, (_, index) => `${index + 1}. 任务`) },
    { bullets: Array.from({ length: 3 }, (_, index) => `${index + 1}. ${'长'.repeat(240)}`) },
    { bullets: [`1. ${'多行\n'.repeat(80)}`] },
  ])(
    'lets long numbered lists paginate instead of keeping an arbitrarily large block together',
    ({ bullets }) => {
      const report = parseReport(
        JSON.stringify({
          title: '行动',
          sections: [{ heading: '后续行动', paragraphs: ['说明'], bullets }],
        }),
      );
      const doc = new DOMParser().parseFromString(reportToHtml(report), 'text/html');
      for (const p of doc.querySelectorAll('p')) expect(p.style.pageBreakAfter).toBe('');
      if (bullets[0].includes('\n')) {
        expect(doc.querySelectorAll('p')[1].style.pageBreakInside).toBe('');
        expect(doc.querySelectorAll('br')).toHaveLength(79);
      }
    },
  );

  it('keeps a compact table together with its final introduction but releases its last row', () => {
    const report = parseReport(
      JSON.stringify({
        title: '表格分页',
        sections: [
          {
            heading: '检查结果',
            paragraphs: ['较早的说明。', '紧邻表格的引导段。'],
            table: {
              headers: ['事项', '数量'],
              rows: Array.from({ length: 6 }, (_, i) => [`项目 ${i + 1}`, String(i)]),
            },
          },
        ],
      }),
    );
    const doc = new DOMParser().parseFromString(reportToHtml(report), 'text/html');
    const bodyParagraphs = [...doc.body.children].filter(
      (node) => node.tagName === 'P',
    ) as HTMLParagraphElement[];
    expect(bodyParagraphs[0].style.pageBreakAfter).toBe('');
    expect(bodyParagraphs[1].style.pageBreakAfter).toBe('avoid');
    expect(doc.querySelector('h2')!.style.pageBreakAfter).toBe('avoid');
    expect(doc.querySelector('table')!.style.pageBreakInside).toBe('avoid');
    const rows = [...doc.querySelectorAll('tr')];
    for (const row of rows.slice(0, -1)) {
      expect(row.style.pageBreakInside).toBe('avoid');
      for (const p of row.querySelectorAll('p')) expect(p.style.pageBreakAfter).toBe('avoid');
    }
    for (const p of rows[rows.length - 1].querySelectorAll('p'))
      expect(p.style.pageBreakAfter).toBe('');
    expect([...doc.querySelectorAll('td')].map((cell) => cell.textContent)).toEqual(
      report.sections[0].table!.rows.flat(),
    );
  });

  it.each([
    { headers: ['事项'], rows: Array.from({ length: 7 }, (_, i) => [String(i)]) },
    { headers: Array.from({ length: 7 }, (_, i) => `列 ${i}`), rows: [Array(7).fill('短')] },
    { headers: ['事项', '详情'], rows: Array.from({ length: 6 }, () => ['事项', '长'.repeat(90)]) },
  ])('lets long, wide or dense tables paginate without chaining the introduction', (table) => {
    const report = parseReport(
      JSON.stringify({
        title: '长表',
        sections: [{ heading: '详情', paragraphs: ['说明'], table }],
      }),
    );
    const doc = new DOMParser().parseFromString(reportToHtml(report), 'text/html');
    expect(doc.querySelector('table')!.style.pageBreakInside).toBe('');
    for (const p of doc.querySelectorAll('p')) expect(p.style.pageBreakAfter).toBe('');
    expect(doc.querySelectorAll('tbody tr')).toHaveLength(table.rows.length);
  });

  it.each([
    [['1. 首项：保留 12 个检查。', '2. 次项 <待补充>&复核'], true],
    [['1、第一步', '2、第二步'], true],
    [['（1）核验', '（2）交接'], true],
    [['3) 核验', '4) 交接'], true],
    [['1.5% 目标', '2.5% 目标'], false],
    [['1. 核验', '没有编号'], false],
    [['1. 核验', '3. 交接'], false],
    [['1. 核验', '2、交接'], false],
    [[], false],
  ])(
    'recognizes explicit consecutive numbering without stripping authored text',
    (bullets, numbered) => {
      expect(hasExplicitListNumbers(bullets)).toBe(numbered);
      const report = parseReport(
        JSON.stringify({
          title: '步骤',
          sections: [{ heading: '后续行动', paragraphs: ['说明'], bullets }],
        }),
      );
      const doc = new DOMParser().parseFromString(reportToHtml(report), 'text/html');
      expect(doc.querySelectorAll('li')).toHaveLength(numbered ? 0 : bullets.length);
      for (const bullet of bullets) {
        expect(doc.body.textContent).toContain(bullet);
        expect(reportToText(report)).toContain(numbered ? bullet : `• ${bullet}`);
        if (numbered) expect(reportToText(report)).not.toContain(`• ${bullet}`);
      }
    },
  );

  it('retains every section, table cell and fact in a complex report', () => {
    const report = parseReport(JSON.stringify(COMPLEX_REPORT));
    expect(report.sections).toHaveLength(6);
    expect(report.sections.filter((section) => section.table)).toHaveLength(2);
    const html = new DOMParser().parseFromString(reportToHtml(report), 'text/html');
    expect(html.querySelectorAll('h2')).toHaveLength(6);
    expect(html.querySelectorAll('table')).toHaveLength(2);
    expect(html.querySelectorAll('thead th')).toHaveLength(8);
    for (const section of report.sections) {
      expect(html.body.textContent).toContain(section.heading);
      for (const row of section.table?.rows ?? [])
        for (const cell of row) expect(html.body.textContent).toContain(cell);
    }
    expect(reportToText(report)).toContain('87.5%');
  });
  it('treats model HTML, scripts, image handlers and hyperlinks as text', () => {
    const attack =
      '<img src=x onerror="alert(1)"><script>alert(2)</script><a href="javascript:bad">链接</a>&';
    const report = parseReport(
      JSON.stringify({
        title: attack,
        sections: [
          { heading: '测试', paragraphs: [attack], table: { headers: ['字段'], rows: [[attack]] } },
        ],
      }),
    );
    const html = new DOMParser().parseFromString(reportToHtml(report), 'text/html');
    expect(html.querySelector('script, img, a, [onerror]')).toBeNull();
    expect(html.body.textContent).toContain(attack);
  });
  it.each([
    { title: '空报告', sections: [] },
    { title: '报告', sections: [{ heading: '空节' }] },
    {
      title: '报告',
      sections: [{ heading: '表格', table: { headers: ['A', 'B'], rows: [['缺少一列']] } }],
    },
    { title: '报告', sections: [{ heading: '正文', paragraphs: ['x'.repeat(4001)] }] },
  ])('rejects malformed reports instead of silently dropping content', (report) => {
    expect(() => parseReport(JSON.stringify(report))).toThrow();
  });
  it('accepts the full twelve-slide fixture and bounds all demo prompts', () => {
    expect(parseOutline(JSON.stringify({ slides: COMPLEX_SLIDES }))).toHaveLength(12);
    for (const [host, source] of Object.entries(COMPLEX_SOURCES)) {
      expect(source.length).toBeLessThanOrEqual(30000);
      for (const task of TASKS[host as keyof typeof TASKS])
        expect(buildMessages(task, source, '使用合成样稿验证', '英语')).toHaveLength(2);
    }
  });
  it('rejects unrendered content instead of dropping it from a document', () => {
    expect(() =>
      parseReport(JSON.stringify({ ...COMPLEX_REPORT, footnotes: ['不能丢失的注释'] })),
    ).toThrow('暂不支持');
    expect(() => parseTable(JSON.stringify({ ...COMPLEX_TABLE, notes: ['未呈现内容'] }))).toThrow(
      '暂不支持',
    );
    expect(() =>
      parseOutline(
        JSON.stringify({ slides: [{ title: '页面', bullets: ['要点'], notes: '未呈现内容' }] }),
      ),
    ).toThrow('第 1 页');
  });
});

describe('typed spreadsheet output', () => {
  it('keeps numeric values, booleans, blank cells and explicit formulas distinct', () => {
    const table = parseTable(
      JSON.stringify({
        name: '数据',
        headers: ['文本', '数值', '状态', '缺失', '公式'],
        rows: [['001', 0, false, null, { formula: '=B2+1' }]],
      }),
    );
    expect(tableMatrix(table)).toEqual([
      ["'文本", "'数值", "'状态", "'缺失", "'公式"],
      ["'001", 0, false, '', '=B2+1'],
    ]);
    expect(table.formats).toEqual(Array(5).fill('general'));
  });
  it('never evaluates strings as formulas and preserves literal apostrophes', () => {
    const cells = ['=HYPERLINK("https://example.test")', '+123', '-123', '@SUM(A1)', "'原文"];
    const table = parseTable(
      JSON.stringify({ name: '文本测试', headers: ['A', 'B', 'C', 'D', 'E'], rows: [cells] }),
    );
    expect(tableMatrix(table)[1]).toEqual(cells.map((cell) => `'${cell}`));
    expect(tableToText(table)).toContain("'=HYPERLINK");
  });
  it('retains 48 data rows and 96 formulas with reproducible reference arithmetic', () => {
    const table = parseTable(JSON.stringify(COMPLEX_TABLE));
    expect(table.rows).toHaveLength(48);
    expect(table.rows.flat().filter((cell) => cell && typeof cell === 'object')).toHaveLength(96);
    for (const index of [0, 23, 47]) {
      const row = table.rows[index];
      expect(row[5]).toEqual({ formula: `=D${index + 2}-E${index + 2}` });
      expect(row[6]).toEqual({ formula: `=IFERROR(F${index + 2}/D${index + 2},0)` });
      expect((row[3] as number) - (row[4] as number)).toBe(400 + index * 18.125);
    }
    expect(tableToText(table)).toContain('"合成数据\n保留两行备注"');
  });
  it.each([
    { name: '非法/名称', headers: ['A'], rows: [[1]] },
    { name: '表', headers: ['A', 'a'], rows: [[1, 2]] },
    { name: '表', headers: ['A', 'B'], rows: [[1]] },
    { name: '表', headers: ['A'], rows: [[{ formula: '=WEBSERVICE("x")' }]] },
    { name: '表', headers: ['A'], formats: ['text'], rows: [[{ formula: '=1+1' }]] },
    { name: '表', headers: ['A'], formats: ['unknown'], rows: [[1]] },
    { name: '表', headers: ['A'], rows: [[{}]] },
    { name: '表', headers: ['A'], rows: Array(201).fill([1]) },
    {
      name: '表',
      headers: Array.from({ length: 12 }, (_, i) => String(i)),
      rows: Array(200).fill(Array(12).fill(1)),
    },
  ])('rejects unsafe or inconsistent table input', (table) => {
    expect(() => parseTable(JSON.stringify(table))).toThrow();
  });
  it('reports a precise cell for malformed formulas', () => {
    expect(() =>
      parseTable(
        JSON.stringify({ name: '表', headers: ['金额'], rows: [[1], [{ formula: '=SUM(A1' }]] }),
      ),
    ).toThrow('第 2 行第 1 列');
    expect(columnName(0)).toBe('A');
    expect(columnName(11)).toBe('L');
    expect(columnName(26)).toBe('AA');
  });
});
