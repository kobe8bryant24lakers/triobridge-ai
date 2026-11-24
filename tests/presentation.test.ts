import { describe, expect, it } from 'vitest';
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import {
  hasCredentialText,
  materialFileAllowed,
  materialMessages,
  parseDelimited,
  readMaterials,
} from '../src/presentation/materials';
import { chartData, numeric, tableFromText } from '../src/presentation/charts';
import { parseDeck } from '../src/presentation/content';
import {
  alignProposal,
  assertSnapshotUnchanged,
  parseEdit,
  reorderedIds,
  replaceProposal,
  styleProposal,
  validateEdit,
} from '../src/presentation/editing';
import { DEFAULT_THEME, renderDeck, WIDE } from '../src/presentation/render';
import { createDemoPresentation } from '../src/office/presentation-demo';
import type { DeckContent, MaterialContext } from '../src/presentation/types';

const empty: MaterialContext = { materials: [], includeImages: false };
const context: MaterialContext = {
  includeImages: false,
  materials: [
    {
      id: 'm1',
      name: '季度数据.csv',
      kind: 'spreadsheet',
      text: '',
      tables: [
        {
          id: 'm1-t1',
          name: '经营数据',
          headers: ['月份', '营收', '利润率'],
          rows: [
            ['一月', 0, 0.125],
            ['二月', -2.5, 0.2],
            ['三月', null, 0.275],
          ],
        },
      ],
    },
  ],
};
function file(name: string, data: Uint8Array | string): File {
  const bytes = typeof data === 'string' ? strToU8(data) : data;
  const f = new File([new Uint8Array(bytes)], name);
  Object.defineProperty(f, 'arrayBuffer', { value: async () => Uint8Array.from(bytes).buffer });
  return f;
}
describe('local materials, source boundaries and fidelity', () => {
  it('parses quoted commas, escaped quotes, multiline cells, UTF-8 BOM and tabs without executing formulas', () => {
    expect(parseDelimited('\uFEFF名称,值\r\n"东,部",1\r\n"换\n行","a""b"')).toEqual([
      ['名称', '值'],
      ['东,部', '1'],
      ['换\n行', 'a"b'],
    ]);
    expect(parseDelimited('月\t值\n一月\t=1+2')).toEqual([
      ['月', '值'],
      ['一月', '=1+2'],
    ]);
    expect(() => parseDelimited('a,b\n"unterminated')).toThrow('未闭合');
    expect(() => parseDelimited('a,b\n"x"oops,y')).toThrow('分隔符');
    expect(() => tableFromText('a,b\n1,2,3')).toThrow('列数');
  });
  it('excludes hidden files and credential text, preserves selected valid files and reports each failure', async () => {
    for (const name of [
      '.env',
      '.git/config.json',
      'data/.private/a.txt',
      'secrets.txt',
      'node_modules/a.md',
      'report.exe',
    ])
      expect(materialFileAllowed(name)).toBe(false);
    expect(materialFileAllowed('业务素材/经营报告.docx')).toBe(true);
    expect(hasCredentialText('api_key="example-private-value"')).toBe(true);
    const result = await readMaterials([
      file('note.md', '演示素材：收入增长。'),
      file('private.txt', 'api_key="example-private-value"'),
      file('unsupported.exe', 'binary'),
    ]);
    expect(result.materials.map((m) => m.name)).toEqual(['note.md']);
    expect(result.issues).toHaveLength(2);
    expect(
      JSON.stringify(
        materialMessages({ ...empty, materials: result.materials }, 'system', 'brief'),
      ),
    ).not.toContain('example-private-value');
  });
  it('reads DOCX paragraphs/table and rejects entities and oversized decompressed XML', async () => {
    const w = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
    const bytes = zipSync({
      'word/document.xml': strToU8(
        `<w:document xmlns:w="${w}"><w:body><w:p><w:r><w:t>项目概况</w:t></w:r></w:p><w:tbl><w:tr><w:tc><w:p><w:r><w:t>名称</w:t></w:r></w:p></w:tc></w:tr><w:tr><w:tc><w:p><w:r><w:t>项目甲</w:t></w:r></w:p></w:tc></w:tr></w:tbl></w:body></w:document>`,
      ),
    });
    const result = await readMaterials([file('report.docx', bytes)]);
    expect(result.materials[0].tables[0].rows).toEqual([['项目甲']]);
    expect(result.materials[0].text).toContain('项目概况');
    const bad = await readMaterials([
      file(
        'bad.docx',
        zipSync({
          'word/document.xml': strToU8(
            '<!DOCTYPE x [<!ENTITY ext SYSTEM "https://example.test">]><x/>',
          ),
        }),
      ),
    ]);
    expect(bad.issues[0].message).toContain('XML');
    const large = await readMaterials([
      file('large.docx', zipSync({ 'word/document.xml': new Uint8Array(8_000_001) })),
    ]);
    expect(large.issues[0].message).toContain('解压');
  });
  it('reads XLSX typed cells and cached formulas, retaining missing caches as empty', async () => {
    const x = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
      r = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
    const zip = zipSync({
      'xl/workbook.xml': strToU8(
        `<workbook xmlns="${x}" xmlns:r="${r}"><sheets><sheet name="数据" r:id="r1"/></sheets></workbook>`,
      ),
      'xl/_rels/workbook.xml.rels': strToU8(
        '<Relationships><Relationship Id="r1" Target="worksheets/sheet1.xml"/></Relationships>',
      ),
      'xl/worksheets/sheet1.xml': strToU8(
        `<worksheet xmlns="${x}"><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>月份</t></is></c><c r="B1" t="inlineStr"><is><t>数值</t></is></c></row><row r="2"><c r="A2" t="inlineStr"><is><t>一月</t></is></c><c r="B2"><f>1-1</f><v>0</v></c></row><row r="3"><c r="A3" t="inlineStr"><is><t>二月</t></is></c><c r="B3"><f>1+1</f></c></row></sheetData></worksheet>`,
      ),
    });
    const result = await readMaterials([file('source.xlsx', zip)]);
    expect(result.materials[0].tables[0].rows).toEqual([
      ['一月', 0],
      ['二月', null],
    ]);
    expect(result.materials[0].warning).toContain('1');
  });
  it('rejects excessive context rather than truncating and only includes opted-in thumbnails', () => {
    expect(() => materialMessages(empty, 'system', 'x'.repeat(60001))).toThrow('没有自动截断');
    const image = {
      id: 'img',
      name: '图片.png',
      kind: 'image' as const,
      text: '',
      tables: [],
      image: { data: 'full-image', modelData: 'data:image/jpeg;base64,eA==', width: 4, height: 4 },
    };
    expect(
      typeof materialMessages({ materials: [image], includeImages: false }, 'system', 'brief')[1]
        .content,
    ).toBe('string');
    const multimodal = materialMessages(
      { materials: [image], includeImages: true },
      'system',
      'brief',
    )[1].content;
    expect(JSON.stringify(multimodal)).toContain('data:image/jpeg;base64,eA==');
    expect(JSON.stringify(multimodal)).not.toContain('full-image');
    expect(() =>
      materialMessages({ materials: Array(5).fill(image), includeImages: true }, 'system', 'brief'),
    ).toThrow('最多 4 张');
  });
});
describe('chart and model output safeguards', () => {
  it('keeps zero, missing, negative and percent distinct; rejects formulas and invalid numeric cells', () => {
    const table = context.materials[0].tables[0];
    expect(chartData(table, 0, [1], 'line').series[0].values).toEqual([0, -2.5, null]);
    expect(numeric('1,234.56', '')).toBe(1234.56);
    expect(numeric('12.5%', '')).toBe(0.125);
    for (const value of ['=1+2', '1,23', '1.2.3', 'NaN', '12元', true])
      expect(() => numeric(value, '值')).toThrow();
    expect(() => chartData(table, 0, [1], 'pie')).toThrow('非负完整');
    expect(() => chartData(table, 0, [1, 1], 'column')).toThrow('不同');
  });
  it('requires exact page count, known sources and valid layout data', () => {
    const valid = {
      title: '经营汇报',
      slides: [
        {
          title: '月度趋势',
          layout: 'chart',
          chart: { type: 'line', tableId: 'm1-t1', category: 0, series: [1] },
          notes: '原始数据',
          sources: ['m1'],
        },
      ],
    };
    expect(parseDeck(JSON.stringify(valid), context, 1).slides[0].chart?.tableId).toBe('m1-t1');
    expect(() => parseDeck(JSON.stringify(valid), context, 2)).toThrow('需要 2 页');
    expect(() => parseDeck(JSON.stringify({ ...valid, run: 'alert(1)' }), context)).toThrow('字段');
    expect(() =>
      parseDeck(
        JSON.stringify({
          title: '标题',
          slides: [{ title: '页', layout: 'twoColumn', sources: [] }],
        }),
        empty,
      ),
    ).toThrow('缺少');
    expect(() =>
      parseDeck(
        JSON.stringify({
          title: '标题',
          slides: [{ title: '页', layout: 'cover', sources: ['unknown'] }],
        }),
        empty,
      ),
    ).toThrow('未知来源');
  });
  it('generates a real editable chart with identical embedded worksheet values, including blanks', async () => {
    const deck: DeckContent = {
      title: '数据一致性',
      slides: [
        {
          title: '趋势',
          layout: 'chart',
          chart: { type: 'line', tableId: 'm1-t1', category: 0, series: [1, 2] },
          notes: '验收',
          sources: ['m1'],
        },
        {
          title: '流程',
          layout: 'process',
          steps: [
            { title: '检查', detail: '核对原始数据' },
            { title: '应用', detail: '确认后写入' },
          ],
          sources: [],
          notes: '',
        },
      ],
    };
    const base64 = await renderDeck(deck, context, DEFAULT_THEME, WIDE),
      zip = unzipSync(Uint8Array.from(atob(base64), (c) => c.charCodeAt(0)));
    expect(Object.keys(zip).filter((n) => /^ppt\/slides\/slide\d+\.xml$/.test(n))).toHaveLength(2);
    const chart = strFromU8(zip['ppt/charts/chart1.xml']);
    expect(chart).toContain('lineChart');
    expect(chart).toContain('<c:v>-2.5</c:v>');
    expect(chart).toContain('<c:v>0</c:v>');
    expect(chart).not.toMatch(/<c:v>(null|undefined|NaN)<\/c:v>/);
    const workbookName = Object.keys(zip).find((n) => /^ppt\/embeddings\/.*xlsx$/.test(n))!;
    const workbook = unzipSync(zip[workbookName]);
    const sheet = strFromU8(workbook['xl/worksheets/sheet1.xml']);
    expect(sheet).toContain('<v>-2.5</v>');
    expect(sheet).toContain('<v>0.125</v>');
    expect(sheet).not.toMatch(/<v>(null|undefined|NaN)<\/v>/);
    const doc = new DOMParser().parseFromString(sheet, 'application/xml');
    expect(doc.querySelector('c[r="B2"] v')?.textContent).toBe('0');
    expect(doc.querySelector('c[r="B3"] v')?.textContent).toBe('-2.5');
    expect(doc.querySelector('c[r="B4"] v')?.textContent || '').toBe('');
    expect(strFromU8(zip['ppt/notesSlides/notesSlide1.xml'])).toContain('季度数据.csv');
  });
  it('rejects content that its layout would discard and tables too dense to read', async () => {
    expect(() =>
      parseDeck(
        JSON.stringify({
          title: 'x',
          slides: [{ title: 'x', layout: 'cover', bullets: ['must not disappear'], sources: [] }],
        }),
        empty,
      ),
    ).toThrow('版式');
    const dense = structuredClone(context);
    dense.materials[0].tables[0].headers = ['甲', '乙', '丙', '丁', '戊', '己'];
    dense.materials[0].tables[0].rows = Array.from({ length: 10 }, () =>
      Array(6).fill('文字'.repeat(39)),
    );
    await expect(
      renderDeck(
        {
          title: 'x',
          slides: [
            { title: '密集表格', layout: 'table', tableId: 'm1-t1', notes: '', sources: ['m1'] },
          ],
        },
        dense,
      ),
    ).rejects.toThrow('表格文字过密');
  });
});
describe('scoped edits and stale target protection', () => {
  it('rejects unknown IDs/properties and off-slide geometry before writes', async () => {
    const snap = await createDemoPresentation(() => {}).read('selected');
    expect(() =>
      parseEdit(
        '{"summary":"x","changes":[{"slideId":"other","shapeId":"title","text":"x"}]}',
        snap,
      ),
    ).toThrow('范围以外');
    expect(() =>
      validateEdit(
        { summary: 'x', changes: [{ slideId: 'demo-1', shapeId: 'title', left: 900 }] },
        snap,
      ),
    ).toThrow('超出页面');
    expect(() =>
      validateEdit(
        { summary: 'x', changes: [{ slideId: 'demo-1', shapeId: 'title', code: 'bad' }] },
        snap,
      ),
    ).toThrow('字段');
    expect(() =>
      validateEdit({ summary: 'x', changes: [], slideOrder: ['demo-1', 'demo-1'] }, snap),
    ).toThrow('每页且仅一次');
  });
  it('detects style/geometry/content/order changes, but ignores irrelevant selection details', async () => {
    const snap = await createDemoPresentation(() => {}).read('all');
    for (const key of ['left', 'text'] as const) {
      const next = structuredClone(snap);
      Object.assign(next.slides[0].shapes[0], { [key]: key === 'left' ? 88 : '修改' });
      expect(() => assertSnapshotUnchanged(snap, next)).toThrow('已变化');
    }
    const next = structuredClone(snap);
    next.slides[0].shapes[0].font!.size = 31;
    expect(() => assertSnapshotUnchanged(snap, next)).toThrow('已变化');
    expect(() => assertSnapshotUnchanged(snap, { ...snap, selectedShapeIds: [] })).not.toThrow();
    const subset = { ...snap, slides: [snap.slides[0], snap.slides[2]] };
    expect(reorderedIds(subset, ['demo-3', 'demo-1'])).toEqual(['demo-3', 'demo-2', 'demo-1']);
  });
  it('produces actual style/replace/alignment changes and disallows reapplying an old snapshot', async () => {
    const bridge = createDemoPresentation(() => {}),
      snapshot = await bridge.read('selected');
    expect(styleProposal(snapshot, 'Aptos', '#28604A').changes.length).toBeGreaterThan(0);
    expect(alignProposal(snapshot, 'left').changes).toHaveLength(2);
    const plan = replaceProposal(snapshot, '24 小时', '12 小时');
    expect(plan.changes).toHaveLength(1);
    await bridge.apply(snapshot, plan);
    await expect(bridge.apply(snapshot, plan)).rejects.toThrow('已变化');
  });
});
