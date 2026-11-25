import { afterEach, describe, expect, it, vi } from 'vitest';
import { createOfficeBridge } from '../src/office/bridge';
import { wordSelectionWarning } from '../src/office/word-safety';
import { COMPLEX_REPORT, COMPLEX_TABLE } from '../src/office/complex-samples';

const wordXml = (content = '<w:p><w:r><w:t>Original</w:t></w:r></w:p>') =>
  `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>${content}</w:body></w:document>`;

afterEach(() => vi.unstubAllGlobals());

function office(supported: string[] = []) {
  const document = {
    getSelectedDataAsync: vi.fn((_type, cb) => cb({ status: 'succeeded', value: 'Selected text' })),
    setSelectedDataAsync: vi.fn((_text, _options, cb) => cb({ status: 'succeeded' })),
    addHandlerAsync: vi.fn((_type, _handler, cb) => cb({ status: 'succeeded' })),
    removeHandlerAsync: vi.fn(),
  };
  vi.stubGlobal('Office', {
    context: {
      requirements: {
        isSetSupported: (name: string, version: string) => supported.includes(`${name}:${version}`),
      },
      document,
    },
    CoercionType: { Text: 'text' },
    AsyncResultStatus: { Succeeded: 'succeeded' },
    EventType: { DocumentSelectionChanged: 'selection' },
  });
  return document;
}

describe('Office API version fallback', () => {
  it('keeps basic PowerPoint text processing without newer APIs', async () => {
    office(['TextCoercion:1.1']);
    const bridge = await createOfficeBridge('powerpoint');
    expect(bridge.capabilities).toMatchObject({
      read: true,
      replace: true,
      slides: false,
    });
    expect((await bridge.readSelection()).text).toBe('Selected text');
    await expect(bridge.prepareWrite('slides')).rejects.toThrow('不支持');
  });
  it('only enables slide creation at PowerPointApi 1.4', async () => {
    office(['PowerPointApi:1.4']);
    expect((await createOfficeBridge('powerpoint')).capabilities.slides).toBe(true);
  });
  it('does not offer unsafe replacement when selection events are unavailable', async () => {
    const doc = office(['TextCoercion:1.1']);
    doc.addHandlerAsync.mockImplementation((_type, _handler, cb) => cb({ status: 'failed' }));
    const bridge = await createOfficeBridge('powerpoint');
    expect(bridge.capabilities.replace).toBe(false);
    expect((await bridge.readSelection()).replace).toBeUndefined();
  });
  it('detects movement even when two selected texts happen to be equal', async () => {
    const doc = office(['TextCoercion:1.1']);
    const bridge = await createOfficeBridge('powerpoint');
    const source = await bridge.readSelection();
    const handler = doc.addHandlerAsync.mock.calls[0][1];
    handler();
    const plan = await bridge.prepareWrite('replace', source);
    await expect(plan.commit('updated')).rejects.toThrow('已经变化');
    expect(doc.setSelectedDataAsync).not.toHaveBeenCalled();
  });
  it('detects edits within a PowerPoint selection before writing', async () => {
    const doc = office(['TextCoercion:1.1']);
    const bridge = await createOfficeBridge('powerpoint');
    const source = await bridge.readSelection();
    doc.getSelectedDataAsync.mockImplementation((_type, cb) =>
      cb({ status: 'succeeded', value: 'Edited text' }),
    );
    await expect(source.replace!('updated')).rejects.toThrow('已经变化');
    expect(doc.setSelectedDataAsync).not.toHaveBeenCalled();
  });
  it('writes only after explicit commit and unregisters its own handler', async () => {
    const doc = office(['TextCoercion:1.1']);
    const bridge = await createOfficeBridge('powerpoint');
    const plan = await bridge.prepareWrite('replace', await bridge.readSelection());
    expect(doc.setSelectedDataAsync).not.toHaveBeenCalled();
    await plan.commit('updated');
    expect(doc.setSelectedDataAsync).toHaveBeenCalledWith(
      'updated',
      { coercionType: 'text' },
      expect.any(Function),
    );
    bridge.dispose!();
    expect(doc.removeHandlerAsync).toHaveBeenCalled();
  });
  it('keeps Word text-only fallback readable without enabling untracked rich-text writes', async () => {
    const doc = office(['TextCoercion:1.1']);
    const bridge = await createOfficeBridge('word');
    const source = await bridge.readSelection();
    expect(source.text).toBe('Selected text');
    expect(source.warning).toContain('不支持自动写入排版结果');
    expect(source.replace).toBeUndefined();
    expect(bridge.capabilities).toMatchObject({ read: true, replace: false, insert: false });
    await expect(bridge.prepareWrite('replace', source, '# 中文标题')).rejects.toThrow('不支持');
    await expect(bridge.prepareWrite('insert', source, '# 中文标题')).rejects.toThrow('不支持');
    expect(doc.setSelectedDataAsync).not.toHaveBeenCalled();
  });
});

describe('Word tracked range and stale content protection', () => {
  const leadingBoundary = '<p class="trioWordBefore"><br></p>';
  const trailingBoundary = '<p class="trioWordAfter"><br></p>';
  function setup(boundaryApi = false) {
    office(['WordApi:1.1', ...(boundaryApi ? ['WordApi:1.3'] : [])]);
    const paragraph = { insertHtml: vi.fn() };
    const boundary = { before: 'After', after: 'Before' };
    const paragraphStart = {},
      paragraphEnd = {};
    const paragraphs = {
      items: [{ style: 'Normal' }, { style: 'Quote' }],
      load: vi.fn(),
      getFirst: () => ({ getRange: () => paragraphStart }),
      getLast: () => ({ getRange: () => paragraphEnd }),
    };
    const point = {
      paragraphs,
      getRange: (_location: string): unknown => point,
      compareLocationWith: vi.fn((other) => ({
        value: other === paragraphStart ? boundary.before : boundary.after,
      })),
    };
    const range = {
      text: 'Original',
      paragraphs,
      getRange: vi.fn((_location: string) => point),
      load: vi.fn(),
      insertText: vi.fn(),
      insertHtml: vi.fn(),
      insertParagraph: vi.fn(() => paragraph),
      getOoxml: vi.fn(() => ({ value: wordXml() })),
    };
    const context = {
      document: { getSelection: () => range },
      sync: vi.fn().mockResolvedValue(undefined),
      trackedObjects: { add: vi.fn(), remove: vi.fn() },
    };
    vi.stubGlobal('Word', {
      InsertLocation: { replace: 'Replace', end: 'End', after: 'After' },
      run: vi.fn((...args) => args[args.length - 1](context)),
    });
    return { range, context, paragraph, boundary, point, paragraphs };
  }
  it('writes back to the tracked source rather than a later cursor location', async () => {
    const { range, context } = setup();
    const bridge = await createOfficeBridge('word');
    const source = await bridge.readSelection();
    context.document.getSelection = () => ({
      ...range,
      text: 'Different selection',
      insertText: vi.fn(),
    });
    await source.replace!('New text');
    expect(range.insertText).toHaveBeenCalledWith('New text', 'Replace');
    expect(range.insertHtml).not.toHaveBeenCalled();
    await source.dispose!();
    expect(context.trackedObjects.remove).toHaveBeenCalledWith(range);
  });
  it('refuses to overwrite user edits after generating a preview', async () => {
    const { range } = setup();
    const source = await (await createOfficeBridge('word')).readSelection();
    range.text = 'User changed this';
    await expect(source.replace!('Generated text')).rejects.toThrow('已经变化');
    expect(range.insertText).not.toHaveBeenCalled();
  });
  it('inserts at the reviewed anchor without replacing the selection', async () => {
    const { range } = setup();
    const plan = await (await createOfficeBridge('word')).prepareWrite('insert');
    await plan.commit('New text');
    expect(range.insertText).toHaveBeenCalledWith('New text', 'End');
    expect(range.insertHtml).not.toHaveBeenCalled();
    expect(range.insertParagraph).not.toHaveBeenCalled();
  });
  it('replaces the tracked selection with formatted Chinese Markdown only after confirmation', async () => {
    const { range, context } = setup();
    const bridge = await createOfficeBridge('word');
    const source = await bridge.readSelection();
    const output = '# 季度经营回顾\n\n**重点结论**：保持原始日期。\n\n- 完成复盘\n- 确认计划';
    const plan = await bridge.prepareWrite('replace', source, output);
    expect(plan.warning).toContain('按 Markdown 排版');
    const laterSelection = { ...range, text: 'Later cursor', insertHtml: vi.fn() };
    context.document.getSelection = () => laterSelection;
    expect(range.insertHtml).not.toHaveBeenCalled();
    await plan.commit(output);
    expect(range.insertHtml).toHaveBeenCalledWith(
      expect.stringContaining('季度经营回顾'),
      'Replace',
    );
    const html = range.insertHtml.mock.calls[0][0];
    expect(html).toContain(leadingBoundary + '<h1');
    expect(html.endsWith(trailingBoundary)).toBe(true);
    expect(html).toContain('<h1');
    expect(html).toContain('<strong>重点结论</strong>');
    expect(html).toContain('<li');
    expect(html).not.toContain('# 季度经营回顾');
    expect(range.insertText).not.toHaveBeenCalled();
    expect(range.insertParagraph).not.toHaveBeenCalled();
    expect(range.insertHtml).toHaveBeenCalledTimes(1);
    expect(laterSelection.insertHtml).not.toHaveBeenCalled();
  });
  it.each(['Original paragraph text', ''])(
    'inserts a bounded Markdown block at the exact reviewed range end or cursor (%j)',
    async (before) => {
      const { range, paragraph } = setup();
      range.text = before;
      range.insertParagraph.mockImplementation(() => {
        throw new Error('Must not move insertion after the whole parent paragraph');
      });
      const output = '# 中文标题\n\n正文中的**重点**。';
      const bridge = await createOfficeBridge('word');
      const plan = await bridge.prepareWrite('insert', undefined, output);
      expect(plan.before).toBe(before);
      expect(plan.label).toContain('新起段落');
      expect(range.insertParagraph).not.toHaveBeenCalled();
      await plan.commit(output);
      expect(range.insertParagraph).not.toHaveBeenCalled();
      expect(paragraph.insertHtml).not.toHaveBeenCalled();
      expect(range.insertHtml).toHaveBeenCalledWith(
        expect.stringContaining(leadingBoundary + '<h1'),
        'End',
      );
      expect(range.insertHtml.mock.calls[0][0]).toContain('中文标题');
      expect(range.insertHtml.mock.calls[0][0].endsWith(trailingBoundary)).toBe(true);
      expect(range.insertHtml).toHaveBeenCalledTimes(1);
      expect(range.insertText).not.toHaveBeenCalled();
      expect(range.text).toBe(before);
    },
  );
  it.each(['replace', 'insert'] as const)(
    'keeps inline emphasis in the existing paragraph during %s',
    async (kind) => {
      const { range, paragraph, point, paragraphs } = setup(true);
      const bridge = await createOfficeBridge('word');
      const source = await bridge.readSelection();
      const output = '中文 **重点** 与原文衔接';
      const plan = await bridge.prepareWrite(kind, source, output);
      expect(plan.label).not.toContain('新起段落');
      await plan.commit(output);
      expect(range.insertHtml).toHaveBeenCalledWith(
        '中文 <strong>重点</strong> 与原文衔接',
        kind === 'replace' ? 'Replace' : 'End',
      );
      expect(range.insertHtml).toHaveBeenCalledTimes(1);
      expect(range.insertText).not.toHaveBeenCalled();
      expect(range.insertParagraph).not.toHaveBeenCalled();
      expect(paragraph.insertHtml).not.toHaveBeenCalled();
      expect(point.compareLocationWith).not.toHaveBeenCalled();
      expect(paragraphs.load).not.toHaveBeenCalled();
    },
  );
  it('isolates a list replacement from neighboring paragraph text in one write', async () => {
    const { range } = setup();
    const source = await (await createOfficeBridge('word')).readSelection();
    await source.replace!('- 第一项\n- 第二项');
    const html = range.insertHtml.mock.calls[0][0];
    expect(html).toContain(leadingBoundary + '<ul');
    expect(html.endsWith('</ul>' + trailingBoundary)).toBe(true);
    expect(range.insertHtml).toHaveBeenCalledTimes(1);
    expect(range.insertText).not.toHaveBeenCalled();
    expect(range.insertParagraph).not.toHaveBeenCalled();
  });
  it.each([false, true])(
    'preserves distinct localized paragraph styles around a block replacement (exact boundaries: %s)',
    async (precise) => {
      const { range, paragraphs } = setup(precise);
      paragraphs.items = [{ style: '正文' }, { style: 'QA Edit 自定义' }];
      const source = await (await createOfficeBridge('word')).readSelection();
      await source.replace!('## 新标题\n\n- 新项目');
      const html = range.insertHtml.mock.calls[0][0];
      expect(paragraphs.load).toHaveBeenCalledWith('items/style');
      expect(html).toContain('p.trioWordBefore{mso-style-name:"正文";}');
      expect(html).toContain('p.trioWordAfter{mso-style-name:"QA Edit 自定义";}');
      expect(html).toContain(leadingBoundary + '<h2');
      expect(html.endsWith(trailingBoundary)).toBe(true);
      expect(range.insertHtml).toHaveBeenCalledTimes(1);
      expect(range.insertParagraph).not.toHaveBeenCalled();
    },
  );
  it('uses the end paragraph style on both sides of an insertion into a multi-paragraph selection', async () => {
    const { range, paragraphs } = setup();
    paragraphs.items = [{ style: 'Heading 1' }, { style: 'Quote' }];
    const output = '## 新标题';
    const plan = await (await createOfficeBridge('word')).prepareWrite('insert', undefined, output);
    await plan.commit(output);
    const html = range.insertHtml.mock.calls[0][0];
    expect(html).toContain('p.trioWordBefore{mso-style-name:"Quote";}');
    expect(html).toContain('p.trioWordAfter{mso-style-name:"Quote";}');
    expect(html).not.toContain('Heading 1');
    expect(range.insertHtml).toHaveBeenCalledTimes(1);
  });
  it('escapes document style names as CSS strings without allowing HTML or CSS injection', async () => {
    const { range, paragraphs } = setup(true);
    paragraphs.items = [{ style: 'QA "quoted"\\style\n</style><img src=x>&' }];
    const source = await (await createOfficeBridge('word')).readSelection();
    await source.replace!('## 新标题');
    const html = range.insertHtml.mock.calls[0][0];
    expect(html).toContain(
      'QA \\22 quoted\\22 \\5c style\\a \\3c /style\\3e \\3c img src=x\\3e \\26 ',
    );
    expect(html.match(/<style>/g)).toHaveLength(1);
    expect(html.match(/<\/style>/g)).toHaveLength(1);
    expect(html).not.toContain('<img');
  });
  it('leaves the source untouched if its outside paragraph style cannot be confirmed', async () => {
    const { range, paragraphs } = setup(true);
    paragraphs.items = [];
    const source = await (await createOfficeBridge('word')).readSelection();
    await expect(source.replace!('## 新标题')).rejects.toThrow('无法确认原段落样式');
    expect(range.insertHtml).not.toHaveBeenCalled();
    expect(range.insertParagraph).not.toHaveBeenCalled();
    expect(range.insertText).not.toHaveBeenCalled();
  });
  it.each([
    ['After', 'Before', true, true],
    ['Equal', 'Before', false, true],
    ['After', 'Equal', true, false],
    ['Equal', 'Equal', false, false],
    ['Equal', 'After', false, false],
  ] as const)(
    'adds only necessary boundaries for a Word 1.3 replacement (%s / %s)',
    async (before, after, leading, trailing) => {
      const { range, boundary, point } = setup(true);
      boundary.before = before;
      boundary.after = after;
      const source = await (await createOfficeBridge('word')).readSelection();
      await source.replace!('# 标题\n\n正文');
      const html = range.insertHtml.mock.calls[0][0];
      expect(html.includes(leadingBoundary)).toBe(leading);
      expect(html.endsWith(trailingBoundary)).toBe(trailing);
      expect(html.includes('<style>')).toBe(leading || trailing);
      expect(point.compareLocationWith).toHaveBeenCalledTimes(2);
      expect(range.insertHtml).toHaveBeenCalledTimes(1);
      expect(range.insertParagraph).not.toHaveBeenCalled();
      expect(range.insertText).not.toHaveBeenCalled();
    },
  );
  it('does not add a trailing blank paragraph when inserting a block at document end', async () => {
    const { range, boundary } = setup(true);
    boundary.before = 'After';
    boundary.after = 'Equal';
    const output = '# 文末标题';
    const plan = await (await createOfficeBridge('word')).prepareWrite('insert', undefined, output);
    await plan.commit(output);
    expect(range.getRange).toHaveBeenCalledWith('End');
    const [html, location] = range.insertHtml.mock.calls[0];
    expect(html.includes(leadingBoundary)).toBe(true);
    expect(html.endsWith(trailingBoundary)).toBe(false);
    expect(location).toBe('End');
    expect(range.insertParagraph).not.toHaveBeenCalled();
  });
  it('rejects a stale Markdown insertion anchor before creating any paragraph', async () => {
    const { range, paragraph } = setup();
    const output = '# 标题\n\n正文';
    const plan = await (await createOfficeBridge('word')).prepareWrite('insert', undefined, output);
    range.text = 'User changed this';
    await expect(plan.commit(output)).rejects.toThrow('已经变化');
    expect(range.insertParagraph).not.toHaveBeenCalled();
    expect(paragraph.insertHtml).not.toHaveBeenCalled();
  });
  it.each(['replace', 'insert'] as const)(
    'does not pre-edit or retry the source when the single Markdown %s fails',
    async (kind) => {
      const { range, paragraph, context } = setup(true);
      const bridge = await createOfficeBridge('word');
      const source = await bridge.readSelection();
      const output = '## 新标题\n\n- 新项目';
      const plan = await bridge.prepareWrite(kind, source, output);
      context.sync
        .mockResolvedValueOnce(undefined)
        .mockRejectedValueOnce(new Error('Word import failed'));
      await expect(plan.commit(output)).rejects.toThrow('Word import failed');
      expect(range.insertHtml).toHaveBeenCalledTimes(1);
      expect(range.text).toBe('Original');
      expect(range.insertText).not.toHaveBeenCalled();
      expect(range.insertParagraph).not.toHaveBeenCalled();
      expect(paragraph.insertHtml).not.toHaveBeenCalled();
    },
  );
  it('rechecks protected Word objects before rich-text replacement', async () => {
    const { range } = setup();
    const source = await (await createOfficeBridge('word')).readSelection();
    range.getOoxml.mockReturnValue({ value: wordXml('<w:drawing/>') });
    await expect(source.replace!('# 新标题\n\n**新文字**')).rejects.toThrow('保留结构');
    expect(range.insertHtml).not.toHaveBeenCalled();
    expect(range.insertText).not.toHaveBeenCalled();
  });
  it('isolates report headings in a new paragraph so existing text is not restyled', async () => {
    const { range, context, paragraph } = setup();
    const bridge = await createOfficeBridge('word');
    const raw = JSON.stringify(COMPLEX_REPORT);
    await expect(bridge.prepareWrite('document', undefined, '{}')).rejects.toThrow();
    expect(context.trackedObjects.add).not.toHaveBeenCalled();
    const plan = await bridge.prepareWrite('document', undefined, raw);
    expect(range.insertHtml).not.toHaveBeenCalled();
    expect(range.insertParagraph).not.toHaveBeenCalled();
    await plan.commit(raw);
    expect(range.insertParagraph).toHaveBeenCalledWith('', 'After');
    expect(paragraph.insertHtml).toHaveBeenCalledWith(expect.stringContaining('<table'), 'Replace');
    expect(range.insertHtml).not.toHaveBeenCalled();
    expect(range.insertText).not.toHaveBeenCalled();
    await plan.dispose!();
    expect(context.trackedObjects.remove).toHaveBeenCalledWith(range);
  });
  it.each([
    'tbl',
    'drawing',
    'pict',
    'object',
    'sdt',
    'fldChar',
    'fldSimple',
    'hyperlink',
    'footnoteReference',
    'endnoteReference',
  ])(
    'keeps %s objects available as context without destructive text replacement',
    async (element) => {
      const { range, context } = setup();
      range.getOoxml.mockReturnValue({
        value: wordXml(`<w:${element}/><w:p><w:r><w:t>Original</w:t></w:r></w:p>`),
      });
      const source = await (await createOfficeBridge('word')).readSelection();
      expect(source.text).toBe('Original');
      expect(source.replace).toBeUndefined();
      expect(source.warning).toContain('保留结构');
      expect(context.trackedObjects.add).not.toHaveBeenCalled();
    },
  );
  it('rechecks inserted objects even if the selected text did not change', async () => {
    const { range } = setup();
    const source = await (await createOfficeBridge('word')).readSelection();
    range.getOoxml.mockReturnValue({ value: wordXml('<w:drawing/>') });
    await expect(source.replace!('Replacement')).rejects.toThrow('保留结构');
    expect(range.insertText).not.toHaveBeenCalled();
  });
  it('inspects namespace-aware body content and fails safely on broken OOXML', () => {
    expect(
      wordSelectionWarning(wordXml().replace(/w:/g, 'doc:').replace('xmlns:w', 'xmlns:doc')),
    ).toBeUndefined();
    expect(wordSelectionWarning('<broken')).toContain('无法确认');
    expect(
      wordSelectionWarning(
        wordXml('<m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"/>'),
      ),
    ).toContain('保留结构');
    expect(
      wordSelectionWarning(
        wordXml().replace('</w:document>', '<w:styles><w:tbl/></w:styles></w:document>'),
      ),
    ).toBeUndefined();
  });
});

describe('Excel structured table creation', () => {
  function setup() {
    office(['ExcelApi:1.1', 'ExcelApi:1.2']);
    const range = {
      formulas: [] as unknown[][],
      numberFormat: [] as string[][],
      text: [
        ['A', 'B'],
        ['1', '2'],
      ],
      format: { wrapText: false, columnWidth: 0, autofitRows: vi.fn() },
      load: vi.fn(),
    };
    const nativeTable = { style: '' };
    const newSheet = {
      getRange: vi.fn(() => range),
      tables: { add: vi.fn(() => nativeTable) },
      activate: vi.fn(),
    };
    const sheets = {
      items: [{ name: '数据' }, { name: '数据 (2)' }],
      load: vi.fn(),
      add: vi.fn(() => newSheet),
    };
    const context = {
      workbook: {
        worksheets: sheets,
        application: { load: vi.fn(), calculationMode: 'Automatic', calculationState: 'Done' },
      },
      sync: vi.fn().mockResolvedValue(undefined),
    };
    vi.stubGlobal('Excel', {
      run: (callback: (ctx: typeof context) => unknown) => callback(context),
    });
    return { range, newSheet, sheets, context, nativeTable };
  }
  const raw = JSON.stringify({
    name: '数据',
    headers: ['项目', '值'],
    formats: ['text', 'number'],
    rows: [['001', { formula: '=1+1' }]],
  });
  it('previews a unique name and writes only after commit to a new worksheet', async () => {
    const { sheets, newSheet, range, nativeTable } = setup();
    const plan = await (await createOfficeBridge('excel')).prepareWrite('table', undefined, raw);
    expect(plan.label).toContain('数据 (3)');
    expect(plan.label).toContain('A1:B2');
    expect(sheets.add).not.toHaveBeenCalled();
    await plan.commit(raw);
    expect(sheets.add).toHaveBeenCalledWith('数据 (3)');
    expect(newSheet.getRange).toHaveBeenCalledWith('A1:B2');
    expect(range.formulas).toEqual([
      ["'项目", "'值"],
      ["'001", '=1+1'],
    ]);
    expect(range.numberFormat).toEqual([
      ['@', '@'],
      ['@', '#,##0.00'],
    ]);
    expect(newSheet.tables.add).toHaveBeenCalledWith('A1:B2', true);
    expect(nativeTable.style).toBe('TableStyleMedium2');
    expect(newSheet.activate).toHaveBeenCalledTimes(1);
  });
  it('rejects a name collision or changed payload before creating any worksheet', async () => {
    const { sheets } = setup();
    const plan = await (await createOfficeBridge('excel')).prepareWrite('table', undefined, raw);
    await expect(plan.commit(raw.replace('001', '002'))).rejects.toThrow('确认时');
    sheets.items.push({ name: '数据 (3)' });
    await expect(plan.commit(raw)).rejects.toThrow('名称已被使用');
    expect(sheets.add).not.toHaveBeenCalled();
  });
  it('reports calculated formula errors while acknowledging the sheet was created', async () => {
    const { range } = setup();
    range.text[1][1] = '#DIV/0!';
    const plan = await (await createOfficeBridge('excel')).prepareWrite('table', undefined, raw);
    expect(await plan.commit(raw)).toMatchObject({
      warning: expect.stringContaining('B2 #DIV/0!'),
    });
  });
  it('does not claim validation or change workbook mode when Excel calculation is manual or pending', async () => {
    const { context } = setup();
    context.workbook.application.calculationMode = 'Manual';
    const plan = await (await createOfficeBridge('excel')).prepareWrite('table', undefined, raw);
    expect(await plan.commit(raw)).toMatchObject({
      warning: expect.stringContaining('手动计算模式'),
    });
    expect(context.workbook.application.calculationMode).toBe('Manual');
    office(['ExcelApi:1.1', 'ExcelApi:1.2', 'ExcelApi:1.9']);
    context.workbook.application.calculationMode = 'Automatic';
    context.workbook.application.calculationState = 'Pending';
    const pendingPlan = await (
      await createOfficeBridge('excel')
    ).prepareWrite('table', undefined, raw);
    expect(await pendingPlan.commit(raw)).toMatchObject({
      warning: expect.stringContaining('待完成的计算'),
    });
  });
  it('names the partial worksheet on a write failure and does not retry automatically', async () => {
    const { context, sheets } = setup();
    const plan = await (await createOfficeBridge('excel')).prepareWrite('table', undefined, raw);
    context.sync
      .mockResolvedValueOnce(undefined)
      .mockResolvedValueOnce(undefined)
      .mockRejectedValueOnce(new Error('write failed'));
    await expect(plan.commit(raw)).rejects.toThrow('数据 (3)');
    expect(sheets.add).toHaveBeenCalledTimes(1);
  });
  it('creates the complete 48-row fixture and gates formatting on ExcelApi 1.2', async () => {
    const { range, newSheet } = setup();
    range.text = Array.from({ length: 49 }, () => Array(8).fill('0'));
    const source = JSON.stringify(COMPLEX_TABLE);
    const plan = await (await createOfficeBridge('excel')).prepareWrite('table', undefined, source);
    await plan.commit(source);
    expect(newSheet.getRange).toHaveBeenCalledWith('A1:H49');
    expect(range.formulas).toHaveLength(49);
    expect(range.formulas[48][5]).toBe('=D49-E49');
    office(['ExcelApi:1.1']);
    const legacy = await createOfficeBridge('excel');
    expect(legacy.capabilities.table).toBe(false);
    await expect(legacy.prepareWrite('table', undefined, source)).rejects.toThrow('不支持');
  });
});

describe('Excel single-cell write protection', () => {
  function setup() {
    office(['ExcelApi:1.1']);
    const range = {
      address: 'Sheet1!C7',
      rowCount: 1,
      columnCount: 1,
      formulas: [['=1+1']],
      values: [[2]],
      text: [['2']],
      load: vi.fn(),
      worksheet: { id: 'sheet-id', name: 'Sheet1', load: vi.fn() },
    };
    const getRange = vi.fn(() => range);
    const context = {
      workbook: {
        getSelectedRange: () => range,
        worksheets: { getItem: vi.fn(() => ({ getRange })) },
      },
      sync: vi.fn().mockResolvedValue(undefined),
    };
    vi.stubGlobal('Excel', {
      run: (cb: (ctx: typeof context) => unknown) => cb(context),
    });
    return { range, context, getRange };
  }
  it('uses the captured worksheet ID and address even if selection moves', async () => {
    const { range, context, getRange } = setup();
    const plan = await (await createOfficeBridge('excel')).prepareWrite('formula');
    expect(plan.label).toContain('Sheet1!C7');
    context.workbook.getSelectedRange = () => ({
      ...range,
      address: 'Sheet1!D8',
    });
    await plan.commit('=SUM(A1:A3)');
    expect(context.workbook.worksheets.getItem).toHaveBeenCalledWith('sheet-id');
    expect(getRange).toHaveBeenCalledWith('C7');
    expect(range.formulas).toEqual([['=SUM(A1:A3)']]);
  });
  it('refuses multi-cell targets, oversized reads and changed values', async () => {
    const { range } = setup();
    const bridge = await createOfficeBridge('excel');
    range.rowCount = 2;
    await expect(bridge.prepareWrite('formula')).rejects.toThrow('一个目标单元格');
    range.rowCount = 3000;
    await expect(bridge.readSelection()).rejects.toThrow('2,000');
    range.rowCount = 1;
    const plan = await bridge.prepareWrite('formula');
    range.values = [[3]];
    await expect(plan.commit('=3+3')).rejects.toThrow('已经变化');
    expect(range.formulas).toEqual([['=1+1']]);
  });
  it('retains numeric precision, formatted dates and percentages, and multiline cell boundaries', async () => {
    const { range } = setup();
    Object.assign(range, {
      address: "'区域 销售'!A1:D2",
      rowCount: 2,
      columnCount: 4,
      values: [
        ['单价', '完成率', '日期', '备注'],
        [1234.5678, 0.12567, 46000, '含\t制表符\n第二行'],
      ],
      text: [
        ['单价', '完成率', '日期', '备注'],
        ['¥1,234.57', '12.6%', '2025-12-09', '含\t制表符\n第二行'],
      ],
      formulas: [
        ['单价', '完成率', '日期', '备注'],
        [1234.5678, '=1/8+0.00067', 46000, '含\t制表符\n第二行'],
      ],
    });
    const source = await (await createOfficeBridge('excel')).readSelection();
    const content = JSON.parse(source.text);
    expect(content.values[1]).toEqual([1234.5678, 0.12567, 46000, '含\t制表符\n第二行']);
    expect(content.displayValues[1][1]).toBe('12.6%');
    expect(content.formulas[1][1]).toBe('=1/8+0.00067');
    expect(source.label).toContain('2 行 × 4 列');
  });
  it('reads formulas with empty displayed results and shows zero or false before overwriting', async () => {
    const { range } = setup();
    Object.assign(range, { values: [['']], text: [['']], formulas: [['=IF(A1="","",A1)']] });
    const bridge = await createOfficeBridge('excel');
    expect(JSON.parse((await bridge.readSelection()).text).formulas[0][0]).toContain('=IF');
    for (const value of [0, false]) {
      Object.assign(range, { values: [[value]], text: [['']], formulas: [[value]] });
      expect((await bridge.prepareWrite('formula')).before).toBe(String(value));
      expect(JSON.parse((await bridge.readSelection()).text).values[0][0]).toBe(value);
    }
  });
  it('reads exactly 2,000 cells without whitespace exhausting the context budget', async () => {
    const { range } = setup();
    Object.assign(range, {
      rowCount: 200,
      columnCount: 10,
      address: 'Sheet1!A1:J200',
      values: Array.from({ length: 200 }, () => Array(10).fill(1)),
      text: Array.from({ length: 200 }, () => Array(10).fill('1')),
      formulas: Array.from({ length: 200 }, () => Array(10).fill(1)),
    });
    const source = await (await createOfficeBridge('excel')).readSelection();
    expect(JSON.parse(source.text).values.flat()).toHaveLength(2000);
    expect(source.text.length).toBeLessThan(30000);
  });
});

describe('PowerPoint outline write boundaries', () => {
  function setup(modern = false) {
    office(['PowerPointApi:1.4', ...(modern ? ['PowerPointApi:1.10'] : [])]);
    const existingDelete = vi.fn();
    const newPlaceholderDeletes: ReturnType<typeof vi.fn>[] = [];
    const textBoxes: {
      text: string;
      options: unknown;
      textFrame: {
        textRange: { font: Record<string, unknown> };
        autoSizeSetting?: string;
      };
    }[] = [];
    const slides = [
      {
        id: 'existing-slide',
        shapes: {
          items: [{ delete: existingDelete }],
          load: vi.fn(),
          addTextBox: vi.fn(),
        },
      },
    ];
    const collection = {
      get items() {
        return [...slides];
      },
      load: vi.fn(),
      getCount: () => ({ value: slides.length }),
      add: vi.fn(() => {
        const remove = vi.fn();
        newPlaceholderDeletes.push(remove);
        slides.push({
          id: `new-${slides.length}`,
          shapes: {
            items: [{ delete: remove }],
            load: vi.fn(),
            addTextBox: vi.fn((text, options) => {
              const shape = {
                text,
                options,
                textFrame: { textRange: { font: {} } },
              };
              textBoxes.push(shape);
              return shape;
            }),
          },
        });
      }),
      getItemAt: (index: number) => slides[index],
      getItem: vi.fn((id: string) => slides.find((slide) => slide.id === id)!),
    };
    const context = {
      presentation: {
        slides: collection,
        pageSetup: { slideWidth: 540, slideHeight: 960, load: vi.fn() },
      },
      sync: vi.fn().mockResolvedValue(undefined),
    };
    vi.stubGlobal('PowerPoint', {
      run: (callback: (ctx: typeof context) => unknown) => callback(context),
    });
    return { slides, collection, context, existingDelete, textBoxes, newPlaceholderDeletes };
  }
  it('validates before adding slides and leaves existing content untouched', async () => {
    const { slides, collection, existingDelete, textBoxes, newPlaceholderDeletes } = setup();
    const plan = await (await createOfficeBridge('powerpoint')).prepareWrite('slides');
    await expect(plan.commit('{"slides":[]}')).rejects.toThrow();
    expect(collection.add).not.toHaveBeenCalled();
    await plan.commit(
      JSON.stringify({
        slides: [
          { title: 'A', bullets: ['One'] },
          { title: 'B', bullets: ['Two'] },
        ],
      }),
    );
    expect(slides).toHaveLength(3);
    expect(existingDelete).not.toHaveBeenCalled();
    expect(textBoxes.map((box) => box.text)).toEqual(['A', '• One', 'B', '• Two']);
    expect(
      textBoxes.every((box) => box.textFrame.autoSizeSetting === 'AutoSizeTextToFitShape'),
    ).toBe(true);
    expect(newPlaceholderDeletes.every((remove) => remove.mock.calls.length === 1)).toBe(true);
  });
  it('stops before deleting any shapes if another edit changes the page order during creation', async () => {
    const { slides, context, existingDelete, newPlaceholderDeletes } = setup();
    let syncs = 0;
    context.sync.mockImplementation(async () => {
      if (++syncs === 2) slides.reverse();
    });
    const plan = await (await createOfficeBridge('powerpoint')).prepareWrite('slides');
    await expect(
      plan.commit(JSON.stringify({ slides: [{ title: '新页', bullets: ['要点'] }] })),
    ).rejects.toThrow('页序发生变化');
    expect(existingDelete).not.toHaveBeenCalled();
    expect(newPlaceholderDeletes[0]).not.toHaveBeenCalled();
  });
  it('reserves readable line spacing before native auto-fit for dense Chinese slides', async () => {
    const { context, textBoxes } = setup(true);
    context.presentation.pageSetup.slideWidth = 720;
    context.presentation.pageSetup.slideHeight = 540;
    const plan = await (await createOfficeBridge('powerpoint')).prepareWrite('slides');
    await plan.commit(
      JSON.stringify({
        slides: [
          { title: '验收范围', bullets: ['保留原始资料', '核对全部数据'] },
          {
            title: '复杂项目验收'.repeat(8),
            bullets: Array.from({ length: 5 }, () => '验收数据完整性'.repeat(21)),
          },
        ],
      }),
    );
    expect(textBoxes[0].textFrame.textRange.font.size).toBe(28);
    expect(textBoxes[1].textFrame.textRange.font.size).toBe(16);
    expect(textBoxes[2].textFrame.textRange.font.size).toBeGreaterThanOrEqual(18);
    expect(textBoxes[2].textFrame.textRange.font.size).toBeLessThan(28);
    expect(textBoxes[3].textFrame.textRange.font.size).toBe(12);
    expect(textBoxes[3].textFrame.autoSizeSetting).toBe('AutoSizeTextToFitShape');
  });
  it('creates twelve dense slides by ID within a portrait presentation', async () => {
    const { context, textBoxes, collection, existingDelete } = setup(true);
    const plan = await (await createOfficeBridge('powerpoint')).prepareWrite('slides');
    await plan.commit(
      JSON.stringify({
        slides: Array.from({ length: 12 }, (_, i) => ({
          title: `第 ${i + 1} 页`,
          bullets: Array(5).fill('长内容'.repeat(50)),
        })),
      }),
    );
    expect(textBoxes).toHaveLength(24);
    expect(collection.getItem).toHaveBeenCalledTimes(12);
    for (const box of textBoxes) {
      const rect = box.options as { left: number; top: number; width: number; height: number };
      expect(rect.left + rect.width).toBeLessThanOrEqual(context.presentation.pageSetup.slideWidth);
      expect(rect.top + rect.height).toBeLessThanOrEqual(
        context.presentation.pageSetup.slideHeight,
      );
    }
    expect(existingDelete).not.toHaveBeenCalled();
  });
});
