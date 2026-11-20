import { afterEach, describe, expect, it, vi } from 'vitest';
import { createOfficeBridge } from '../src/office/bridge';

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
});

describe('Word tracked range and stale content protection', () => {
  function setup() {
    office(['WordApi:1.1']);
    const range = { text: 'Original', load: vi.fn(), insertText: vi.fn() };
    const context = {
      document: { getSelection: () => range },
      sync: vi.fn().mockResolvedValue(undefined),
      trackedObjects: { add: vi.fn(), remove: vi.fn() },
    };
    vi.stubGlobal('Word', {
      InsertLocation: { replace: 'Replace', end: 'End' },
      run: vi.fn((...args) => args[args.length - 1](context)),
    });
    return { range, context };
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
});

describe('PowerPoint outline write boundaries', () => {
  it('validates before adding slides and leaves existing content untouched', async () => {
    office(['PowerPointApi:1.4']);
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
        shapes: {
          items: [{ delete: existingDelete }],
          load: vi.fn(),
          addTextBox: vi.fn(),
        },
      },
    ];
    const collection = {
      getCount: () => ({ value: slides.length }),
      add: vi.fn(() => {
        const remove = vi.fn();
        newPlaceholderDeletes.push(remove);
        slides.push({
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
    };
    const context = {
      presentation: { slides: collection },
      sync: vi.fn().mockResolvedValue(undefined),
    };
    vi.stubGlobal('PowerPoint', {
      run: (callback: (ctx: typeof context) => unknown) => callback(context),
    });
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
});
