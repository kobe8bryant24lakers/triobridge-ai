import { afterEach, describe, expect, it, vi } from 'vitest';
import { createPresentationAccess } from '../src/office/presentation';

function office() {
  const writes: string[] = [];
  const load = () => {};
  function shape(id: string, value: string, top = 40, size = 22.5) {
    const font = new Proxy(
      { name: 'Arial', size, color: '#000000', bold: false, load },
      {
        set(target, key, value) {
          writes.push(`font:${String(key)}`);
          Reflect.set(target, key, value);
          return true;
        },
      },
    );
    const range = {
      text: value,
      font,
      load,
      paragraphFormat: { horizontalAlignment: 'Left', load },
      getSubstring(start: number, length: number) {
        return {
          set text(next: string) {
            writes.push('text');
            range.text = range.text.slice(0, start) + next + range.text.slice(start + length);
          },
        };
      },
    };
    return {
      id,
      name: `Text ${id}`,
      type: 'TextBox',
      left: 40,
      top,
      width: 600,
      height: 55,
      textFrame: { hasText: true, load, textRange: range },
    };
  }
  const heading = shape('title', '经营概览'),
    footer = shape('source', '来源：经营表', 500, 9);
  const slides = ['a', 'b', 'c'].map((id) => ({
    id,
    shapes: {
      items: id === 'b' ? [heading, footer] : [shape(`${id}-text`, `页 ${id}`)],
      load,
      getItem(shapeId: string) {
        return this.items.find((s) => s.id === shapeId)!;
      },
    },
    moveTo(index: number) {
      writes.push('order');
      const at = slides.indexOf(this);
      slides.splice(at, 1);
      slides.splice(index, 0, this);
    },
  }));
  const collection = {
    get items() {
      return slides;
    },
    load,
    getItem: (id: string) => slides.find((s) => s.id === id)!,
  };
  let partial = false,
    insertCount = 1;
  const ctx = {
    presentation: {
      slides: collection,
      pageSetup: { slideWidth: 720, slideHeight: 540, load },
      getSelectedSlides: vi.fn(() => ({ items: [slides.find((s) => s.id === 'b')!], load })),
      getSelectedShapes: vi.fn(() => ({ items: [heading], load })),
      insertSlidesFromBase64: () => {
        writes.push('insert');
        for (let i = 0; i < insertCount; i++) slides.push({ ...slides[0], id: `inserted${i}` });
      },
    },
    sync: async () => {
      if (partial && writes.length) throw new Error('Simulated disconnected host after write');
    },
  };
  vi.stubGlobal('Office', { context: { requirements: { isSetSupported: () => true } } });
  vi.stubGlobal('PowerPoint', {
    run: async (fn: (c: typeof ctx) => unknown) => fn(ctx),
    InsertSlideFormatting: { keepSourceFormatting: 'KeepSourceFormatting' },
  });
  return {
    access: createPresentationAccess(),
    writes,
    heading,
    footer,
    slides,
    presentation: ctx.presentation,
    partial: () => {
      partial = true;
    },
    wrongInsertCount: () => {
      insertCount = 2;
    },
  };
}
afterEach(() => vi.unstubAllGlobals());
describe('native PowerPoint bridge transaction boundaries', () => {
  it('reads selected objects, recognizes 4:3 titles and preserves small footers', async () => {
    const f = office(),
      snap = await f.access.read('selected');
    expect(snap.size).toEqual({ width: 720, height: 540 });
    expect(snap.slides.map((s) => s.id)).toEqual(['b']);
    expect(snap.slides[0].shapes.map((s) => s.role)).toEqual(['title', 'footer']);
    expect(snap.selectedShapeIds).toEqual(['title']);
    expect(f.writes).toEqual([]);
  });
  it('refuses a moved object before applying edits', async () => {
    const f = office(),
      snap = await f.access.read('selected');
    f.heading.left += 1;
    await expect(
      f.access.apply(snap, {
        summary: '改标题',
        changes: [{ slideId: 'b', shapeId: 'title', text: '更新概览' }],
      }),
    ).rejects.toThrow('尚未应用');
    expect(f.writes).toEqual([]);
  });
  it('appends without reading selection or objects, even after unrelated objects change', async () => {
    const f = office();
    f.presentation.getSelectedSlides.mockImplementation(() => {
      throw new Error('Invalid selection');
    });
    f.presentation.getSelectedShapes.mockImplementation(() => {
      throw new Error('Invalid selection');
    });
    for (const slide of f.slides)
      slide.shapes.load = () => {
        throw new Error('Unsupported object');
      };
    const target = await f.access.prepareInsert();
    expect(target).toEqual({ slideIds: ['a', 'b', 'c'], size: { width: 720, height: 540 } });
    f.heading.left += 1;
    await f.access.insert('test', 1, target);
    expect(f.writes).toEqual(['insert']);
    expect(f.presentation.getSelectedSlides).not.toHaveBeenCalled();
    expect(f.presentation.getSelectedShapes).not.toHaveBeenCalled();
  });
  it.each(['order', 'size'] as const)(
    'rejects a changed deck %s before insertion',
    async (change) => {
      const f = office(),
        target = await f.access.prepareInsert();
      if (change === 'order') f.slides.reverse();
      else f.presentation.pageSetup.slideWidth = 960;
      await expect(f.access.insert('test', 1, target)).rejects.toThrow('尚未插入');
      expect(f.writes).toEqual([]);
    },
  );
  it('prepares an empty presentation without requiring a selected slide', async () => {
    const f = office();
    f.slides.length = 0;
    const target = await f.access.prepareInsert();
    expect(target.slideIds).toEqual([]);
    await f.access.insert('test', 1, target);
    expect(f.slides.map((slide) => slide.id)).toEqual(['inserted0']);
    expect(f.presentation.getSelectedSlides).not.toHaveBeenCalled();
  });
  it('writes allowed text/style changes, then verifies data against a fresh read', async () => {
    const f = office(),
      snap = await f.access.read('selected');
    await f.access.apply(snap, {
      summary: '更新标题',
      changes: [
        {
          slideId: 'b',
          shapeId: 'title',
          text: '季度经营概览',
          fontSize: 30,
          fontColor: '#28604A',
        },
      ],
    });
    const after = await f.access.read('all');
    expect(after.slides[1].shapes[0].text).toBe('季度经营概览');
    expect(after.slides[1].shapes[0].font?.size).toBe(30);
    expect(after.slides[1].shapes[1]).toEqual(snap.slides[0].shapes[1]);
    expect(after.slideIds).toEqual(['a', 'b', 'c']);
  });
  it('reports possible partial application when Office fails after mutations begin', async () => {
    const f = office(),
      snap = await f.access.read('selected');
    f.partial();
    await expect(
      f.access.apply(snap, {
        summary: 'x',
        changes: [{ slideId: 'b', shapeId: 'title', fontSize: 30 }],
      }),
    ).rejects.toThrow('可能已应用部分改动');
    expect(f.writes).toEqual(['font:size']);
  });
  it('verifies appended IDs/count and detects ambiguous insertion results', async () => {
    const f = office();
    await f.access.insert('test', 1, await f.access.prepareInsert());
    expect(f.slides.map((s) => s.id)).toEqual(['a', 'b', 'c', 'inserted0']);
    const bad = office();
    bad.wrongInsertCount();
    await expect(bad.access.insert('test', 1, await bad.access.prepareInsert())).rejects.toThrow(
      '不要直接重复插入',
    );
  });
  it('reorders the whole deck and verifies IDs independently of the old indexes', async () => {
    const f = office(),
      snap = await f.access.read('all');
    await f.access.apply(snap, { summary: '页序', changes: [], slideOrder: ['c', 'a', 'b'] });
    expect(f.slides.map((s) => s.id)).toEqual(['c', 'a', 'b']);
  });
});
