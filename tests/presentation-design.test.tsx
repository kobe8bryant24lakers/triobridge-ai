import { describe, expect, it, vi } from 'vitest';
import { render, screen } from '@testing-library/react';
import userEvent from '@testing-library/user-event';
import Pptx from 'pptxgenjs';
import { ThemeDesigner } from '../src/components/ThemeDesigner';
import { SlideRefinement } from '../src/components/SlideRefinement';
import { DeckReview } from '../src/components/DeckReview';
import { DraftLayout } from '../src/components/DraftLayout';
import { builtInThemes, parseThemes, contrast } from '../src/presentation/design-themes';
import { optimizeLayout, swapPictureColumns } from '../src/presentation/layout-optimization';
import { PptxWorkingCopy } from '../src/presentation/pptx-editor';
import { bytes64, unpack } from '../src/presentation/package-xml';
import type { DeckContent, PresentationSnapshot } from '../src/presentation/types';
import type { PresentationRuntime } from '../src/components/presentation-runtime';

const material = { materials: [], includeImages: false };
const deck: DeckContent = {
  title: 'Review',
  slides: [
    {
      title: '已有标题',
      layout: 'bullets',
      bullets: ['甲', '乙', '丙'],
      notes: '保留备注',
      sources: [],
    },
  ],
};
const runtime = (reply: string) =>
  ({
    busy: false,
    demo: false,
    notify: vi.fn(),
    run: async (_label: string, work: (s: AbortSignal) => Promise<void>) =>
      work(new AbortController().signal),
    request: vi.fn(async () => reply),
    context: material,
    modelStamp: 'x',
  }) satisfies PresentationRuntime;
describe('post-generation editing', () => {
  it('compares designs with real titles, reuses a favorite and sends only five titles for AI design', async () => {
    const user = userEvent.setup(),
      apply = vi.fn(),
      themes = builtInThemes(),
      r = runtime(JSON.stringify({ themes }));
    render(
      <ThemeDesigner
        runtime={r}
        samples={Array.from({ length: 7 }, (_, i) => ({
          title: '页' + i,
          detail: '不可发送的正文',
        }))}
        onApply={apply}
      />,
    );
    await user.click(screen.getByRole('button', { name: '快速换主题' }));
    expect(screen.getAllByRole('radio')).toHaveLength(3);
    await user.type(screen.getByLabelText('风格要求'), '夏季发布');
    await user.click(screen.getByRole('button', { name: '按描述生成三套' }));
    const request = JSON.stringify((r.request.mock.calls as unknown[][])[0][0]);
    expect(request).not.toContain('不可发送的正文');
    expect(request).toContain('页4');
    expect(request).not.toContain('页5');
    await user.click(screen.getByRole('radio', { name: themes[1].name }));
    await user.click(screen.getByRole('button', { name: '保存此主题' }));
    expect(apply).not.toHaveBeenCalled();
    await user.click(screen.getByText('微调所选主题'));
    await user.selectOptions(screen.getByLabelText('标题字体'), 'Georgia');
    await user.click(screen.getByRole('button', { name: '应用到全部 7 页' }));
    expect(apply).toHaveBeenCalledWith({ ...themes[1], headingFont: 'Georgia' });
    expect(() =>
      parseThemes(JSON.stringify({ themes: [themes[0], themes[0], themes[0]] })),
    ).toThrow();
    expect(contrast(themes[1].background, themes[1].ink)).toBeGreaterThan(4.5);
  });
  it('applies only approved edits on one slide, and surfaces missing evidence', async () => {
    const u = userEvent.setup(),
      apply = vi.fn(),
      select = vi.fn(),
      r = runtime('{"edits":[{"id":"title","value":"新标题"},{"id":"b0","value":"新的要点"}]}');
    render(
      <>
        <SlideRefinement
          page={deck.slides[0]}
          index={0}
          runtime={r}
          disabled={false}
          onApply={apply}
        />
        <DeckReview deck={deck} busy={false} onSelect={select} />
      </>,
    );
    await u.click(screen.getByRole('button', { name: '优化本页内容' }));
    await u.click(screen.getByRole('button', { name: '标题突出结论' }));
    await u.click(screen.getByRole('button', { name: '生成本页建议' }));
    await u.click(screen.getByRole('checkbox', { name: /要点/ }));
    await u.click(screen.getByRole('button', { name: '采用所选 1 处修改' }));
    expect(apply).toHaveBeenCalledWith({ ...deck.slides[0], title: '新标题' });
    await u.click(screen.getByRole('button', { name: '检查成稿' }));
    await u.click(screen.getByRole('button', { name: '1. 已有标题' }));
    expect(select).toHaveBeenCalledWith(0);
  });
  it('previews a new native layout without rewriting the source content', async () => {
    const u = userEvent.setup(),
      apply = vi.fn();
    render(
      <DraftLayout
        busy={false}
        deck={deck}
        theme={builtInThemes()[0]}
        context={material}
        onApply={apply}
      />,
    );
    await u.click(screen.getByRole('button', { name: '优化布局' }));
    expect(screen.getByLabelText('重排效果预览')).toBeInTheDocument();
    expect(apply).not.toHaveBeenCalled();
    await u.click(screen.getByRole('button', { name: '应用所选 1 页布局' }));
    expect(apply).toHaveBeenCalledWith({
      ...deck,
      slides: [{ ...deck.slides[0], arrangement: 'bulletsInColumns' }],
    });
  });
  it('restyles working copies transactionally and preserves native chart data, notes and original bytes', async () => {
    const p = new Pptx();
    const slide = p.addSlide();
    slide.addText('Original', { x: 0.6, y: 0.2, w: 7, h: 0.8, objectName: 'Title' });
    slide.addChart(p.ChartType.bar, [{ name: 'Data', labels: ['A', 'B'], values: [3.95, 0.24] }], {
      x: 1,
      y: 1,
      w: 5,
      h: 4,
    });
    slide.addNotes('Original notes');
    const bytes = bytes64((await p.write({ outputType: 'base64' })) as string),
      original = new PptxWorkingCopy(bytes),
      before = await original.read(),
      style = builtInThemes()[1];
    const copy = original.themed(before, style),
      after = await copy.read();
    expect(after.slides[0].effectiveBackground).toEqual({ kind: 'solid', color: style.background });
    expect(after.slides[0].shapes.find((s) => s.chart)?.chart).toEqual(
      before.slides[0].shapes.find((s) => s.chart)?.chart,
    );
    const dst = unpack(new Uint8Array(await (await copy.backup()).arrayBuffer())),
      src = unpack(bytes);
    for (const name of Object.keys(src).filter(
      (n) => n.endsWith('.xlsx') || /notesSlide1\.xml$/.test(n),
    ))
      expect(dst[name]).toEqual(src[name]);
    expect(new Uint8Array(await (await original.backup()).arrayBuffer())).toEqual(bytes);
  });
  it('proposes only bounded geometry and swaps an intact picture/text pair', () => {
    const base = { type: 'Shape', name: 'N', width: 40, height: 14 };
    const snap: PresentationSnapshot = {
      scope: 'all',
      size: { width: 200, height: 100 },
      slideIds: ['a'],
      selectedShapeIds: [],
      skipped: 0,
      slides: [
        {
          id: 'a',
          index: 0,
          title: 'Title',
          shapes: [
            { ...base, id: 'p', type: 'Image', left: 15, top: 25, width: 65, height: 48 },
            { ...base, id: 'b', role: 'body', text: 'Text', left: 110, top: 34 },
            { ...base, id: 'x', role: 'footer', text: 'Footer', left: 8, top: 88 },
          ],
        },
      ],
    };
    const swap = swapPictureColumns(snap);
    expect(swap.changes).toEqual([
      { slideId: 'a', shapeId: 'p', left: 120 },
      { slideId: 'a', shapeId: 'b', left: 50 },
    ]);
    snap.slides[0].shapes[1].left = 174;
    expect(optimizeLayout(snap).changes).toEqual([{ slideId: 'a', shapeId: 'b', left: 160 }]);
  });
});
