import { describe, expect, it } from 'vitest';
import Pptx from 'pptxgenjs';
import { PptxWorkingCopy } from '../src/presentation/pptx-editor';
import {
  A,
  C,
  P,
  R,
  X,
  OfficePackage,
  unpack,
  bytes64,
  element,
  nodes,
} from '../src/presentation/package-xml';
import {
  localBackgroundRequest,
  applyRecipe,
  sampleStyle,
  slideIssues,
} from '../src/presentation/slide-tools';
import { mergeProposals } from '../src/presentation/page-edits';
import { renderDeck, DEFAULT_THEME } from '../src/presentation/render';
import type { ChartData, PageEdit } from '../src/presentation/types';
const data: ChartData = {
  categories: ['East', 'West', 'South'],
  series: [
    { name: 'Results', values: [0, null, 0.25] },
    { name: 'Target', values: [1, 2, 4] },
  ],
};
async function fixture(n = 2) {
  const deck = new Pptx();
  deck.layout = 'LAYOUT_WIDE';
  for (let i = 0; i < n; i++) {
    const slide = deck.addSlide();
    slide.addText('Q3 ' + (i + 1), {
      x: 1,
      y: 0.2,
      w: 8,
      h: 0.7,
      objectName: 'Title',
      color: '18324A',
      fontSize: 30,
    });
    slide.addText('Q3 overview.\nMore detail.', {
      x: 1,
      y: 1,
      w: 4,
      h: 1.5,
      objectName: 'Body',
      fontFace: 'Arial',
      fontSize: 18,
      color: 'FFFFFF',
    });
    if (i === 0)
      slide.addChart(
        deck.ChartType.bar,
        [{ name: 'Original', labels: ['A', 'B'], values: [1, 2] }],
        { x: 1, y: 3, w: 5, h: 3, valAxisLabelFormatCode: '0.0%' },
      );
  }
  return bytes64((await deck.write({ outputType: 'base64' })) as string);
}
async function readBlob(b: Blob) {
  return new Uint8Array(await b.arrayBuffer());
}
describe('native PPTX working copy', () => {
  it('reopens twenty slides with one exception, preserving original shape and chart parts', async () => {
    const original = await fixture(20),
      copy = new PptxWorkingCopy(original),
      snap = await copy.read();
    const plan = localBackgroundRequest('把整份PPT背景改为米白色，但第5页改为深蓝色', snap)!;
    expect(plan.actions).toHaveLength(20);
    await copy.apply(snap, plan);
    const saved = await readBlob(await copy.backup()),
      pkg = new OfficePackage(unpack(saved)),
      before = new OfficePackage(unpack(original));
    const now = await new PptxWorkingCopy(saved).read();
    expect(now.slides.map((s) => s.background)).toEqual(
      snap.slides.map((_, i) => ({ kind: 'solid', color: i === 4 ? '#17365D' : '#F5F1E8' })),
    );
    expect(now.slides.map((s) => s.shapes)).toEqual(snap.slides.map((s) => s.shapes));
    expect(pkg.files['ppt/charts/chart1.xml']).toEqual(before.files['ppt/charts/chart1.xml']);
    await expect(copy.apply(snap, plan)).rejects.toThrow();
  });
  it('combines recipe, chart refresh, scoped style matching, native gradient and insertion/split', async () => {
    const original = await fixture(),
      copy = new PptxWorkingCopy(original),
      snap = await copy.read(),
      [a, b] = snap.slides;
    const chart = a.shapes.find((s) => s.chart)!,
      body = b.shapes.find((s) => s.role === 'body')!;
    expect(chart.chart!.data.categories).toEqual(['A', 'B']);
    const recipe = applyRecipe(
      {
        name: 'Report',
        fill: { kind: 'gradient', from: '#17365D', to: '#6098B0' },
        titleColor: '#FFFFFF',
        find: 'Q3',
        replacement: 'Q4',
      },
      snap,
      [a],
    );
    const style = sampleStyle(snap, a, [b], {
      background: false,
      title: true,
      body: true,
      geometry: true,
      alignment: true,
    });
    const actions: PageEdit[] = [
      { kind: 'chart', slideId: a.id, shapeId: chart.id, data },
      {
        kind: 'insert',
        slideId: a.id,
        title: 'Follow-up',
        body: 'Forecast',
        chart: { kind: 'line', unit: '%', data },
      },
      {
        kind: 'split',
        slideId: b.id,
        shapeId: body.id,
        at: body.text!.indexOf('\n'),
        title: 'Details',
      },
    ];
    await copy.apply(
      snap,
      mergeProposals(mergeProposals(recipe, style), {
        summary: 'Refresh and split',
        changes: [],
        actions,
      }),
    );
    const bytes = await readBlob(await copy.backup()),
      now = await new PptxWorkingCopy(bytes).read(),
      pkg = new OfficePackage(unpack(bytes));
    expect(now.slides.map((s) => s.title)).toEqual(['Q4 1', 'Follow-up', 'Q3 2', 'Details']);
    expect(now.slides[0].shapes.find((s) => s.id === chart.id)).toMatchObject({
      left: chart.left,
      width: chart.width,
      chart: { data },
    });
    expect(now.slides[1].background).toEqual(now.slides[0].background);
    expect(now.slides[1].shapes.find((s) => s.chart)?.chart?.data).toEqual(data);
    expect(now.slides[2].shapes.find((s) => s.id === body.id)?.text).toBe('Q3 overview.');
    expect(now.slides[3].shapes.find((s) => s.id === body.id)?.text).toBe('More detail.');
    const firstPath = 'ppt/slides/slide1.xml',
      chartPath = pkg.target(
        firstPath,
        nodes(pkg.doc(firstPath), C, 'chart')[0].getAttributeNS(R, 'id'),
      )!,
      cd = pkg.doc(chartPath);
    expect(nodes(cd, C, 'numFmt').some((e) => e.getAttribute('formatCode') === '0.0%')).toBe(true);
    const embedded = pkg.target(
      chartPath,
      nodes(cd, C, 'externalData')[0].getAttributeNS(R, 'id'),
    )!;
    const sheet = new OfficePackage(unpack(pkg.files[embedded])).doc('xl/worksheets/sheet1.xml');
    const cells = Object.fromEntries(
      nodes(sheet, X, 'c').map((c) => [c.getAttribute('r'), c.textContent]),
    );
    expect(cells).toMatchObject({
      A2: 'East',
      B1: 'Results',
      B2: '0',
      B3: '',
      B4: '0.25',
      C3: '2',
    });
  });
  it('isolates shared chart and workbook parts before refreshing one slide', async () => {
    const pkg = new OfficePackage(unpack(await fixture())),
      one = 'ppt/slides/slide1.xml',
      two = 'ppt/slides/slide2.xml';
    const sourceFrame = nodes(pkg.doc(one), P, 'graphicFrame')[0],
      chartPath = pkg.target(one, nodes(sourceFrame, C, 'chart')[0].getAttributeNS(R, 'id'))!;
    const doc = pkg.edit(two),
      frame = doc.importNode(sourceFrame, true);
    nodes(frame, C, 'chart')[0].setAttributeNS(R, 'r:id', pkg.relate(two, chartPath, R + '/chart'));
    nodes(doc, P, 'spTree')[0].append(frame);
    const copy = new PptxWorkingCopy(pkg.pack()),
      before = await copy.read(),
      [first, second] = before.slides;
    const original = second.shapes.find((s) => s.chart)!;
    await copy.apply(before, {
      summary: 'Refresh only slide one',
      changes: [],
      actions: [
        { kind: 'chart', slideId: first.id, shapeId: first.shapes.find((s) => s.chart)!.id, data },
      ],
    });
    const after = await copy.read();
    expect(after.slides[0].shapes.find((s) => s.chart)?.chart?.data).toEqual(data);
    expect(after.slides[1].shapes.find((s) => s.chart)?.chart?.data).toEqual(original.chart!.data);
    const saved = new OfficePackage(unpack(await readBlob(await copy.backup()))),
      origWorkbook = pkg.target(
        chartPath,
        nodes(pkg.doc(chartPath), C, 'externalData')[0].getAttributeNS(R, 'id'),
      )!;
    expect(saved.files[chartPath]).toEqual(pkg.files[chartPath]);
    expect(saved.files[origWorkbook]).toEqual(pkg.files[origWorkbook]);
  });
  it('uses local background images, copies the original background, restores the master and reads comments', async () => {
    const picture =
      'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+j4hoAAAAASUVORK5CYII=';
    const ctx = {
      materials: [
        {
          id: 'pic',
          kind: 'image' as const,
          name: 'image.png',
          text: '',
          tables: [],
          image: { data: picture, modelData: picture, width: 1, height: 1 },
        },
      ],
      includeImages: false,
    };
    const base = await renderDeck(
      {
        title: 'Report',
        slides: [
          { title: 'First', layout: 'cover', sources: [], notes: '' },
          {
            title: 'Second',
            layout: 'closing',
            sources: [],
            notes: '',
            background: { kind: 'solid', color: '#FAFAFA' },
          },
        ],
      },
      ctx,
      { ...DEFAULT_THEME, fill: { kind: 'image', imageId: 'pic', transparency: 38 } },
    );
    const raw = new OfficePackage(unpack(bytes64(base))),
      comment = 'ppt/comments/comment1.xml',
      path = 'ppt/slides/slide1.xml';
    const doc = document.implementation.createDocument(P, 'p:cmLst'),
      c = element(doc, P, 'p:cm'),
      tx = element(doc, P, 'p:text');
    tx.textContent = 'Shorten the title';
    c.setAttribute('authorId', '0');
    c.setAttribute('idx', '1');
    c.setAttribute('dt', '2026-01-01T12:00:00Z');
    c.append(element(doc, P, 'p:pos', { x: 100, y: 100 }), tx);
    const authorPath = 'ppt/commentAuthors.xml',
      authors = document.implementation.createDocument(P, 'p:cmAuthorLst');
    authors.documentElement.append(
      element(authors, P, 'p:cmAuthor', {
        id: 0,
        name: 'Reviewer',
        initials: 'QA',
        lastIdx: 1,
        clrIdx: 0,
      }),
    );
    raw.saveDoc(authorPath, authors);
    raw.relate('ppt/presentation.xml', authorPath, R + '/commentAuthors');
    raw.type(
      authorPath,
      'application/vnd.openxmlformats-officedocument.presentationml.commentAuthors+xml',
    );
    doc.documentElement.append(c);
    raw.saveDoc(comment, doc);
    raw.relate(path, comment, R + '/comments');
    raw.type(comment, 'application/vnd.openxmlformats-officedocument.presentationml.comments+xml');
    const copy = new PptxWorkingCopy(raw.pack()),
      snap = await copy.read(),
      [first, second] = snap.slides;
    expect(slideIssues(snap).map((i) => i.message)).toContain('Shorten the title');
    expect(first.background).toMatchObject({ kind: 'image', transparency: 38 });
    await copy.apply(snap, {
      summary: 'Copy original',
      changes: [],
      actions: [
        { kind: 'background', slideId: first.id, fill: { kind: 'master' } },
        { kind: 'background', slideId: second.id, fill: { kind: 'reference', sourceId: first.id } },
      ],
    });
    const next = await copy.read();
    expect(next.slides[0].background).toEqual({ kind: 'master' });
    expect(next.slides[0].effectiveBackground).toEqual(first.effectiveBackground);
    expect(next.slides[1].background).toEqual(first.background);
    const saved = new OfficePackage(unpack(await readBlob(await copy.backup())));
    expect(nodes(saved.doc('ppt/slides/slide2.xml'), A, 'alphaModFix')[0].getAttribute('amt')).toBe(
      '62000',
    );
  });
});
