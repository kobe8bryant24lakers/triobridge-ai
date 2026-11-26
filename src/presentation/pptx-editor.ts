import { t } from '../i18n';
import type {
  EditProposal,
  PageEdit,
  PageSize,
  PresentationAccess,
  PresentationSnapshot,
  ShapeState,
  SlideState,
} from './types';
import { assertSnapshotUnchanged, originalValue, reorderedIds, validateEdit } from './editing';
import {
  A,
  C,
  P,
  R,
  OfficePackage,
  base64,
  bytes64,
  child,
  element,
  nodes,
  unpack,
  relPath,
} from './package-xml';
import { copyBackground, readBackground, writeBackground } from './package-background';
import { refreshChart } from './package-charts';
import {
  chartPart,
  emu,
  identity,
  idOf,
  readShape,
  shapeNodes,
  shapeTree,
  writeShape,
} from './package-shapes';
import { addNativeChart, DEFAULT_THEME } from './render';
import { preserveChartData } from './chart-workbooks';

const mime = 'application/vnd.openxmlformats-officedocument.presentationml.presentation';
const unsupported = () => {
  throw Error(t('请在工作副本中使用页面工具。', 'Use the slide tools in this working copy.'));
};
function mainPart(pkg: OfficePackage) {
  const rel = pkg.rels('').find((r) => r.getAttribute('Type') === R + '/officeDocument');
  const path = rel && pkg.target('', rel.getAttribute('Id'));
  if (
    !path ||
    pkg.doc(path).documentElement.namespaceURI !== P ||
    pkg.contentType(path) !==
      'application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml'
  )
    throw Error(t('请选择可编辑的 .pptx 演示稿。', 'Choose an editable .pptx presentation.'));
  return path;
}
function slideEntries(pkg: OfficePackage, main: string) {
  const list = nodes(pkg.doc(main), P, 'sldIdLst')[0];
  return list
    ? Array.from(list.children)
        .filter((e) => e.namespaceURI === P && e.localName === 'sldId')
        .map((e) => {
          const path = pkg.target(main, e.getAttributeNS(R, 'id'));
          if (!path || !pkg.has(path))
            throw Error(t('幻灯片引用无效。', 'Invalid slide relationship.'));
          return { id: e.getAttribute('id')!, path, entry: e };
        })
    : [];
}
function slideComments(pkg: OfficePackage, path: string) {
  return pkg
    .rels(path)
    .filter((r) => /\/comments?$/i.test(r.getAttribute('Type') || ''))
    .flatMap((rel) => {
      const cp = pkg.target(path, rel.getAttribute('Id'));
      if (!cp) return [];
      const cm = Array.from(pkg.doc(cp).getElementsByTagNameNS('*', 'cm'));
      return cm
        .filter((c) => !['resolved', 'closed'].includes(c.getAttribute('status') || ''))
        .flatMap((c, i) => {
          const anchor = Array.from(c.getElementsByTagName('*')).find(
              (e) => e.localName === 'spMk',
            ),
            shapeId = anchor?.getAttribute('id');
          const replies = Array.from(c.getElementsByTagNameNS('*', 'reply'));
          return [c, ...replies]
            .filter((e) => !['resolved', 'closed'].includes(e.getAttribute('status') || ''))
            .map((entry, j) => {
              const legacy = child(entry, P, 'text'),
                body = Array.from(entry.children).find((e) => e.localName === 'txBody');
              const text =
                legacy?.textContent ||
                (body
                  ? nodes(body, A, 'p')
                      .map((p) =>
                        nodes(p, A, 't')
                          .map((e) => e.textContent)
                          .join(''),
                      )
                      .join('\n')
                  : '');
              return {
                id: cp + '#' + (entry.getAttribute('id') || i + '/' + j),
                text,
                ...(shapeId ? { shapeId } : {}),
              };
            })
            .filter((entry) => entry.text.trim());
        });
    });
}
function readPackage(pkg: OfficePackage): PresentationSnapshot {
  const main = mainPart(pkg),
    entries = slideEntries(pkg, main),
    sz = nodes(pkg.doc(main), P, 'sldSz')[0];
  if (!entries.length || entries.length > 500)
    throw Error(t('工作副本支持 1–500 页。', 'Working copies support 1–500 slides.'));
  const width = Number(sz?.getAttribute('cx')) / emu,
    height = Number(sz?.getAttribute('cy')) / emu;
  const slides = entries.map(({ id, path }, index): SlideState => {
    const shapes = shapeNodes(pkg.doc(path)).map((e) => readShape(pkg, path, e));
    if (!shapes.some((s) => s.role === 'title')) {
      const first = shapes.find(
        (s) => s.text?.trim() && s.role !== 'footer' && s.top < height * 0.3,
      );
      if (first) first.role = 'title';
    }
    return {
      id,
      index,
      shapes,
      title:
        shapes.find((s) => s.role === 'title')?.text ||
        shapes.find((s) => s.text)?.text?.slice(0, 80) ||
        t('第 ' + (index + 1) + ' 页', 'Slide ' + (index + 1)),
      ...readBackground(pkg, path),
      comments: slideComments(pkg, path),
    };
  });
  return {
    scope: 'all',
    slides,
    slideIds: entries.map((s) => s.id),
    size: width > 0 && height > 0 ? { width, height } : null,
    selectedShapeIds: [],
    skipped: slides.reduce(
      (sum, s) =>
        sum +
        s.shapes.filter((o) => ['Group', 'Graphic', 'Chart'].includes(o.type) && !o.chart).length,
      0,
    ),
  };
}
function textShape(d: Document, id: string, role: 'title' | 'body', size: PageSize, text: string) {
  const e = element(d, P, 'p:sp'),
    nv = element(d, P, 'p:nvSpPr');
  nv.append(
    element(d, P, 'p:cNvPr', {
      id,
      name: role === 'title' ? 'TrioBridge Title' : 'TrioBridge Content',
    }),
    element(d, P, 'p:cNvSpPr', { txBox: 1 }),
    element(d, P, 'p:nvPr'),
  );
  const sp = element(d, P, 'p:spPr'),
    xf = element(d, A, 'a:xfrm');
  xf.append(
    element(d, A, 'a:off', {
      x: Math.round(size.width * 0.06 * emu),
      y: Math.round(size.height * (role === 'title' ? 0.06 : 0.25) * emu),
    }),
    element(d, A, 'a:ext', {
      cx: Math.round(size.width * 0.88 * emu),
      cy: Math.round(size.height * (role === 'title' ? 0.14 : 0.57) * emu),
    }),
  );
  sp.append(xf);
  const b = element(d, P, 'p:txBody');
  b.append(element(d, A, 'a:bodyPr'), element(d, A, 'a:lstStyle'));
  e.append(nv, sp, b);
  writeShape(
    e,
    { slideId: '', shapeId: id, text, fontSize: role === 'title' ? 30 : 20 },
    {} as ShapeState,
  );
  return e;
}
async function addChart(
  pkg: OfficePackage,
  path: string,
  tree: Element,
  id: string,
  chart: NonNullable<Extract<PageEdit, { kind: 'insert' }>['chart']>,
  size: PageSize,
  withBody: boolean,
) {
  const { default: Pptx } = await import('pptxgenjs'),
    deck = new Pptx(),
    slide = deck.addSlide();
  addNativeChart(
    deck,
    slide,
    chart.data,
    chart.kind,
    {
      x: (size.width * (withBody ? 0.43 : 0.06)) / 72,
      y: (size.height * 0.24) / 72,
      w: (size.width * (withBody ? 0.51 : 0.88)) / 72,
      h: (size.height * 0.63) / 72,
    },
    DEFAULT_THEME,
    chart.unit,
  );
  const raw = preserveChartData(
    (await deck.write({ outputType: 'base64', compression: true })) as string,
    [chart.data],
  );
  const source = new OfficePackage(unpack(bytes64(raw))),
    srcSlide = 'ppt/slides/slide1.xml',
    frame = nodes(source.doc(srcSlide), P, 'graphicFrame')[0],
    srcChart = chartPart(source, srcSlide, frame)!;
  const dest = pkg.unique('ppt/charts', 'chart', '.xml'),
    cd = source.doc(srcChart),
    external = nodes(cd, C, 'externalData')[0],
    srcBook = source.target(srcChart, external.getAttributeNS(R, 'id'))!,
    book = pkg.unique('ppt/embeddings', 'Workbook', '.xlsx');
  pkg.files[book] = source.files[srcBook];
  pkg.type(book, source.contentType(srcBook)!);
  external.setAttributeNS(R, 'r:id', pkg.relate(dest, book, R + '/package'));
  pkg.saveDoc(dest, cd);
  pkg.type(dest, source.contentType(srcChart)!);
  const shape = tree.ownerDocument.importNode(frame, true);
  identity(shape).setAttribute('id', id);
  nodes(shape, C, 'chart')[0].setAttributeNS(R, 'r:id', pkg.relate(path, dest, R + '/chart'));
  tree.append(shape);
}
function attachSlide(pkg: OfficePackage, main: string, path: string, afterId: string) {
  const doc = pkg.edit(main),
    entries = slideEntries(pkg, main),
    after = entries.find((s) => s.id === afterId)!.entry,
    id = String(Math.max(255, ...entries.map((e) => Number(e.id))) + 1);
  const entry = element(doc, P, 'p:sldId', { id });
  entry.setAttributeNS(R, 'r:id', pkg.relate(main, path, R + '/slide'));
  after.parentElement!.insertBefore(entry, after.nextSibling);
  const section = Array.from(doc.getElementsByTagNameNS('*', 'sldId')).find(
    (e) => e.namespaceURI !== P && e.getAttribute('id') === afterId,
  );
  if (section) {
    const se = section.cloneNode(false) as Element;
    se.setAttribute('id', id);
    section.parentElement!.insertBefore(se, section.nextSibling);
  }
  return id;
}
async function insertPage(
  pkg: OfficePackage,
  main: string,
  path: string,
  afterId: string,
  action: Extract<PageEdit, { kind: 'insert' | 'split' }>,
  slide: SlideState,
  size: PageSize,
) {
  const dest = pkg.unique('ppt/slides', 'slide', '.xml'),
    doc = pkg.doc(path).cloneNode(true) as Document,
    tree = shapeTree(doc);
  pkg.saveDoc(dest, doc);
  pkg.type(dest, pkg.contentType(path)!);
  const target =
    action.kind === 'split' ? slide.shapes.find((s) => s.id === action.shapeId) : undefined;
  const title = slide.shapes.find((s) => s.role === 'title' && s.id !== target?.id),
    body = target || slide.shapes.find((s) => s.role === 'body' && s.text);
  let next =
    Math.max(1, ...nodes(doc, P, 'cNvPr').map((e) => Number(e.getAttribute('id')) || 0)) + 1;
  const keep = new Set([
    title?.id,
    body?.id,
    ...slide.shapes
      .filter((s) => s.role === 'footer' || /logo|标志|徽标/i.test(s.name))
      .map((s) => s.id),
  ]);
  shapeNodes(doc)
    .filter((e) => !keep.has(idOf(e)))
    .forEach((e) => e.remove());
  for (const e of Array.from(doc.documentElement.children))
    if (!['cSld', 'clrMapOvr'].includes(e.localName)) e.remove();
  const cs = nodes(doc, P, 'cSld')[0];
  cs.setAttribute('name', action.title);
  child(cs, P, 'custDataLst')?.remove();
  child(cs, P, 'extLst')?.remove();
  child(cs, P, 'controls')?.remove();
  // Retain layout, background media and any relationships used by the retained text or logo.
  const used = new Set(
    Array.from(doc.getElementsByTagName('*')).flatMap((e) =>
      Array.from(e.attributes)
        .filter((a) => a.namespaceURI === R)
        .map((a) => a.value),
    ),
  );
  if (pkg.has(relPath(path))) {
    const rels = pkg.doc(relPath(path)).cloneNode(true) as Document;
    for (const r of Array.from(rels.documentElement.children))
      if (r.getAttribute('Type') !== R + '/slideLayout' && !used.has(r.getAttribute('Id')!))
        r.remove();
    pkg.saveDoc(relPath(dest), rels);
  }
  const continuation =
    action.kind === 'split' ? target!.text!.slice(action.at).trimStart() : action.body;
  for (const [role, original, value] of [
    ['title', title, action.title],
    ['body', body, continuation],
  ] as const) {
    const e = original && shapeNodes(doc).find((e) => idOf(e) === original.id);
    if (e && original) {
      if (!value) {
        e.remove();
        continue;
      }
      writeShape(
        e,
        {
          slideId: '',
          shapeId: original.id,
          text: value,
          ...(role === 'body' && action.kind === 'insert' && action.chart
            ? {
                left: size.width * 0.06,
                top: size.height * 0.24,
                width: size.width * 0.32,
                height: size.height * 0.63,
              }
            : {}),
        },
        original,
      );
    } else if (value) {
      const shape = textShape(doc, String(next++), role, size, value);
      tree.append(shape);
      if (role === 'body' && action.kind === 'insert' && action.chart)
        writeShape(
          shape,
          { slideId: '', shapeId: idOf(shape), width: size.width * 0.32 },
          {} as ShapeState,
        );
    }
  }
  if (action.kind === 'insert' && action.chart)
    await addChart(pkg, dest, tree, String(next++), action.chart, size, !!action.body);
  if (action.kind === 'split') {
    const originalShape = shapeNodes(pkg.edit(path)).find((e) => idOf(e) === target!.id)!;
    writeShape(
      originalShape,
      { slideId: slide.id, shapeId: target!.id, text: target!.text!.slice(0, action.at).trimEnd() },
      target!,
    );
  }
  return attachSlide(pkg, main, dest, afterId);
}
export class PptxWorkingCopy implements PresentationAccess {
  readonly capabilities = {
    read: true,
    selected: false,
    reorder: true,
    insert: false,
    create: false,
    thumbnail: false,
    backup: false,
    backgrounds: true,
    structure: true,
    charts: true,
  };
  private pkg: OfficePackage;
  private version = 0;
  private revisions = new WeakMap<PresentationSnapshot, number>();
  constructor(private bytes: Uint8Array) {
    this.pkg = new OfficePackage(unpack(bytes));
    mainPart(this.pkg);
  }
  read = async (): Promise<PresentationSnapshot> => {
    const snapshot = readPackage(this.pkg);
    this.revisions.set(snapshot, this.version);
    return snapshot;
  };
  backup = async () => new Blob([new Uint8Array(this.bytes)], { type: mime });
  image = (id: string) =>
    this.pkg.files[id] && /^image\/(png|jpeg|gif)$/i.test(this.pkg.contentType(id) || '')
      ? 'data:' + this.pkg.contentType(id) + ';base64,' + base64(this.pkg.files[id])
      : undefined;
  prepareInsert = async () => unsupported();
  insert = async () => unsupported();
  create = async () => unsupported();
  thumbnail = async () => unsupported();
  apply = async (
    before: PresentationSnapshot,
    input: EditProposal,
    images: Record<string, string> = {},
  ) => {
    const version = this.revisions.get(before);
    const check = () => {
      if (version !== this.version)
        throw Error(
          t('工作副本已更新，请重新读取页面。', 'The working copy changed. Read the slides again.'),
        );
    };
    check();
    assertSnapshotUnchanged(before, readPackage(this.pkg));
    const plan = validateEdit(input, before),
      pkg = new OfficePackage(unpack(this.bytes)),
      main = mainPart(pkg);
    const entries = slideEntries(pkg, main),
      paths = new Map(entries.map((e) => [e.id, e.path])),
      cache = new Map<string, string>();
    for (const patch of plan.changes) {
      const state = before.slides
        .find((s) => s.id === patch.slideId)!
        .shapes.find((s) => s.id === patch.shapeId)!;
      const shape = shapeNodes(pkg.edit(paths.get(patch.slideId)!)).find(
        (e) => idOf(e) === patch.shapeId,
      )!;
      writeShape(shape, patch, state);
    }
    if (plan.slideOrder) {
      pkg.edit(main);
      const ids = reorderedIds(before, plan.slideOrder);
      const list = entries[0].entry.parentElement!;
      for (const id of ids) list.append(entries.find((e) => e.id === id)!.entry);
    }
    const actions = plan.actions || [];
    for (const action of actions) {
      const path = paths.get(action.slideId)!;
      if (action.kind === 'background') {
        if (action.fill.kind === 'reference')
          copyBackground(pkg, paths.get(action.fill.sourceId)!, path, this.pkg);
        else writeBackground(pkg, path, action.fill, images, cache);
      } else if (action.kind === 'chart') {
        const shape = shapeNodes(pkg.doc(path)).find((e) => idOf(e) === action.shapeId)!;
        const graphic = nodes(shape, C, 'chart')[0];
        let chart = chartPart(pkg, path, shape)!;
        const references = entries.flatMap((e) =>
          nodes(pkg.doc(e.path), C, 'chart').filter(
            (c) => pkg.target(e.path, c.getAttributeNS(R, 'id')) === chart,
          ),
        );
        if (references.length > 1) {
          const part = pkg.unique(chart.slice(0, chart.lastIndexOf('/')), 'chart', '.xml');
          pkg.saveDoc(part, pkg.doc(chart).cloneNode(true) as Document);
          pkg.type(part, pkg.contentType(chart)!);
          if (pkg.has(relPath(chart)))
            pkg.saveDoc(relPath(part), pkg.doc(relPath(chart)).cloneNode(true) as Document);
          pkg.edit(path);
          graphic.setAttributeNS(R, 'r:id', pkg.relate(path, part, R + '/chart'));
          chart = part;
        }
        refreshChart(pkg, chart, action.data);
      }
    }
    const anchors = new Map<string, string>(),
      added = new Map<PageEdit, string>();
    for (const a of actions)
      if (a.kind === 'insert' || a.kind === 'split') {
        if (!before.size) throw Error(t('缺少页面尺寸。', 'Slide dimensions are missing.'));
        if (
          a.kind === 'split' &&
          (plan.changes.some(
            (p) => p.slideId === a.slideId && p.shapeId === a.shapeId && p.text !== undefined,
          ) ||
            actions.filter((b) => b.kind === 'split' && b.slideId === a.slideId).length > 1)
        )
          throw Error(
            t(
              '先确认正文修改，再按新文字拆页。',
              'Apply the text edit before splitting the revised body.',
            ),
          );
        const id = await insertPage(
          pkg,
          main,
          paths.get(a.slideId)!,
          anchors.get(a.slideId) || a.slideId,
          a,
          before.slides.find((s) => s.id === a.slideId)!,
          before.size,
        );
        anchors.set(a.slideId, id);
        added.set(a, id);
      }
    const bytes = pkg.pack(),
      reopened = new OfficePackage(unpack(bytes)),
      now = readPackage(reopened);
    const failed = () => {
      throw Error(
        t(
          '写入结果与预览不一致，工作副本未更新。',
          'The saved file differs from the preview. The working copy was not updated.',
        ),
      );
    };
    const expected = reorderedIds(before, plan.slideOrder).flatMap((id) => [
      id,
      ...Array.from(added)
        .filter(([a]) => a.slideId === id)
        .map(([, newId]) => newId),
    ]);
    if (JSON.stringify(expected) !== JSON.stringify(now.slideIds)) failed();
    for (const p of plan.changes) {
      const s = now.slides.find((s) => s.id === p.slideId)?.shapes.find((s) => s.id === p.shapeId);
      if (!s) failed();
      for (const [k, v] of Object.entries(p).filter(([k]) => !['shapeId', 'slideId'].includes(k))) {
        const actual = originalValue(s!, k);
        if (
          typeof v === 'number' && typeof actual === 'number'
            ? Math.abs(v - actual) > 0.02
            : actual !== v
        )
          failed();
      }
    }
    for (const a of actions) {
      const slide = now.slides.find((s) => s.id === a.slideId)!;
      if (
        a.kind === 'chart' &&
        JSON.stringify(slide.shapes.find((s) => s.id === a.shapeId)?.chart?.data) !==
          JSON.stringify(a.data)
      )
        failed();
      if (a.kind === 'background' && a.fill.kind === 'reference') {
        const source = a.fill.sourceId,
          expectedFill = before.slides.find((s) => s.id === source)?.effectiveBackground;
        if (
          expectedFill &&
          JSON.stringify(slide.effectiveBackground) !== JSON.stringify(expectedFill)
        )
          failed();
      }
      if (a.kind === 'background' && a.fill.kind !== 'reference') {
        const fill =
          a.fill.kind === 'image' ? { ...a.fill, imageId: cache.get(a.fill.imageId) } : a.fill;
        if (JSON.stringify(slide.background) !== JSON.stringify(fill)) failed();
      }
      if (a.kind === 'split' || a.kind === 'insert') {
        const newSlide = now.slides.find((s) => s.id === added.get(a));
        if (!newSlide?.shapes.some((s) => s.text === a.title)) failed();
        if (a.kind === 'split') {
          const old = before.slides
            .find((s) => s.id === a.slideId)!
            .shapes.find((s) => s.id === a.shapeId)!.text!;
          if (
            slide.shapes.find((s) => s.id === a.shapeId)?.text !== old.slice(0, a.at).trimEnd() ||
            !newSlide?.shapes.some((s) => s.text === old.slice(a.at).trimStart())
          )
            failed();
        }
        if (a.kind === 'insert' && a.body && !newSlide?.shapes.some((s) => s.text === a.body))
          failed();
        if (
          a.kind === 'insert' &&
          a.chart &&
          !newSlide?.shapes.some(
            (s) => JSON.stringify(s.chart?.data) === JSON.stringify(a.chart?.data),
          )
        )
          failed();
      }
    }
    check();
    this.bytes = bytes;
    this.pkg = reopened;
    this.version++;
  };
}
