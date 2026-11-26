import { t } from '../i18n';
import type {
  BackgroundFill,
  ChartData,
  ChartKind,
  EditProposal,
  PageEdit,
  PresentationSnapshot,
  ShapePatch,
  SlideBackground,
  SlideState,
} from './types';
import { fields, list, object, text } from './validation';

export function hexColor(input: unknown): string {
  if (typeof input !== 'string' || !/^#[A-F\d]{6}$/i.test(input))
    throw Error(
      t('请输入六位 HEX 色值，例如 #18324A。', 'Enter a six-digit HEX color such as #18324A.'),
    );
  return input.toUpperCase();
}
export function parseFill(input: unknown): BackgroundFill {
  const v = object(input, t('页面背景', 'Slide background'));
  switch (v.kind) {
    case 'master':
      fields(v, ['kind'], 'background');
      return { kind: 'master' };
    case 'solid':
      fields(v, ['kind', 'color'], 'background');
      return { kind: 'solid', color: hexColor(v.color) };
    case 'gradient':
      fields(v, ['kind', 'from', 'to'], 'background');
      return { kind: 'gradient', from: hexColor(v.from), to: hexColor(v.to) };
    case 'image': {
      fields(v, ['kind', 'imageId', 'transparency'], 'background');
      const transparency = v.transparency;
      if (
        typeof transparency !== 'number' ||
        !Number.isInteger(transparency) ||
        transparency < 0 ||
        transparency > 100
      )
        throw Error(t('图片透明度需要在 0–100% 之间。', 'Picture transparency must be 0–100%.'));
      return { kind: 'image', imageId: text(v.imageId, 160, 'imageId'), transparency };
    }
    default:
      throw Error(t('未知的背景类型。', 'Unknown background type.'));
  }
}
export function backgroundLabel(fill?: SlideBackground | { kind: 'reference'; sourceId: string }) {
  if (!fill) return t('未读取', 'Unavailable');
  switch (fill.kind) {
    case 'master':
      return t('跟随母版', 'Follow master');
    case 'solid':
      return fill.color;
    case 'gradient':
      return fill.from + ' → ' + fill.to;
    case 'image':
      return t('图片 · 透明度 ', 'Picture · transparency ') + fill.transparency + '%';
    case 'reference':
      return t('复制参考页背景', 'Copy reference background');
    case 'other':
      return fill.label;
  }
}
export function pagesForRange(slides: SlideState[], input: string, exclude = '') {
  function parse(s: string) {
    const picked = new Set<number>();
    for (const part of s
      .trim()
      .split(/[,，、\s]+/)
      .filter(Boolean)) {
      const m = /^(\d+)(?:[-–—至](\d+))?$/.exec(part);
      if (!m) throw Error(t('页码格式示例：1,3-8,12。', 'Enter slide numbers such as 1,3-8,12.'));
      const a = Number(m[1]),
        b = Number(m[2] || m[1]);
      if (a < 1 || b < a || b > 500)
        throw Error(t('页码超出范围。', 'Slide numbers are out of range.'));
      for (let n = a; n <= b; n++) {
        if (!slides.some((sl) => sl.index + 1 === n))
          throw Error(
            t('请先读取包含这些页码的范围。', 'Read a range containing these slide numbers first.'),
          );
        picked.add(n);
      }
    }
    return picked;
  }
  const include = parse(input),
    skip = parse(exclude);
  const result = slides.filter(
    (s) => (!include.size || include.has(s.index + 1)) && !skip.has(s.index + 1),
  );
  if (!result.length)
    throw Error(t('请至少保留一个目标页面。', 'Choose at least one target slide.'));
  return result;
}
export function checkedData(value: unknown, kind: ChartKind): ChartData {
  const v = object(value, 'chart'),
    categories = list(v.categories, 30, 'categories').map((c) => text(c, 300, 'category'));
  const series = list(v.series, 6, 'series').map((item) => {
    const s = object(item, 'series');
    if (
      !Array.isArray(s.values) ||
      s.values.length !== categories.length ||
      s.values.some((x) => x !== null && (typeof x !== 'number' || !Number.isFinite(x)))
    )
      throw Error(t('图表数据与分类不匹配。', 'Chart values must match the categories.'));
    return { name: text(s.name, 300, 'series'), values: s.values as (number | null)[] };
  });
  if (
    !categories.length ||
    !series.length ||
    (['pie', 'doughnut'].includes(kind) &&
      (series.length !== 1 ||
        categories.length > 8 ||
        series[0].values.some((x) => x === null || x < 0) ||
        !series[0].values.some((x) => x !== null && x > 0)))
  )
    throw Error(t('此数据不适用于当前图表类型。', 'The data does not fit this chart type.'));
  return { categories, series };
}
export function actionKey(a: PageEdit) {
  return JSON.stringify([a.kind, a.slideId, 'shapeId' in a ? a.shapeId : '']);
}
export function parseActions(value: unknown, snapshot: PresentationSnapshot): PageEdit[] {
  const seen = new Set<string>();
  return list(value, 600, t('页面改动', 'Slide edits'))
    .map((input): PageEdit => {
      const v = object(input, 'edit'),
        slide = snapshot.slides.find((s) => s.id === v.slideId);
      if (!slide)
        throw Error(t('目标页面在读取范围以外。', 'Target slide is outside the reviewed range.'));
      const slideId = slide.id,
        shape = slide.shapes.find((s) => s.id === v.shapeId);
      let result: PageEdit;
      switch (v.kind) {
        case 'background': {
          fields(v, ['kind', 'slideId', 'fill'], 'background');
          const f = object(v.fill, 'fill');
          if (!slide.background)
            throw Error(
              t(
                '当前入口未读取背景，请打开 PPTX 文件编辑。',
                'Open the PPTX file to edit backgrounds in this host.',
              ),
            );
          if (f.kind === 'reference') {
            fields(f, ['kind', 'sourceId'], 'reference');
            const ref = snapshot.slides.find((s) => s.id === f.sourceId);
            if (!ref?.background)
              throw Error(t('请重新选择参考页。', 'Select a reference slide again.'));
            result = { kind: 'background', slideId, fill: { kind: 'reference', sourceId: ref.id } };
          } else result = { kind: 'background', slideId, fill: parseFill(f) };
          break;
        }
        case 'chart': {
          fields(v, ['kind', 'slideId', 'shapeId', 'data'], 'chart');
          if (!shape?.chart)
            throw Error(
              t('请选择已读取数据的原生图表。', 'Select a native chart with readable data.'),
            );
          result = {
            kind: 'chart',
            slideId,
            shapeId: shape.id,
            data: checkedData(v.data, shape.chart.kind),
          };
          break;
        }
        case 'insert': {
          fields(v, ['kind', 'slideId', 'title', 'body', 'chart'], 'insert');
          result = {
            kind: 'insert',
            slideId,
            title: text(v.title, 160, 'title'),
            body: text(v.body, 4000, 'body', true),
          };
          if (v.chart !== undefined) {
            const c = object(v.chart, 'chart');
            if (!['column', 'bar', 'line', 'pie', 'doughnut'].includes(String(c.kind)))
              throw Error(t('图表类型无效。', 'Invalid chart type.'));
            const kind = c.kind as ChartKind;
            result.chart = {
              kind,
              data: checkedData(c.data, kind),
              unit: text(c.unit ?? '', 32, 'unit', true),
            };
          }
          break;
        }
        case 'split': {
          fields(v, ['kind', 'slideId', 'shapeId', 'at', 'title'], 'split');
          if (
            !shape?.text ||
            shape.role !== 'body' ||
            typeof v.at !== 'number' ||
            !Number.isInteger(v.at) ||
            !shape.text.slice(0, v.at).trim() ||
            !shape.text.slice(v.at).trim() ||
            v.at <= 0 ||
            v.at >= shape.text.length
          )
            throw Error(
              t(
                '拆页位置必须将文字分成两段非空内容。',
                'Choose a split point with text on both sides.',
              ),
            );
          result = {
            kind: 'split',
            slideId,
            shapeId: shape.id,
            at: v.at,
            title: text(v.title, 160, 'title'),
          };
          break;
        }
        default:
          throw Error(t('未知的页面改动。', 'Unknown slide edit.'));
      }
      const key = actionKey(result);
      if (seen.has(key)) throw Error(t('页面改动重复。', 'Duplicate slide edit.'));
      seen.add(key);
      return result;
    })
    .filter((a) => {
      const slide = snapshot.slides.find((s) => s.id === a.slideId)!;
      if (a.kind === 'background')
        return (
          !(a.fill.kind === 'reference' && a.fill.sourceId === a.slideId) &&
          JSON.stringify(slide.background) !== JSON.stringify(a.fill)
        );
      if (a.kind === 'chart')
        return (
          JSON.stringify(slide.shapes.find((s) => s.id === a.shapeId)?.chart?.data) !==
          JSON.stringify(a.data)
        );
      return true;
    });
}
export function mergeProposals(base: EditProposal, next: EditProposal): EditProposal {
  const shapes = new Map<string, ShapePatch>();
  for (const p of [...base.changes, ...next.changes]) {
    const key = JSON.stringify([p.slideId, p.shapeId]);
    shapes.set(key, { ...shapes.get(key), ...p });
  }
  const pages = new Map<string, PageEdit>();
  for (const a of [...(base.actions || []), ...(next.actions || [])]) pages.set(actionKey(a), a);
  return {
    summary: next.summary,
    changes: [...shapes.values()],
    ...(pages.size ? { actions: [...pages.values()] } : {}),
    ...(next.slideOrder || base.slideOrder
      ? { slideOrder: next.slideOrder || base.slideOrder }
      : {}),
  };
}
export function contrast(a: string, b: string) {
  const luminance = (c: string) =>
    [1, 3, 5]
      .map((i) => parseInt(c.slice(i, i + 2), 16) / 255)
      .map((v) => (v <= 0.04045 ? v / 12.92 : ((v + 0.055) / 1.055) ** 2.4))
      .reduce((s, x, i) => s + x * [0.2126, 0.7152, 0.0722][i], 0);
  const x = luminance(a),
    y = luminance(b);
  return (Math.max(x, y) + 0.05) / (Math.min(x, y) + 0.05);
}
