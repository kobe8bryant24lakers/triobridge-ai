import { t } from '../i18n';
import { validateEdit, replaceProposal } from './editing';
import { contrast, mergeProposals, pagesForRange, parseFill, hexColor } from './page-edits';
import type {
  BackgroundFill,
  EditProposal,
  PresentationSnapshot,
  ShapePatch,
  ShapeState,
  SlideState,
} from './types';

export interface EditRecipe {
  name: string;
  fill?: BackgroundFill;
  pictureName?: string;
  titleColor?: string;
  find?: string;
  replacement?: string;
}
export function applyRecipe(
  recipe: EditRecipe,
  snapshot: PresentationSnapshot,
  targets: SlideState[],
) {
  let plan: EditProposal = {
    summary: recipe.name || t('批量编辑', 'Batch edit'),
    changes: [],
    actions: [],
  };
  if (recipe.fill) {
    const fill = parseFill(recipe.fill);
    plan.actions = targets.map((s) => ({ kind: 'background', slideId: s.id, fill }));
  }
  if (recipe.titleColor) {
    const fontColor = hexColor(recipe.titleColor);
    plan.changes = targets.flatMap((s) =>
      s.shapes
        .filter((sh) => sh.role === 'title' && sh.text)
        .map((sh) => ({ slideId: s.id, shapeId: sh.id, fontColor })),
    );
  }
  if (recipe.find)
    plan = mergeProposals(
      plan,
      replaceProposal({ ...snapshot, slides: targets }, recipe.find, recipe.replacement || ''),
    );
  plan.summary =
    recipe.name || t('批量修改页面背景、标题或用词。', 'Update backgrounds, titles or wording.');
  return validateEdit(plan, snapshot);
}
export interface SampleStyle {
  background: boolean;
  title: boolean;
  body: boolean;
  geometry: boolean;
  alignment: boolean;
}
export function sampleStyle(
  snapshot: PresentationSnapshot,
  ref: SlideState,
  targets: SlideState[],
  opts: SampleStyle,
): EditProposal {
  const changes: ShapePatch[] = [];
  for (const slide of targets.filter((s) => s.id !== ref.id)) {
    for (const role of ['title', 'body'] as const) {
      const source = ref.shapes.filter((s) => s.role === role && s.text?.trim()),
        dest = slide.shapes.filter((s) => s.role === role && s.text?.trim());
      dest.forEach((s, i) => {
        const match =
          source.find(
            (o) => o.name === s.name && source.filter((o2) => o2.name === s.name).length === 1,
          ) ||
          source[i] ||
          source[source.length - 1];
        if (!match) return;
        const p: ShapePatch = { slideId: slide.id, shapeId: s.id };
        if (opts[role]) {
          if (match.font?.name) p.fontName = match.font.name;
          if (match.font?.size && match.font.size >= 8 && match.font.size <= 96)
            p.fontSize = match.font.size;
          if (match.font?.color?.match(/^#[a-f\d]{6}$/i)) p.fontColor = match.font.color;
          if (typeof match.font?.bold === 'boolean') p.bold = match.font.bold;
        }
        if (opts.alignment && ['Left', 'Center', 'Right'].includes(match.alignment || ''))
          p.alignment = match.alignment as ShapePatch['alignment'];
        if (
          opts.geometry &&
          snapshot.size &&
          i < source.length &&
          match.width >= 1 &&
          match.height >= 1
        ) {
          p.left = match.left;
          p.top = match.top;
          p.width = match.width;
          p.height = match.height;
        }
        if (Object.keys(p).length > 2) changes.push(p);
      });
    }
  }
  return validateEdit(
    {
      summary: t(
        '套用第 ' + (ref.index + 1) + ' 页的所选样式。',
        'Apply the selected styles from slide ' + (ref.index + 1) + '.',
      ),
      changes,
      actions: opts.background
        ? targets
            .filter((s) => s.id !== ref.id)
            .map((s) => ({
              kind: 'background',
              slideId: s.id,
              fill: { kind: 'reference', sourceId: ref.id },
            }))
        : [],
    },
    snapshot,
  );
}
export function splitPoint(text: string) {
  const center = text.length / 2;
  const breaks = Array.from(text.matchAll(/[\n。！？.!?；;]\s*|\s+/g))
    .map((m) => m.index! + m[0].length)
    .filter((i) => i > 0 && i < text.length);
  return (
    breaks.sort((a, b) => Math.abs(a - center) - Math.abs(b - center))[0] ??
    Math.max(1, Math.floor(center))
  );
}
export interface SlideIssue {
  slide: SlideState;
  shape?: ShapeState;
  message: string;
  type: 'comment' | 'density' | 'bounds' | 'contrast' | 'small';
  fix?: EditProposal;
  commentId?: string;
}
export function slideIssues(snapshot: PresentationSnapshot): SlideIssue[] {
  const issues: SlideIssue[] = [];
  for (const slide of snapshot.slides) {
    for (const c of slide.comments || [])
      issues.push({
        slide,
        shape: slide.shapes.find((s) => s.id === c.shapeId),
        message: c.text,
        type: 'comment',
        commentId: c.id,
      });
    for (const shape of slide.shapes) {
      const { size } = snapshot,
        id = { slideId: slide.id, shapeId: shape.id };
      const add = (
        type: SlideIssue['type'],
        message: string,
        p?: Omit<ShapePatch, 'slideId' | 'shapeId'>,
      ) =>
        issues.push({
          slide,
          shape,
          type,
          message,
          ...(p ? { fix: { summary: message, changes: [{ ...id, ...p }] } } : {}),
        });
      if (
        size &&
        (shape.left < -0.1 ||
          shape.top < -0.1 ||
          shape.left + shape.width > size.width + 0.1 ||
          shape.top + shape.height > size.height + 0.1)
      ) {
        const width = Math.min(size.width, shape.width),
          height = Math.min(size.height, shape.height);
        add('bounds', t('对象超出页面，移回可见范围。', 'Move the object back inside the slide.'), {
          left: Math.max(0, Math.min(size.width - width, shape.left)),
          top: Math.max(0, Math.min(size.height - height, shape.top)),
          ...(width !== shape.width ? { width } : {}),
          ...(height !== shape.height ? { height } : {}),
        });
      }
      if (shape.role === 'footer' || !shape.text?.trim()) continue;
      if (shape.font?.size && shape.font.size < 14)
        add('small', t('字号小于 14 磅，调至 14 磅。', 'Increase text below 14 pt to 14 pt.'), {
          fontSize: 14,
        });
      const points = shape.font?.size || 18;
      const estLines = shape.text
        .split(/\n/)
        .reduce(
          (n, l) =>
            n +
            Math.max(
              1,
              Math.ceil(
                (Array.from(l).reduce((x, c) => x + (/[^\x00-\x7F]/.test(c) ? 1 : 0.55), 0) *
                  points) /
                  Math.max(1, shape.width),
              ),
            ),
          0,
        );
      if (
        shape.role === 'body' &&
        shape.text.trim().length > 1 &&
        (shape.text.length > 500 ||
          (shape.height > 0 && shape.width > 0 && estLines * points * 1.15 > shape.height + points))
      ) {
        const message = t('正文较密，拆为两页。', 'Split dense body text over two slides.');
        issues.push({
          slide,
          shape,
          type: 'density',
          message,
          fix: {
            summary: message,
            changes: [],
            actions: [
              {
                kind: 'split',
                ...id,
                at: splitPoint(shape.text),
                title: slide.title.slice(0, 145) + t('（续）', ' (continued)'),
              },
            ],
          },
        });
      }
      const bg = slide.effectiveBackground || slide.background;
      if (shape.font?.color && bg && (bg.kind === 'solid' || bg.kind === 'gradient')) {
        const colors = bg.kind === 'solid' ? [bg.color] : [bg.from, bg.to];
        const score = (c: string) => Math.min(...colors.map((b) => contrast(c, b)));
        const limit = points >= 18 || (shape.font.bold && points >= 14) ? 3 : 4.5;
        if (score(shape.font.color) < limit) {
          const fontColor = score('#000000') >= score('#FFFFFF') ? '#000000' : '#FFFFFF';
          if (score(fontColor) >= limit)
            add(
              'contrast',
              t(
                '文字与背景对比度低，调整文字颜色。',
                'Increase text contrast against the slide background.',
              ),
              { fontColor },
            );
          else
            add(
              'contrast',
              t(
                '渐变两端亮度差异大，请检查文字可读性。',
                'Check text readability across this gradient.',
              ),
            );
        }
      }
    }
  }
  return issues;
}
const colors: Record<string, string> = {
  米白色: '#F5F1E8',
  米白: '#F5F1E8',
  奶油色: '#FFFDD0',
  深蓝色: '#17365D',
  深蓝: '#17365D',
  藏青色: '#17365D',
  白色: '#FFFFFF',
  黑色: '#000000',
  红色: '#CC3333',
  蓝色: '#2865AA',
  绿色: '#28604A',
  cream: '#FFFDD0',
  navy: '#17365D',
  white: '#FFFFFF',
  black: '#000000',
  blue: '#2865AA',
  green: '#28604A',
};
export function parseBackgroundClauses(brief: string) {
  if (!/背景|background/i.test(brief)) return;
  const clauses: { range: string; fill: BackgroundFill }[] = [];
  const chunks = brief
    .trim()
    .split(/[。；;\n]+|[,，]\s*(?=(?:但|其中|第|but|except|slide))/i)
    .filter(Boolean);
  for (const [i, chunk] of chunks.entries()) {
    if (i === 0 && !/背景|background/i.test(chunk)) return;
    const hits = Array.from(
      chunk.matchAll(new RegExp('#[0-9a-f]{6}\\b|' + Object.keys(colors).join('|'), 'gi')),
    );
    const master = /母版|master/i.test(chunk),
      gradient = /渐变|gradient/i.test(chunk);
    if (!master && hits.length !== (gradient ? 2 : 1)) return;
    const c = hits.map((h) => colors[h[0].toLowerCase()] || h[0].toUpperCase());
    const fill: BackgroundFill = master
      ? { kind: 'master' }
      : gradient
        ? { kind: 'gradient', from: c[0], to: c[1] }
        : { kind: 'solid', color: c[0] };
    const stripped = chunk
      .replace(/#[a-f\d]{6}/gi, '')
      .replace(new RegExp(Object.keys(colors).join('|'), 'gi'), '');
    const range = /第\s*([\d\s,，、\-–至]+)\s*页|slides?\s+([\d,\-–\s]+)/i.exec(stripped);
    const rem = stripped
      .replace(range?.[0] || '\u0001', '')
      .replace(
        /整个文档|整份|全稿|整稿|所有|全部|整套|整个|全文|演示稿|幻灯片|当前|选中|本页|这页|页面|背景色?|颜色|纯色|双色|渐变|恢复|跟随|母版|指定|其中|设置?|切换|改成|改为|换成|调成|应用|使用|统一|设为|改|为|成|把|将|的|到|从|和|但|整|稿|页|pptx?|powerpoint|presentations?|backgrounds?|gradient|master|restore|follow|apply|set|change|the|whole|entire|deck|all|slides?|selected|current|on|in|to|from|and|but|colour|color|[\s,，:：.]/gi,
        '',
      );
    if (rem) return;
    clauses.push({ range: range?.[1] || range?.[2] || '', fill });
  }
  return clauses.length ? clauses : undefined;
}

export function localBackgroundRequest(
  brief: string,
  snapshot: PresentationSnapshot,
  clauses = parseBackgroundClauses(brief),
): EditProposal | undefined {
  if (!clauses) return;
  const actions = new Map<string, NonNullable<EditProposal['actions']>[number]>();
  for (const { range, fill } of clauses) {
    const targets = pagesForRange(snapshot.slides, range);
    for (const s of targets) actions.set(s.id, { kind: 'background', slideId: s.id, fill });
  }
  return actions.size
    ? validateEdit(
        { summary: brief.slice(0, 600), changes: [], actions: [...actions.values()] },
        snapshot,
      )
    : undefined;
}

export function proposedSnapshot(
  snapshot: PresentationSnapshot,
  proposal: EditProposal,
): PresentationSnapshot {
  const view = structuredClone(snapshot);
  for (const p of proposal.changes) {
    const shape = view.slides
      .find((s) => s.id === p.slideId)
      ?.shapes.find((s) => s.id === p.shapeId);
    if (!shape) continue;
    for (const k of ['text', 'left', 'top', 'width', 'height', 'alignment'] as const)
      if (p[k] !== undefined) Object.assign(shape, { [k]: p[k] });
    if (shape.font) {
      if (p.fontName !== undefined) shape.font.name = p.fontName;
      if (p.fontSize !== undefined) shape.font.size = p.fontSize;
      if (p.fontColor !== undefined) shape.font.color = p.fontColor;
      if (p.bold !== undefined) shape.font.bold = p.bold;
    }
  }
  for (const a of proposal.actions || []) {
    const sl = view.slides.find((s) => s.id === a.slideId);
    if (!sl) continue;
    if (a.kind === 'background') {
      const fill = a.fill;
      sl.background =
        fill.kind === 'reference'
          ? snapshot.slides.find((s) => s.id === fill.sourceId)?.background
          : fill;
      sl.effectiveBackground =
        fill.kind === 'reference'
          ? snapshot.slides.find((s) => s.id === fill.sourceId)?.effectiveBackground
          : fill.kind === 'master'
            ? sl.masterBackground
            : fill;
    }
    if (a.kind === 'chart') {
      const sh = sl.shapes.find((s) => s.id === a.shapeId);
      if (sh?.chart) sh.chart.data = a.data;
    }
  }
  return view;
}
