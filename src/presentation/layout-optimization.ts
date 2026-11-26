import { t } from '../i18n';
import { validateEdit } from './editing';
import type {
  DeckSlide,
  EditProposal,
  PresentationSnapshot,
  ShapePatch,
  ShapeState,
} from './types';

export type Arrangement = 'imageRight' | 'bulletsInColumns' | 'stackSteps';
export function suggestedArrangement(slide: DeckSlide): Arrangement | undefined {
  if (slide.layout === 'image' && slide.bullets?.length) return 'imageRight';
  if (slide.layout === 'bullets' && (slide.bullets?.length || 0) >= 3) return 'bulletsInColumns';
  if (slide.layout === 'process' && (slide.steps?.length || 0) >= 3) return 'stackSteps';
}
export function arrangementName(value: Arrangement) {
  return {
    imageRight: t('左侧讲解，右侧配图', 'Text left, image right'),
    bulletsInColumns: t('要点分为两栏', 'Two-column list'),
    stackSteps: t('步骤纵向排列', 'Vertical steps'),
  }[value];
}
type Box = Pick<ShapeState, 'left' | 'top' | 'width' | 'height'>;
const overlap = (a: Box, b: Box) =>
  Math.max(0, Math.min(a.left + a.width, b.left + b.width) - Math.max(a.left, b.left)) *
  Math.max(0, Math.min(a.top + a.height, b.top + b.height) - Math.max(a.top, b.top));
export function optimizeLayout(snapshot: PresentationSnapshot): EditProposal {
  const size = snapshot.size;
  if (!size) throw Error(t('请先读取页面尺寸。', 'Read the slide size first.'));
  const changes: ShapePatch[] = [];
  for (const slide of snapshot.slides) {
    const boxes = slide.shapes.filter(
      (s) =>
        [s.left, s.top, s.width, s.height].every(Number.isFinite) && s.width > 0 && s.height > 0,
    );
    const movable = boxes.filter(
      (s) =>
        s.type === 'Shape' &&
        s.text?.trim() &&
        s.role !== 'footer' &&
        !/logo|标志|徽标/i.test(s.name),
    );
    const placed = new Map(boxes.map((s) => [s.id, { ...s }]));
    for (const s of movable) {
      const b = {
        ...s,
        width: Math.min(s.width, size.width),
        height: Math.min(s.height, size.height),
      };
      b.left = Math.max(0, Math.min(s.left, size.width - b.width));
      b.top = Math.max(0, Math.min(s.top, size.height - b.height));
      const sameColumn = movable.filter(
        (p) =>
          p !== s &&
          p.role === s.role &&
          Math.abs(p.width - s.width) < 8 &&
          Math.abs(p.left - s.left) <= 6 &&
          Math.abs(p.top - s.top) > s.height * 0.6,
      );
      const anchor = Math.min(b.left, ...sameColumn.map((p) => p.left).filter((n) => n >= 0));
      if (anchor >= 0 && anchor + b.width <= size.width) b.left = anchor;
      const inBounds = (r: Box) =>
        r.left >= 0 &&
        r.top >= 0 &&
        r.left + r.width <= size.width &&
        r.top + r.height <= size.height;
      const obstacles = boxes
        .filter(
          (o) =>
            o.id !== s.id &&
            (o.text?.trim() || ['Chart', 'Image', 'Graphic', 'Group'].includes(o.type)),
        )
        .map((o) => placed.get(o.id)!);
      const wouldCollide = (candidate: Box) =>
        obstacles.some((o) => overlap(candidate, o) > overlap(s, o) + 1);
      if (wouldCollide(b)) continue;
      // Only separate shallow collisions between peers. Text over a graphic can be intentional.
      const prev = obstacles
        .filter(
          (o) =>
            o.text &&
            o.top < s.top &&
            o.left === b.left &&
            Math.abs(o.width - b.width) < 8 &&
            o.top + o.height > b.top &&
            o.top + o.height - b.top <= 8,
        )
        .sort((a, c) => c.top - a.top)[0];
      if (prev) {
        const separated = { ...b, top: prev.top + prev.height + 3 };
        if (inBounds(separated) && !wouldCollide(separated)) b.top = separated.top;
      }
      const patch: ShapePatch = { slideId: slide.id, shapeId: s.id };
      for (const k of ['width', 'height', 'left', 'top'] as const)
        if (Math.abs(b[k] - s[k]) > 0.25) patch[k] = b[k];
      if (Object.keys(patch).length > 2) {
        placed.set(s.id, b);
        changes.push(patch);
      }
    }
  }
  return validateEdit(
    {
      summary: t(
        '将越界文本框收回页面，对齐相近的同列文本框，分开轻微相叠的同列段落。',
        'Bring text inside the page, align nearby column edges and separate shallow overlaps between peer text boxes.',
      ),
      changes,
    },
    snapshot,
  );
}

/** Swap an existing side-by-side picture and its text column as groups. */
export function swapPictureColumns(snapshot: PresentationSnapshot): EditProposal {
  if (!snapshot.size) throw Error(t('请先读取页面尺寸。', 'Read the slide size first.'));
  const width = snapshot.size.width,
    changes: ShapePatch[] = [];
  for (const slide of snapshot.slides) {
    const images = slide.shapes.filter(
      (s) =>
        s.type === 'Image' && s.width > width * 0.18 && s.height > snapshot.size!.height * 0.17,
    );
    if (images.length !== 1) continue;
    const pic = images[0],
      words = slide.shapes.filter(
        (s) =>
          s.role === 'body' &&
          s.text?.trim() &&
          s.width > width * 0.1 &&
          s.top >= snapshot.size!.height * 0.19 &&
          s.top < snapshot.size!.height * 0.88,
      );
    if (!words.length) continue;
    const min = Math.min(...words.map((s) => s.left)),
      max = Math.max(...words.map((s) => s.left + s.width));
    if (!(max <= pic.left - 2 || min >= pic.left + pic.width + 2)) continue;
    const top = Math.min(...words.map((s) => s.top)),
      bottom = Math.max(...words.map((s) => s.top + s.height));
    const marks = slide.shapes.filter(
      (s) =>
        s.type === 'Shape' &&
        !s.text &&
        s.width > 0 &&
        s.width < width * 0.025 &&
        s.height < snapshot.size!.height * 0.05 &&
        s.left >= min - width * 0.035 &&
        s.left <= min &&
        s.top >= top - 2 &&
        s.top <= bottom,
    );
    const group = [...words, ...marks],
      groupLeft = Math.min(min, ...marks.map((s) => s.left));
    const shift = width - max - groupLeft,
      imgLeft = width - pic.left - pic.width;
    const moves = [{ s: pic, left: imgLeft }, ...group.map((s) => ({ s, left: s.left + shift }))];
    const keep = slide.shapes.filter(
      (o) =>
        !moves.some((m) => m.s === o) &&
        (o.text?.trim() || ['Chart', 'Image', 'Graphic', 'Group'].includes(o.type)),
    );
    if (
      moves.some(
        ({ s, left }) =>
          !Number.isFinite(left) ||
          left < 0 ||
          left + s.width > width + 0.1 ||
          keep.some((other) => overlap({ ...s, left }, other) > overlap(s, other) + 1),
      )
    )
      continue;
    for (const { s, left } of moves)
      if (Math.abs(left - s.left) > 0.25) changes.push({ slideId: slide.id, shapeId: s.id, left });
  }
  return validateEdit(
    {
      summary: t(
        '调换独立图片与文字栏的位置，保留栏内顺序、图片比例和原始尺寸。',
        'Swap the image and text columns while keeping their internal order and original dimensions.',
      ),
      changes,
    },
    snapshot,
  );
}
