import { t } from '../i18n';
import type { ShapePatch, ShapeState } from './types';
import { A, C, P, R, OfficePackage, child, element, need, nodes } from './package-xml';
import { hierarchy, themeColor, themeFont } from './package-theme';
import { readChart } from './package-charts';

export const emu = 12700;
export const shapeTree = (doc: Document) => nodes(doc, P, 'spTree')[0];
export const shapeNodes = (doc: Document) =>
  Array.from(shapeTree(doc).children).filter((e) =>
    ['sp', 'pic', 'graphicFrame', 'grpSp', 'cxnSp'].includes(e.localName),
  );
export const identity = (e: Element) => nodes(e, P, 'cNvPr')[0];
export const idOf = (e: Element) => identity(e)?.getAttribute('id') || '';
const body = (e: Element) => child(e, P, 'txBody');
const transform = (e: Element) =>
  child(e, P, 'xfrm') ||
  child(child(e, P, e.localName === 'grpSp' ? 'grpSpPr' : 'spPr'), A, 'xfrm');
export const chartPart = (pkg: OfficePackage, path: string, e: Element) =>
  pkg.target(path, nodes(e, C, 'chart')[0]?.getAttributeNS(R, 'id') ?? null);
function placeholderFamily(pkg: OfficePackage, path: string, e: Element) {
  const ph = nodes(e, P, 'ph')[0];
  if (!ph) return [e];
  const idx = ph.getAttribute('idx') || '0',
    type = ph.getAttribute('type') || 'obj';
  return [
    e,
    ...hierarchy(pkg, path)
      .slice(1, 3)
      .map((p) => {
        const all = shapeNodes(pkg.doc(p));
        return (
          all.find((s) => {
            const other = nodes(s, P, 'ph')[0];
            return other && (other.getAttribute('idx') || '0') === idx;
          }) || all.find((s) => nodes(s, P, 'ph')[0]?.getAttribute('type') === type)
        );
      })
      .filter((s): s is Element => Boolean(s)),
  ];
}
const alignRead: Record<string, string> = { l: 'Left', ctr: 'Center', r: 'Right' };
function textOf(e: Element) {
  const b = body(e);
  if (!b) return;
  return Array.from(b.children)
    .filter((c) => c.localName === 'p')
    .map((p) =>
      Array.from(p.children)
        .map((c) =>
          c.localName === 'br'
            ? '\n'
            : c.localName === 'r' || c.localName === 'fld'
              ? nodes(c, A, 't')
                  .map((v) => v.textContent)
                  .join('')
              : '',
        )
        .join(''),
    )
    .join('\n');
}
function geometry(xfrm?: Element) {
  const off = child(xfrm, A, 'off'),
    ext = child(xfrm, A, 'ext');
  return {
    left: Number(off?.getAttribute('x')) / emu,
    top: Number(off?.getAttribute('y')) / emu,
    width: Number(ext?.getAttribute('cx')) / emu,
    height: Number(ext?.getAttribute('cy')) / emu,
  };
}
export function readShape(pkg: OfficePackage, path: string, e: Element): ShapeState {
  const family = placeholderFamily(pkg, path, e),
    name = identity(e)?.getAttribute('name') || '',
    text = textOf(e);
  const ph = nodes(e, P, 'ph')[0]?.getAttribute('type') || '',
    chart = chartPart(pkg, path, e);
  const kind = chart
    ? 'Chart'
    : { sp: 'Shape', pic: 'Image', graphicFrame: 'Graphic', cxnSp: 'Line', grpSp: 'Group' }[
        e.localName
      ] || 'Object';
  const shape: ShapeState = {
    id: idOf(e),
    name,
    type: kind,
    ...geometry(family.map(transform).find(Boolean)),
  };
  if (chart) shape.chart = readChart(pkg, chart);
  if (e.localName === 'pic') {
    const blip = child(child(e, P, 'blipFill'), A, 'blip');
    const part = pkg.target(path, blip?.getAttributeNS(R, 'embed') ?? null);
    if (part) shape.picturePart = part;
  }
  if (text !== undefined && (text || nodes(e, P, 'ph').length)) {
    shape.text = text;
    shape.role =
      /title/i.test(ph) || /title|标题/i.test(name)
        ? 'title'
        : /^(ftr|dt|sldNum)$/.test(ph) || /footer|页脚/i.test(name)
          ? 'footer'
          : 'body';
    const defaults = family.flatMap((s) => {
      const b = body(s);
      return b ? [...nodes(b, A, 'defRPr'), ...nodes(b, A, 'endParaRPr')] : [];
    });
    const parts = hierarchy(pkg, path),
      level = Number(nodes(body(e)!, A, 'pPr')[0]?.getAttribute('lvl') || 0) + 1;
    const styleName =
      shape.role === 'title' ? 'titleStyle' : nodes(e, P, 'ph').length ? 'bodyStyle' : 'otherStyle';
    if (parts[2]) {
      const parent = nodes(pkg.doc(parts[2]), P, styleName)[0],
        style = child(child(parent, A, 'lvl' + level + 'pPr'), A, 'defRPr');
      if (style) defaults.push(style);
    }
    const explicit = nodes(body(e)!, A, 'r').map((r) => child(r, A, 'rPr'));
    const values = (get: (r: Element) => string | null | undefined) => {
      const fallback = defaults.map(get).find((v) => v != null) ?? null;
      const options = (explicit.length ? explicit : [undefined]).map(
        (prop) => (prop && get(prop)) ?? fallback,
      );
      return options.every((v) => v === options[0]) ? options[0] : null;
    };
    const size = values((r) => r.getAttribute('sz')),
      bold = values((r) => r.getAttribute('b')),
      color = values((r) => {
        const c = child(r, A, 'solidFill')?.firstElementChild;
        return c && !c.children.length ? themeColor(pkg, parts, c) : null;
      });
    shape.font = {
      size: size ? Number(size) / 100 : null,
      bold: bold === null ? null : bold === '1' || bold === 'true',
      name: values((r) => {
        const n = child(r, A, 'latin')?.getAttribute('typeface');
        return n ? themeFont(pkg, parts, n, shape.role!) : null;
      }),
      color: color ? '#' + color.toUpperCase() : null,
    };
    const alg = family
      .flatMap((s) => (body(s) ? nodes(body(s)!, A, 'pPr') : []))
      .map((p) => p.getAttribute('algn'))
      .find(Boolean);
    shape.alignment = alg ? alignRead[alg] || null : null;
  }
  return shape;
}
function editRuns(paragraph: Element, value: string) {
  if (child(paragraph, A, 'br') || child(paragraph, A, 'fld')) return false;
  const runs = nodes(paragraph, A, 't'),
    before = runs.map((r) => r.textContent).join('');
  if (!runs.length) return false;
  let left = 0,
    right = 0;
  while (left < Math.min(before.length, value.length) && before[left] === value[left]) left++;
  while (
    right < Math.min(before.length, value.length) - left &&
    before[before.length - 1 - right] === value[value.length - 1 - right]
  )
    right++;
  let offset = 0,
    inserted = false;
  for (const node of runs) {
    const text = node.textContent || '',
      end = offset + text.length;
    const prefix = text.slice(0, Math.max(0, Math.min(text.length, left - offset)));
    const suffix = text.slice(Math.max(0, before.length - right - offset));
    const here = !inserted && left >= offset && left <= end;
    node.textContent = prefix + (here ? value.slice(left, value.length - right) : '') + suffix;
    if (here) inserted = true;
    offset = end;
  }
  return true;
}
function writeText(e: Element, value: string) {
  const b = body(e)!;
  const d = e.ownerDocument;
  const old = Array.from(b.children).filter((c) => c.localName === 'p');
  const lines = value.split(/\r\n?|\n/);
  lines.forEach((v, i) => {
    const match =
      old[i] &&
      nodes(old[i], A, 't')
        .map((t) => t.textContent)
        .join('') === v
        ? old[i]
        : old.find(
            (p) =>
              nodes(p, A, 't')
                .map((t) => t.textContent)
                .join('') === v,
          );
    const p =
      ((match || old[i] || old[old.length - 1])?.cloneNode(true) as Element) ||
      element(d, A, 'a:p');
    if (match || editRuns(p, v)) {
      b.append(p);
      return;
    }
    const template = nodes(p, A, 'rPr')[0]?.cloneNode(true);
    for (const c of Array.from(p.children))
      if (!['pPr', 'endParaRPr'].includes(c.localName)) c.remove();
    const r = element(d, A, 'a:r'),
      tx = element(d, A, 'a:t');
    tx.textContent = v;
    if (template) r.append(template);
    r.append(tx);
    p.insertBefore(r, child(p, A, 'endParaRPr') || null);
    b.append(p);
  });
  old.forEach((p) => p.remove());
}
function shapeTransform(e: Element, state: ShapeState) {
  const old = transform(e);
  if (old) return old;
  const parent =
    e.localName === 'graphicFrame' ? e : child(e, P, e.localName === 'grpSp' ? 'grpSpPr' : 'spPr')!;
  const x = element(
    e.ownerDocument,
    e.localName === 'graphicFrame' ? P : A,
    e.localName === 'graphicFrame' ? 'p:xfrm' : 'a:xfrm',
  );
  parent.insertBefore(
    x,
    e.localName === 'graphicFrame' ? child(e, A, 'graphic') || null : parent.firstChild,
  );
  x.append(
    element(e.ownerDocument, A, 'a:off', {
      x: Math.round(state.left * emu),
      y: Math.round(state.top * emu),
    }),
    element(e.ownerDocument, A, 'a:ext', {
      cx: Math.round(state.width * emu),
      cy: Math.round(state.height * emu),
    }),
  );
  return x;
}
export function writeShape(e: Element, p: ShapePatch, state: ShapeState) {
  if (p.text !== undefined) writeText(e, p.text);
  if (
    p.left !== undefined ||
    p.top !== undefined ||
    p.height !== undefined ||
    p.width !== undefined
  ) {
    const xfrm = shapeTransform(e, state);
    for (const [key, target, attr] of [
      ['left', 'off', 'x'],
      ['top', 'off', 'y'],
      ['width', 'ext', 'cx'],
      ['height', 'ext', 'cy'],
    ] as const)
      if (p[key] !== undefined)
        need(xfrm, A, 'a:' + target).setAttribute(attr, String(Math.round(p[key]! * emu)));
  }
  if ([p.bold, p.fontName, p.fontSize, p.fontColor, p.alignment].every((x) => x === undefined))
    return;
  const b = body(e);
  if (!b) throw Error(t('对象没有可编辑文字。', 'The object has no editable text.'));
  const d = e.ownerDocument;
  for (const para of Array.from(b.children).filter((c) => c.localName === 'p')) {
    const props = need(para, A, 'a:pPr', true);
    if (p.alignment)
      props.setAttribute('algn', { Left: 'l', Center: 'ctr', Right: 'r' }[p.alignment]);
    if ([p.bold, p.fontName, p.fontSize, p.fontColor].every((x) => x === undefined)) continue;
    const def = need(props, A, 'a:defRPr'),
      runs = Array.from(para.children)
        .filter((c) => ['r', 'fld', 'br'].includes(c.localName))
        .map((r) => need(r, A, 'a:rPr', true));
    for (const r of [def, ...runs, ...nodes(para, A, 'endParaRPr')]) {
      if (p.bold !== undefined) r.setAttribute('b', p.bold ? '1' : '0');
      if (p.fontSize !== undefined) r.setAttribute('sz', String(Math.round(p.fontSize * 100)));
      if (p.fontColor) {
        for (const c of Array.from(r.children))
          if (
            ['noFill', 'solidFill', 'gradFill', 'blipFill', 'pattFill', 'grpFill'].includes(
              c.localName,
            )
          )
            c.remove();
        const f = element(d, A, 'a:solidFill');
        f.append(element(d, A, 'a:srgbClr', { val: p.fontColor.slice(1) }));
        r.insertBefore(f, Array.from(r.children).find((c) => c.localName !== 'ln') || null);
      }
      if (p.fontName) {
        for (const key of ['latin', 'ea', 'cs']) {
          let f = child(r, A, key);
          if (!f) {
            f = element(d, A, 'a:' + key);
            const order = [
              'latin',
              'ea',
              'cs',
              'sym',
              'hlinkClick',
              'hlinkMouseOver',
              'rtl',
              'extLst',
            ];
            r.insertBefore(
              f,
              Array.from(r.children).find((c) => order.indexOf(c.localName) > order.indexOf(key)) ||
                null,
            );
          }
          f.setAttribute('typeface', p.fontName);
        }
      }
    }
  }
}
