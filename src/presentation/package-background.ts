import { hierarchy, themeColor } from './package-theme';
import { t } from '../i18n';
import type { BackgroundFill, SlideBackground } from './types';
import { A, P, R, OfficePackage, child, element, nodes, bytes64, xml } from './package-xml';
function resolvedBackground(pkg: OfficePackage, path: string, inherited = false) {
  const parts = hierarchy(pkg, path);
  for (const part of parts.slice(inherited ? 1 : 0, 3)) {
    const original = child(nodes(pkg.doc(part), P, 'cSld')[0], P, 'bg');
    if (!original) continue;
    const bg = original.cloneNode(true) as Element,
      ref = child(bg, P, 'bgRef');
    let origin = part;
    const ph = ref?.firstElementChild;
    if (ref && parts[3]) {
      const idx = Number(ref.getAttribute('idx')),
        list = nodes(pkg.doc(parts[3]), A, idx >= 1001 ? 'bgFillStyleLst' : 'fillStyleLst')[0],
        fill = list?.children[idx >= 1001 ? idx - 1001 : idx - 1];
      if (fill) {
        const pr = element(bg.ownerDocument, P, 'p:bgPr');
        pr.append(bg.ownerDocument.importNode(fill, true));
        bg.replaceChildren(pr);
        origin = parts[3];
      }
    }
    for (const c of nodes(bg, A, 'schemeClr')) {
      const from = c.getAttribute('val') === 'phClr' && ph ? ph : c,
        rgb = themeColor(pkg, parts, from);
      if (!rgb) continue;
      const col = element(bg.ownerDocument, A, 'a:srgbClr', { val: rgb });
      if (from !== c)
        for (const transform of Array.from(from.children)) col.append(transform.cloneNode(true));
      col.append(...Array.from(c.children));
      c.replaceWith(col);
    }
    return { bg, origin };
  }
}
function describeFill(pkg: OfficePackage, path: string, bg: Element): SlideBackground {
  const props = child(bg, P, 'bgPr'),
    solid = child(props, A, 'solidFill'),
    grad = child(props, A, 'gradFill'),
    blip = child(props, A, 'blipFill');
  const rgb = solid && child(solid, A, 'srgbClr')?.getAttribute('val');
  let background: SlideBackground = { kind: 'other', label: t('模板背景', 'Template background') };
  if (rgb && /^[\dA-F]{6}$/i.test(rgb) && !solid!.firstElementChild?.children.length)
    background = { kind: 'solid', color: '#' + rgb.toUpperCase() };
  if (grad) {
    const stops = nodes(grad, A, 'gs'),
      colors = stops.map((e) => child(e, A, 'srgbClr')?.getAttribute('val'));
    background =
      stops.length === 2 &&
      colors.every((c) => c && /^[a-f\d]{6}$/i.test(c)) &&
      stops.every(
        (e, i) =>
          Number(e.getAttribute('pos')) === i * 100000 && !e.firstElementChild?.children.length,
      ) &&
      Number(child(grad, A, 'lin')?.getAttribute('ang') ?? -1) === 0
        ? {
            kind: 'gradient',
            from: '#' + colors[0]!.toUpperCase(),
            to: '#' + colors[1]!.toUpperCase(),
          }
        : { kind: 'other', label: t('自定义渐变', 'Custom gradient') };
  }
  if (blip) {
    const picture = nodes(blip, A, 'blip')[0],
      target = picture && pkg.target(path, picture.getAttributeNS(R, 'embed'));
    if (target)
      background = {
        kind: 'image',
        imageId: target,
        transparency: Math.round(
          100 - Number(nodes(blip, A, 'alphaModFix')[0]?.getAttribute('amt') ?? 100000) / 1000,
        ),
      };
  }
  return background;
}
export function readBackground(pkg: OfficePackage, path: string) {
  const own = child(nodes(pkg.doc(path), P, 'cSld')[0], P, 'bg'),
    resolved = resolvedBackground(pkg, path);
  const effective: SlideBackground = resolved
    ? describeFill(pkg, resolved.origin, resolved.bg)
    : { kind: 'solid', color: '#FFFFFF' };
  const background: SlideBackground = own ? describeFill(pkg, path, own) : { kind: 'master' };
  const inherited = own ? resolvedBackground(pkg, path, true) : resolved;
  const masterBackground: SlideBackground = inherited
    ? describeFill(pkg, inherited.origin, inherited.bg)
    : { kind: 'solid', color: '#FFFFFF' };
  return {
    background,
    masterBackground,
    effectiveBackground: effective,
    ...(own || resolved ? { backgroundStamp: xml(own || resolved!.bg) } : {}),
  };
}
export function writeBackground(
  pkg: OfficePackage,
  path: string,
  fill: BackgroundFill,
  images: Record<string, string>,
  cache: Map<string, string>,
) {
  const d = pkg.edit(path),
    cs = nodes(d, P, 'cSld')[0];
  child(cs, P, 'bg')?.remove();
  if (fill.kind === 'master') return;
  const bg = element(d, P, 'p:bg'),
    pr = element(d, P, 'p:bgPr');
  cs.insertBefore(bg, cs.firstChild);
  bg.append(pr);
  if (fill.kind === 'solid') {
    const f = element(d, A, 'a:solidFill');
    f.append(element(d, A, 'a:srgbClr', { val: fill.color.slice(1) }));
    pr.append(f);
  } else if (fill.kind === 'gradient') {
    const f = element(d, A, 'a:gradFill'),
      gs = element(d, A, 'a:gsLst');
    f.append(gs);
    [fill.from, fill.to].forEach((color, i) => {
      const stop = element(d, A, 'a:gs', { pos: i * 100000 });
      stop.append(element(d, A, 'a:srgbClr', { val: color.slice(1) }));
      gs.append(stop);
    });
    f.append(element(d, A, 'a:lin', { ang: 0, scaled: 1 }));
    pr.append(f);
  } else {
    let target = cache.get(fill.imageId);
    if (!target) {
      const data = images[fill.imageId],
        match = data?.match(/^data:image\/(png|jpeg);base64,([A-Za-z0-9+/=]+)$/);
      if (!match) throw Error(t('请重新选择背景图片。', 'Select the background image again.'));
      target = pkg.unique('ppt/media', 'background', match[1] === 'png' ? '.png' : '.jpeg');
      pkg.files[target] = bytes64(match[2]);
      pkg.type(target, 'image/' + match[1]);
      cache.set(fill.imageId, target);
    }
    const id = pkg.relate(path, target, R + '/image'),
      f = element(d, A, 'a:blipFill'),
      b = element(d, A, 'a:blip');
    b.setAttributeNS(R, 'r:embed', id);
    if (fill.transparency)
      b.append(element(d, A, 'a:alphaModFix', { amt: (100 - fill.transparency) * 1000 }));
    const stretch = element(d, A, 'a:stretch');
    stretch.append(element(d, A, 'a:fillRect'));
    f.append(b, stretch);
    pr.append(f);
  }
  pr.append(element(d, A, 'a:effectLst'));
}
export function copyBackground(
  pkg: OfficePackage,
  from: string,
  to: string,
  original: OfficePackage = pkg,
) {
  const resolved = resolvedBackground(original, from);
  const source = resolved?.origin || from,
    bg = resolved?.bg;
  const d = pkg.edit(to),
    cs = nodes(d, P, 'cSld')[0];
  child(cs, P, 'bg')?.remove();
  if (!bg) {
    writeBackground(pkg, to, { kind: 'solid', color: '#FFFFFF' }, {}, new Map());
    return;
  }
  const copy = d.importNode(bg, true);
  for (const e of [copy, ...Array.from(copy.getElementsByTagName('*'))]) {
    for (const a of Array.from(e.attributes).filter((a) => a.namespaceURI === R)) {
      const rel = original.rels(source).find((r) => r.getAttribute('Id') === a.value),
        target = original.target(source, a.value);
      if (!rel || !target)
        throw Error(
          t('无法复制外部背景引用。', 'External background references cannot be copied.'),
        );
      e.setAttributeNS(R, a.name, pkg.relate(to, target, rel.getAttribute('Type')!));
    }
  }
  cs.insertBefore(copy, cs.firstChild);
}
