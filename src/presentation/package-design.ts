import type { DeckTheme } from './design-themes';
import { chartPalette, readableOn } from './design-themes';
import { A, C, P, OfficePackage, child, nodes, element, need } from './package-xml';
import { hierarchy, themeColor } from './package-theme';
import { chartPart, shapeNodes, idOf, writeShape } from './package-shapes';
import { writeBackground } from './package-background';
import type { PresentationSnapshot } from './types';

function fillColor(parent: Element | undefined, value: string) {
  if (!parent) return;
  const fill = child(parent, A, 'solidFill');
  if (!fill) return;
  fill.replaceChildren(element(parent.ownerDocument, A, 'a:srgbClr', { val: value.slice(1) }));
}
function typeStyle(e: Element, font: string, ink: string) {
  for (const r of ['rPr', 'defRPr', 'endParaRPr'].flatMap((tag) => nodes(e, A, tag))) {
    for (const f of ['latin', 'ea', 'cs']) {
      const face = child(r, A, f);
      if (face) face.setAttribute('typeface', font);
    }
    fillColor(r, ink);
  }
}
/** Work on the isolated package. Text runs, data caches, embeds and slide relationships stay intact. */
export function restylePackage(
  pkg: OfficePackage,
  paths: string[],
  snapshot: PresentationSnapshot,
  style: DeckTheme,
) {
  const colors = chartPalette(style),
    schemes = new Set<string>(),
    charts = new Set<string>();
  paths.forEach((path, index) => {
    const slide = snapshot.slides[index],
      parts = hierarchy(pkg, path);
    if (parts[3]) schemes.add(parts[3]);
    writeBackground(pkg, path, { kind: 'solid', color: style.background }, {}, new Map());
    const backdrops: {
      left: number;
      top: number;
      width: number;
      height: number;
      color?: string;
    }[] = [];
    for (const shape of shapeNodes(pkg.edit(path))) {
      const info = slide.shapes.find((s) => s.id === idOf(shape));
      if (!info) continue;
      const opaque = child(child(shape, P, 'spPr'), A, 'solidFill')?.firstElementChild;
      const shapeFill = opaque && themeColor(pkg, parts, opaque);
      const colorBehind = backdrops
        .slice()
        .reverse()
        .find(
          (b) =>
            info.left >= b.left - 1 &&
            info.top >= b.top - 1 &&
            info.left + info.width <= b.left + b.width + 1 &&
            info.top + info.height <= b.top + b.height + 1,
        );
      if (info.text) {
        const surface = shapeFill ? '#' + shapeFill : colorBehind?.color || style.background;
        const ink =
          info.role === 'title'
            ? readableOn(surface, style.accent)
            : readableOn(surface, style.ink);
        writeShape(
          shape,
          {
            slideId: slide.id,
            shapeId: info.id,
            fontName: info.role === 'title' ? style.headingFont : style.font,
            ...(colorBehind && !colorBehind.color && !shapeFill ? {} : { fontColor: ink }),
          },
          info,
        );
      } else if (shapeFill && info.type === 'Shape') {
        fillColor(child(shape, P, 'spPr'), style.accent);
        backdrops.push({ ...info, color: style.accent });
      } else if (info.type === 'Image') backdrops.push({ ...info });
      const line = child(child(shape, P, 'spPr'), A, 'ln');
      if (line && !info.text) fillColor(line, style.accent);
      const table = nodes(shape, A, 'tbl')[0];
      if (table) {
        const firstRow = child(table, A, 'tblPr')?.getAttribute('firstRow');
        nodes(table, A, 'tr').forEach((row, i) => {
          const header = i === 0 && firstRow !== '0' && firstRow !== 'false';
          const surface = header ? style.accent : style.background;
          for (const cell of nodes(row, A, 'tc')) {
            const props = child(cell, A, 'tcPr');
            if (props) {
              const solid = child(props, A, 'solidFill');
              if (solid) fillColor(props, surface);
            }
            typeStyle(cell, style.font, readableOn(surface, header ? '#FFFFFF' : style.ink));
          }
        });
      }
      const chart = chartPart(pkg, path, shape);
      if (chart) charts.add(chart);
    }
  });
  for (const part of charts) {
    const doc = pkg.edit(part),
      root = doc.documentElement;
    typeStyle(root, style.font, style.ink);
    for (const [i, ser] of nodes(doc, C, 'ser').entries()) {
      const props = child(ser, C, 'spPr');
      fillColor(props, colors[i % colors.length]);
      fillColor(child(props, A, 'ln'), colors[i % colors.length]);
      for (const point of nodes(ser, C, 'dPt')) {
        const n = Number(child(point, C, 'idx')?.getAttribute('val')) || 0;
        fillColor(child(point, C, 'spPr'), colors[n % colors.length]);
      }
    }
  }
  for (const part of schemes) {
    const doc = pkg.edit(part),
      group = nodes(doc, A, 'clrScheme')[0];
    const palette: Record<string, string> = {
      dk1: style.ink,
      lt1: style.background,
      dk2: style.ink,
      lt2: style.background,
      ...Object.fromEntries(colors.map((c, i) => ['accent' + (i + 1), c])),
    };
    for (const el of Array.from(group?.children || []))
      if (palette[el.localName])
        el.replaceChildren(element(doc, A, 'a:srgbClr', { val: palette[el.localName].slice(1) }));
    for (const [tag, font] of [
      ['majorFont', style.headingFont],
      ['minorFont', style.font],
    ]) {
      const family = nodes(doc, A, tag)[0];
      if (family)
        for (const name of ['latin', 'ea', 'cs'])
          need(family, A, 'a:' + name).setAttribute('typeface', font);
    }
  }
}
