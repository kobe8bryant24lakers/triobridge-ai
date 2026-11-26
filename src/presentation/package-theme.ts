import { A, P, R, OfficePackage, child, nodes } from './package-xml';

export function hierarchy(pkg: OfficePackage, slide: string) {
  const parts = [slide];
  for (const type of ['slideLayout', 'slideMaster', 'theme']) {
    const part = parts[parts.length - 1],
      rel = pkg.rels(part).find((r) => r.getAttribute('Type') === R + '/' + type),
      next = rel && pkg.target(part, rel.getAttribute('Id'));
    if (!next) break;
    parts.push(next);
  }
  return parts;
}
const defaultMap: Record<string, string> = { bg1: 'lt1', tx1: 'dk1', bg2: 'lt2', tx2: 'dk2' };
export function themeColor(
  pkg: OfficePackage,
  parts: string[],
  color: Element,
): string | undefined {
  if (color.localName === 'srgbClr') return color.getAttribute('val') || undefined;
  if (color.localName === 'sysClr') return color.getAttribute('lastClr') || undefined;
  if (color.localName !== 'schemeClr' || !parts[3]) return;
  const key = color.getAttribute('val') || '';
  const mapped =
    parts
      .slice(0, 3)
      .map((p) =>
        (
          nodes(pkg.doc(p), A, 'overrideClrMapping')[0] || nodes(pkg.doc(p), P, 'clrMap')[0]
        )?.getAttribute(key),
      )
      .find(Boolean) ||
    defaultMap[key] ||
    key;
  const palette = nodes(pkg.doc(parts[3]), A, 'clrScheme')[0],
    chosen = child(palette, A, mapped)?.firstElementChild;
  return chosen?.getAttribute(chosen.localName === 'sysClr' ? 'lastClr' : 'val') || undefined;
}
export function themeFont(
  pkg: OfficePackage,
  parts: string[],
  name: string | undefined | null,
  role: string,
) {
  if (name && !name.startsWith('+')) return name;
  if (!parts[3]) return null;
  const group = nodes(
    pkg.doc(parts[3]),
    A,
    (name ? name.startsWith('+mj') : role === 'title') ? 'majorFont' : 'minorFont',
  )[0];
  return child(group, A, 'latin')?.getAttribute('typeface') || null;
}
