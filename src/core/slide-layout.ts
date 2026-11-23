/**
 * Choose a conservative starting size before PowerPoint's final text fitting.
 * Native auto-fit alone also compresses line spacing, which crowds dense CJK
 * paragraphs. This estimate reserves leading and text-frame insets; Office
 * still handles actual font metrics, word wrapping, and unusual glyphs.
 */
export function slideFontSize(
  text: string,
  width: number,
  height: number,
  preferred: number,
  minimum: number,
): number {
  const usableWidth = Math.max(1, width - 16);
  const usableHeight = Math.max(1, height - 12);
  const widths = text
    .split('\n')
    .map((line) =>
      Array.from(line).reduce((sum, char) => sum + (/^[\x20-\x7e]$/.test(char) ? 0.55 : 1), 0),
    );
  for (let size = preferred; size > minimum; size -= 0.5) {
    const lines = widths.reduce(
      (sum, units) => sum + Math.max(1, Math.ceil((units * size) / usableWidth)),
      0,
    );
    if (lines * size * 1.4 <= usableHeight) return size;
  }
  return minimum;
}
