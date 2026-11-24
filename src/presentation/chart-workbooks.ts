import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import type { ChartData } from './types';

const C = 'http://schemas.openxmlformats.org/drawingml/2006/chart';
const X = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const nodes = (root: Document | Element, ns: string, name: string) =>
  Array.from(root.getElementsByTagNameNS(ns, name));
function parse(bytes: Uint8Array | undefined) {
  if (!bytes) throw new Error('图表内嵌数据缺失，文件未导出。');
  const doc = new DOMParser().parseFromString(strFromU8(bytes), 'application/xml');
  if (doc.getElementsByTagName('parsererror').length) throw new Error('图表文件结构无效。');
  return doc;
}
const serialize = (doc: Document) => strToU8(new XMLSerializer().serializeToString(doc));

/** Repair PptxGenJS 4.0.1's zero-to-blank worksheet serialization using the
 * validated source data. Also omit missing numeric cache points, never convert
 * them to zero. This keeps Edit Data, cached rendering and input data identical. */
export function preserveChartData(base64: string, data: ChartData[]): string {
  if (!data.length) return base64;
  const zip = unzipSync(Uint8Array.from(atob(base64), (c) => c.charCodeAt(0)));
  const paths = Object.keys(zip)
    .filter((n) => /^ppt\/charts\/chart\d+\.xml$/.test(n))
    .sort((a, b) => Number(a.match(/chart(\d+)/)![1]) - Number(b.match(/chart(\d+)/)![1]));
  if (paths.length !== data.length) throw new Error('图表数量与源数据不一致，文件未导出。');
  paths.forEach((path, i) => {
    const chart = parse(zip[path]),
      series = nodes(chart, C, 'ser'),
      source = data[i];
    if (series.length !== source.series.length) throw new Error('图表系列与源数据不一致。');
    const rels = parse(zip[path.replace(/([^/]+)$/, '_rels/$1.rels')]);
    const externalId = nodes(chart, C, 'externalData')[0]?.getAttributeNS(R, 'id');
    const target = Array.from(rels.documentElement.children)
      .find((r) => r.getAttribute('Id') === externalId)
      ?.getAttribute('Target');
    if (!target || !/^\.\.\/embeddings\/[\w-]+\.xlsx$/.test(target))
      throw new Error('图表工作簿引用无效。');
    const workbookPath = `ppt/${target.slice(3)}`,
      workbook = unzipSync(zip[workbookPath]);
    const sheet = parse(workbook['xl/worksheets/sheet1.xml']);
    const cells = new Map(nodes(sheet, X, 'c').map((c) => [c.getAttribute('r'), c]));
    series.forEach((ser, j) => {
      const val = nodes(ser, C, 'val')[0],
        ref = val && nodes(val, C, 'f')[0]?.textContent;
      const range = ref?.match(/^Sheet1!\$([A-Z]+)\$2:\$\1\$(\d+)$/);
      if (!range || Number(range[2]) !== source.categories.length + 1)
        throw new Error('图表数据区域与源数据不一致。');
      const cache = nodes(val, C, 'numCache')[0];
      if (!cache) throw new Error('图表缓存缺失。');
      nodes(cache, C, 'pt').forEach((p) => p.remove());
      source.series[j].values.forEach((value, row) => {
        const cell = cells.get(`${range[1]}${row + 2}`);
        if (!cell) throw new Error('图表工作簿单元格缺失。');
        nodes(cell, X, 'v').forEach((v) => v.remove());
        cell.removeAttribute('t');
        if (value !== null) {
          const v = sheet.createElementNS(X, 'v');
          v.textContent = String(value);
          cell.append(v);
          const pt = chart.createElementNS(C, 'c:pt'),
            cv = chart.createElementNS(C, 'c:v');
          pt.setAttribute('idx', String(row));
          cv.textContent = String(value);
          pt.append(cv);
          cache.append(pt);
        }
      });
    });
    workbook['xl/worksheets/sheet1.xml'] = serialize(sheet);
    zip[workbookPath] = zipSync(workbook);
    zip[path] = serialize(chart);
  });
  const bytes = zipSync(zip),
    chunks: string[] = [];
  for (let start = 0; start < bytes.length; start += 0x8000)
    chunks.push(String.fromCharCode(...bytes.subarray(start, start + 0x8000)));
  return btoa(chunks.join(''));
}
