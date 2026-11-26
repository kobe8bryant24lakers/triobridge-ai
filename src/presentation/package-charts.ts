import { t } from '../i18n';
import type { ChartData, ChartKind, ShapeState } from './types';
import { C, X, R, OfficePackage, child, element, nodes, unpack, need } from './package-xml';

function points(container: Element | undefined) {
  if (!container) return [];
  const count = Number(nodes(container, C, 'ptCount')[0]?.getAttribute('val') ?? 1);
  if (!Number.isInteger(count) || count < 1 || count > 10000) return [];
  const values: (string | null)[] = Array(count).fill(null);
  for (const p of nodes(container, C, 'pt')) {
    const i = Number(p.getAttribute('idx'));
    if (i >= 0 && i < count) values[i] = child(p, C, 'v')?.textContent ?? null;
  }
  return values;
}
function bookPath(pkg: OfficePackage, path: string) {
  return pkg.target(
    path,
    nodes(pkg.doc(path), C, 'externalData')[0]?.getAttributeNS(R, 'id') ?? null,
  );
}
export function readChart(pkg: OfficePackage, path: string): ShapeState['chart'] {
  if (!bookPath(pkg, path)) return;
  const plot = nodes(pkg.doc(path), C, 'plotArea')[0];
  const groups = plot && Array.from(plot.children).filter((e) => e.localName.endsWith('Chart'));
  if (!groups || groups.length !== 1) return;
  const group = groups[0];
  const kinds: Record<string, ChartKind> = {
    lineChart: 'line',
    pieChart: 'pie',
    doughnutChart: 'doughnut',
    barChart: child(group, C, 'barDir')?.getAttribute('val') === 'bar' ? 'bar' : 'column',
  };
  const kind = kinds[group.localName],
    series = nodes(group, C, 'ser');
  if (!kind || !series.length || series.length > 6) return;
  const cat = child(series[0], C, 'cat');
  if (!cat || nodes(cat, C, 'lvl').length > 1 || nodes(cat, C, 'numRef').length) return;
  const categories = points(cat);
  if (categories.length > 30 || !categories.length || !categories.every((v) => v !== null)) return;
  const result = series.map((ser, i) => {
    const tx = child(ser, C, 'tx'),
      values = points(child(ser, C, 'val'));
    return {
      name: (tx && nodes(tx, C, 'v')[0]?.textContent) || t('系列 ', 'Series ') + (i + 1),
      values: values.map((v) => (v === null ? null : Number(v))),
    };
  });
  if (
    result.some(
      (s) =>
        s.values.length !== categories.length ||
        s.values.some((v) => v !== null && !Number.isFinite(v)),
    )
  )
    return;
  return { kind, data: { categories: categories as string[], series: result } };
}
function column(n: number) {
  let name = '';
  for (n++; n > 0; n = Math.floor((n - 1) / 26))
    name = String.fromCharCode(65 + ((n - 1) % 26)) + name;
  return name;
}
function sheetFormula(name: string, col: number, start: number, end = start) {
  const sheet = /^[A-Za-z_][A-Za-z_0-9]*$/.test(name) ? name : "'" + name.replace(/'/g, "''") + "'";
  const ref = '$' + column(col) + '$';
  return sheet + '!' + ref + start + (end !== start ? ':' + ref + end : '');
}
function workbookData(book: OfficePackage, old: ChartData, data: ChartData, series: Element[]) {
  const main =
    book.target(
      '',
      book
        .rels('')
        .find((r) => r.getAttribute('Type') === R + '/officeDocument')
        ?.getAttribute('Id') ?? null,
    ) || 'xl/workbook.xml';
  const wb = book.edit(main),
    sheets = nodes(wb, X, 'sheet');
  // Reuse a conventional chart data sheet when it contains exclusively the bound grid.
  // A separate worksheet avoids overwriting unrelated cells, formulas or horizontal ranges.
  let sheet = sheets.find((s) =>
    series.every((item, i) => {
      const formula = nodes(child(item, C, 'val')!, C, 'f')[0]?.textContent?.replace(/'/g, '');
      const expected = sheetFormula(
        s.getAttribute('name')!,
        i + 1,
        2,
        old.categories.length + 1,
      ).replace(/'/g, '');
      const cat = nodes(child(item, C, 'cat')!, C, 'f')[0]?.textContent?.replace(/'/g, '');
      return (
        (formula === expected ||
          (old.categories.length === 1 && formula === expected + ':$' + column(i + 1) + '$2')) &&
        cat?.split('!')[1]?.startsWith('$A$2')
      );
    }),
  );
  let path = sheet && book.target(main, sheet.getAttributeNS(R, 'id'));
  if (
    path &&
    nodes(book.doc(path), X, 'c').some((c) => {
      const m = /^([A-Z]+)(\d+)$/.exec(c.getAttribute('r') || '');
      return (
        !m ||
        nodes(c, X, 'f').length > 0 ||
        Number(m[2]) > old.categories.length + 1 ||
        Array.from({ length: old.series.length + 1 }, (_, i) => column(i)).indexOf(m[1]) < 0
      );
    })
  )
    path = undefined;
  if (!path) {
    let name = 'Chart Data',
      n = 1;
    while (sheets.some((s) => s.getAttribute('name') === name)) name = 'Chart Data ' + ++n;
    path = book.unique('xl/worksheets', 'sheet', '.xml');
    const d = document.implementation.createDocument(X, 'worksheet');
    d.documentElement.append(element(d, X, 'sheetData'));
    book.saveDoc(path, d);
    book.type(path, 'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml');
    sheet = element(wb, X, 'sheet', {
      name,
      sheetId: Math.max(0, ...sheets.map((s) => Number(s.getAttribute('sheetId')))) + 1,
    });
    sheet.setAttributeNS(R, 'r:id', book.relate(main, path, R + '/worksheet'));
    nodes(wb, X, 'sheets')[0].append(sheet);
    nodes(wb, X, 'workbookView').forEach((v) => v.setAttribute('activeTab', String(sheets.length)));
  }
  const sheetName = sheet!.getAttribute('name')!,
    d = book.edit(path),
    grid = nodes(d, X, 'sheetData')[0],
    oldCells = new Map(nodes(grid, X, 'c').map((c) => [c.getAttribute('r'), c]));
  const last = column(data.series.length) + (data.categories.length + 1),
    range = 'A1:' + last;
  // Keep per-cell number formats, including percentages. The chart retains its axis and label formats.
  const header = oldCells.get('A1');
  let firstHeader = header?.textContent || '';
  if (header?.getAttribute('t') === 's') {
    const strs = book.rels(main).find((r) => r.getAttribute('Type') === R + '/sharedStrings');
    const sp = strs && book.target(main, strs.getAttribute('Id'));
    firstHeader = sp ? nodes(book.doc(sp), X, 'si')[Number(firstHeader)]?.textContent || '' : '';
  }
  const rows: (string | number | null)[][] = [
    [firstHeader || t('分类', 'Category'), ...data.series.map((s) => s.name)],
    ...data.categories.map((v, i) => [v, ...data.series.map((s) => s.values[i])]),
  ];
  grid.replaceChildren();
  rows.forEach((line, i) => {
    const row = element(d, X, 'row', { r: i + 1 });
    grid.append(row);
    line.forEach((value, j) => {
      const ref = column(j) + (i + 1),
        cell = element(d, X, 'c', { r: ref }),
        style = (oldCells.get(ref) || oldCells.get(column(j) + '2'))?.getAttribute('s');
      if (style) cell.setAttribute('s', style);
      row.append(cell);
      if (typeof value === 'string') {
        cell.setAttribute('t', 'inlineStr');
        const v = element(d, X, 't');
        v.textContent = value;
        const str = element(d, X, 'is');
        str.append(v);
        cell.append(str);
      } else if (value !== null) {
        const v = element(d, X, 'v');
        v.textContent = String(value);
        cell.append(v);
      }
    });
  });
  nodes(d, X, 'dimension').forEach((e) => e.setAttribute('ref', range));
  for (const rel of book.rels(path).filter((r) => r.getAttribute('Type') === R + '/table')) {
    const tablePath = book.target(path, rel.getAttribute('Id'))!;
    const td = book.edit(tablePath),
      table = td.documentElement;
    if (
      table.getAttribute('ref') !==
      'A1:' + column(old.series.length) + (old.categories.length + 1)
    )
      continue;
    table.setAttribute('ref', range);
    child(table, X, 'autoFilter')?.setAttribute('ref', range);
    const cols = child(table, X, 'tableColumns');
    if (cols) {
      cols.replaceChildren();
      cols.setAttribute('count', String(rows[0].length));
      rows[0].forEach((name, i) =>
        cols.append(element(td, X, 'tableColumn', { id: i + 1, name: String(name) })),
      );
    }
  }
  return sheetName;
}
function cacheData(
  ser: Element,
  key: 'tx' | 'cat' | 'val',
  values: (string | number | null)[],
  formula: string,
) {
  const d = ser.ownerDocument,
    container = child(ser, C, key)!,
    numeric = key === 'val';
  const format = nodes(container, C, 'formatCode')[0]?.textContent || 'General';
  const ref = element(d, C, numeric ? 'c:numRef' : 'c:strRef'),
    f = element(d, C, 'c:f'),
    cache = element(d, C, numeric ? 'c:numCache' : 'c:strCache');
  f.textContent = formula;
  ref.append(f, cache);
  container.replaceChildren(ref);
  if (numeric) {
    const code = element(d, C, 'c:formatCode');
    code.textContent = format;
    cache.append(code);
  }
  cache.append(element(d, C, 'c:ptCount', { val: values.length }));
  values.forEach((v, i) => {
    if (v === null) return;
    const pt = element(d, C, 'c:pt', { idx: i }),
      value = element(d, C, 'c:v');
    value.textContent = String(v);
    pt.append(value);
    cache.append(pt);
  });
}
export function refreshChart(pkg: OfficePackage, path: string, data: ChartData) {
  const current = readChart(pkg, path),
    workbook = bookPath(pkg, path);
  if (!current || !workbook)
    throw Error(t('此图表无法写入内嵌工作簿。', 'This chart has no supported editable workbook.'));
  const doc = pkg.edit(path),
    series = nodes(doc, C, 'ser'),
    book = new OfficePackage(unpack(pkg.files[workbook]));
  const sheet = workbookData(book, current.data, data, series);
  const shared = pkg
    .paths()
    .some(
      (other) =>
        other !== path &&
        pkg.contentType(other) ===
          'application/vnd.openxmlformats-officedocument.drawingml.chart+xml' &&
        bookPath(pkg, other) === workbook,
    );
  let destination = workbook;
  if (shared) {
    destination = pkg.unique(workbook.slice(0, workbook.lastIndexOf('/')), 'Workbook', '.xlsx');
    pkg.type(destination, pkg.contentType(workbook)!);
    nodes(doc, C, 'externalData')[0].setAttributeNS(
      R,
      'r:id',
      pkg.relate(path, destination, R + '/package'),
    );
  }
  pkg.files[destination] = book.pack();
  const last = series[series.length - 1],
    group = last.parentElement!,
    after = last.nextSibling;
  let index = Math.max(...series.map((s) => Number(child(s, C, 'idx')?.getAttribute('val')) || 0));
  data.series.forEach((s, i) => {
    const ser = series[i] || (last.cloneNode(true) as Element);
    if (!series[i]) {
      group.insertBefore(ser, after);
      child(ser, C, 'idx')?.setAttribute('val', String(++index));
      child(ser, C, 'order')?.setAttribute('val', String(i));
      nodes(ser, C, 'dPt').forEach((e) => e.remove());
    }
    if (!child(ser, C, 'tx')) {
      const tx = element(doc, C, 'c:tx');
      ser.insertBefore(tx, child(ser, C, 'order')?.nextSibling || null);
    }
    cacheData(ser, 'tx', [s.name], sheetFormula(sheet, i + 1, 1));
    cacheData(ser, 'cat', data.categories, sheetFormula(sheet, 0, 2, data.categories.length + 1));
    cacheData(ser, 'val', s.values, sheetFormula(sheet, i + 1, 2, data.categories.length + 1));
    for (const dp of [...nodes(ser, C, 'dPt'), ...nodes(ser, C, 'dLbl')])
      if (Number(child(dp, C, 'idx')?.getAttribute('val')) >= data.categories.length) dp.remove();
  });
  series.slice(data.series.length).forEach((e) => e.remove());
  const ext = nodes(doc, C, 'externalData')[0];
  need(ext, C, 'c:autoUpdate').setAttribute('val', '0');
}
