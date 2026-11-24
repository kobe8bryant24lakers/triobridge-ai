import type { ChartData, ChartKind, MaterialTable } from './types';
import { parseDelimited } from './materials';

export const CHART_LABELS: Record<ChartKind, string> = {
  column: '柱形图',
  bar: '条形图',
  line: '折线图',
  pie: '饼图',
  doughnut: '圆环图',
};
export function tableFromText(value: string): MaterialTable {
  const matrix = parseDelimited(value);
  if (matrix.length < 2 || matrix[0].length < 2)
    throw new Error('请粘贴含表头、分类列及至少一列数值的数据。');
  if (matrix.some((row) => row.length !== matrix[0].length))
    throw new Error('各行列数不一致，请检查分隔符和空单元格。');
  return { id: 'pasted-table', name: '粘贴的数据', headers: matrix[0], rows: matrix.slice(1) };
}
export function numeric(value: string | number | boolean | null, label: string): number | null {
  if (value === null || (typeof value === 'string' && !value.trim())) return null;
  if (typeof value === 'number' && Number.isFinite(value)) return value;
  if (typeof value === 'string') {
    const s = value.trim();
    // Explicit thousands separators and percentage values; never interpret formulas or currency.
    if (/^[+-]?(?:\d{1,3}(?:,\d{3})+|\d+|\d*\.\d+)(?:\.\d+)?(?:[eE][+-]?\d+)?%?$/.test(s)) {
      const n = Number(s.replace(/,/g, '').replace(/%$/, '')) / (s.endsWith('%') ? 100 : 1);
      if (Number.isFinite(n)) return n;
    }
  }
  throw new Error(`${label}不是数值。请在来源中清理单位、公式或错误值后重试。`);
}
export function chartData(
  table: MaterialTable,
  category: number,
  series: number[],
  kind: ChartKind,
): ChartData {
  if (!Object.prototype.hasOwnProperty.call(CHART_LABELS, kind)) throw new Error('图表类型无效。');
  if (
    !Number.isInteger(category) ||
    category < 0 ||
    category >= table.headers.length ||
    !series.length ||
    series.length > 6 ||
    new Set(series).size !== series.length ||
    series.some((i) => !Number.isInteger(i) || i < 0 || i >= table.headers.length || i === category)
  )
    throw new Error('请选择分类列及 1–6 个不同的数值列。');
  if (!table.rows.length || table.rows.length > 30)
    throw new Error('一张图表支持 1–30 个分类，请选择更小的数据区域。');
  const data: ChartData = {
    categories: table.rows.map((r, i) => {
      const label = String(r[category] ?? '').trim();
      if (!label || label.length > 80) throw new Error(`第 ${i + 2} 行的分类名称为空或过长。`);
      return label;
    }),
    series: series.map((index) => ({
      name: table.headers[index],
      values: table.rows.map((row, i) =>
        numeric(row[index] ?? null, `第 ${i + 2} 行「${table.headers[index]}」`),
      ),
    })),
  };
  if (data.series.some((s) => s.values.every((v) => v === null)))
    throw new Error('所选数值列全为空，请检查数据。');
  if (kind === 'pie' || kind === 'doughnut') {
    if (series.length !== 1 || table.rows.length > 8)
      throw new Error('饼图和圆环图需要一个数值列，最多 8 个分类。');
    const values = data.series[0].values;
    if (values.some((v) => v === null || v < 0) || values.every((v) => v === 0))
      throw new Error('饼图和圆环图需要非负完整数据，且至少一项大于零。');
  }
  return data;
}
