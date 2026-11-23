import { parseFormula } from './output';
import { MAX_RESULT_LENGTH } from './tasks';
import type { ColumnFormat, ReportContent, ReportSection, TableCell, TableContent } from './types';

function jsonObject(raw: string): Record<string, unknown> {
  if (raw.length > MAX_RESULT_LENGTH)
    throw new Error('结构化结果超过 100,000 字符，请缩小任务范围。');
  try {
    const value = JSON.parse(raw.trim().replace(/^```(?:json)?\s*\n([\s\S]*?)\n```$/i, '$1'));
    if (value && typeof value === 'object' && !Array.isArray(value)) return value;
  } catch {
    /* The editable result remains available to repair. */
  }
  throw new Error('结果需要有效的 JSON 对象。可编辑修正，或重新生成。');
}
function text(value: unknown, limit: number, label: string): string {
  if (typeof value !== 'string' || !value.trim() || value.length > limit)
    throw new Error(`${label}需要 1–${limit} 字符。`);
  return value.trim();
}
function knownFields(value: Record<string, unknown>, fields: string[], label: string) {
  const extra = Object.keys(value).filter((key) => !fields.includes(key));
  if (extra.length)
    throw new Error(
      `${label}包含暂不支持的字段：${extra.join('、')}。请将内容移入正文或表格后重试。`,
    );
}
function texts(value: unknown, maxItems: number, maxLength: number, label: string): string[] {
  if (!Array.isArray(value) || value.length > maxItems)
    throw new Error(`${label}最多 ${maxItems} 项。`);
  return value.map((item, i) => text(item, maxLength, `${label}第 ${i + 1} 项`));
}
function headers(value: unknown, max: number, label: string): string[] {
  const result = texts(value, max, 80, label);
  if (
    !result.length ||
    new Set(result.map((item) => item.toLocaleLowerCase())).size !== result.length
  )
    throw new Error(`${label}需要 1–${max} 个不重复的列名。`);
  return result;
}

export function parseReport(raw: string): ReportContent {
  const data = jsonObject(raw);
  knownFields(data, ['title', 'sections'], '报告');
  const title = text(data.title, 120, '文档标题');
  if (!Array.isArray(data.sections) || !data.sections.length || data.sections.length > 12)
    throw new Error('报告需要 1–12 个章节。');
  const sections = data.sections.map((value: unknown, i): ReportSection => {
    const label = `第 ${i + 1} 节`;
    if (!value || typeof value !== 'object' || Array.isArray(value))
      throw new Error(`${label}格式无效。`);
    const section = value as Record<string, unknown>;
    knownFields(section, ['heading', 'paragraphs', 'bullets', 'table'], label);
    const result: ReportSection = {
      heading: text(section.heading, 100, `${label}标题`),
      paragraphs: texts(section.paragraphs ?? [], 8, 4000, `${label}段落`),
      bullets: texts(section.bullets ?? [], 12, 500, `${label}要点`),
    };
    if (section.table !== undefined) {
      if (!section.table || typeof section.table !== 'object' || Array.isArray(section.table))
        throw new Error(`${label}表格格式无效。`);
      const table = section.table as Record<string, unknown>;
      knownFields(table, ['headers', 'rows'], `${label}表格`);
      const columns = headers(table.headers, 8, `${label}表头`);
      if (!Array.isArray(table.rows) || !table.rows.length || table.rows.length > 60)
        throw new Error(`${label}表格需要 1–60 行数据。`);
      const rows = table.rows.map((row, index) => {
        if (
          !Array.isArray(row) ||
          row.length !== columns.length ||
          !row.every((cell) => typeof cell === 'string' && cell.length <= 500)
        )
          throw new Error(
            `${label}表格第 ${index + 1} 行需要 ${columns.length} 个文本单元格，每格最多 500 字符。`,
          );
        return row as string[];
      });
      result.table = { headers: columns, rows };
    }
    if (!result.paragraphs.length && !result.bullets.length && !result.table)
      throw new Error(`${label}需要正文、要点或表格。`);
    return result;
  });
  return { title, sections };
}

export function parseTable(raw: string): TableContent {
  const data = jsonObject(raw);
  knownFields(data, ['name', 'headers', 'rows', 'formats'], '数据表');
  const name = text(data.name, 31, '工作表名称');
  if (/[\\/*?:\[\]\u0000-\u001f]/.test(name) || name.startsWith("'") || name.endsWith("'"))
    throw new Error('工作表名称不能包含 \\ / * ? : [ ]、控制字符，或以单引号开头结尾。');
  const columns = headers(data.headers, 12, '表头');
  if (
    !Array.isArray(data.rows) ||
    !data.rows.length ||
    data.rows.length > 200 ||
    (data.rows.length + 1) * columns.length > 2000
  )
    throw new Error('数据表需要 1–200 行，含表头最多 2,000 个单元格。');
  const rows = data.rows.map((row, r): TableCell[] => {
    if (!Array.isArray(row) || row.length !== columns.length)
      throw new Error(`第 ${r + 1} 行需要 ${columns.length} 个单元格，请补齐空值或移除多余列。`);
    return row.map((cell, c): TableCell => {
      const label = `第 ${r + 1} 行第 ${c + 1} 列`;
      if (cell === null || typeof cell === 'boolean') return cell;
      if (typeof cell === 'number' && Number.isFinite(cell)) return cell;
      if (typeof cell === 'string' && cell.length <= 4000) return cell;
      if (
        cell &&
        typeof cell === 'object' &&
        Object.keys(cell).length === 1 &&
        typeof cell.formula === 'string'
      ) {
        try {
          return { formula: parseFormula(cell.formula) };
        } catch (error) {
          throw new Error(`${label}：${(error as Error).message}`);
        }
      }
      throw new Error(`${label}需要文本、数字、布尔值、null 或 {"formula":"=公式"}。`);
    });
  });
  const validFormats = ['general', 'text', 'integer', 'number', 'percent'];
  const formats = data.formats ?? columns.map(() => 'general');
  if (
    !Array.isArray(formats) ||
    formats.length !== columns.length ||
    formats.some((format) => !validFormats.includes(format))
  )
    throw new Error('列格式需要与表头一一对应，可选 general、text、integer、number、percent。');
  rows.forEach((row, r) =>
    row.forEach((cell, c) => {
      if (formats[c] === 'text' && cell && typeof cell === 'object')
        throw new Error(
          `第 ${r + 1} 行第 ${c + 1} 列包含公式，请将该列格式改为 general 或数字格式。`,
        );
    }),
  );
  return { name, headers: columns, rows, formats: formats as ColumnFormat[] };
}

export function cellText(cell: TableCell): string {
  return cell !== null && typeof cell === 'object' ? cell.formula : String(cell ?? '');
}

/** Keep explicit, consecutive list numbers as authored; never strip model text. */
export function hasExplicitListNumbers(items: string[]): boolean {
  if (!items.length) return false;
  const prefixes = items.map((item) => {
    const match = item.match(
      /^\s*(?:(\d{1,3})([.．])(?=\s+\S)|(\d{1,3})([、)）])(?=\s*\S)|([（(])(\d{1,3})[）)](?=\s*\S))/,
    );
    return match
      ? {
          number: Number(match[1] ?? match[3] ?? match[6]),
          style: match[2] ?? match[4] ?? match[5],
        }
      : null;
  });
  const first = prefixes[0];
  return (
    !!first &&
    prefixes.every(
      (prefix, index) =>
        !!prefix && prefix.style === first.style && prefix.number === first.number + index,
    )
  );
}

export function reportToText(report: ReportContent): string {
  return [
    report.title,
    ...report.sections.map((section) =>
      [
        section.heading,
        ...section.paragraphs,
        ...section.bullets.map((bullet) =>
          hasExplicitListNumbers(section.bullets) ? bullet : `• ${bullet}`,
        ),
        ...(section.table
          ? [section.table.headers.join('\t'), ...section.table.rows.map((row) => row.join('\t'))]
          : []),
      ].join('\n\n'),
    ),
  ].join('\n\n');
}
export function tableToText(table: TableContent): string {
  return [table.headers, ...table.rows]
    .map((row) =>
      row.map((cell) =>
        typeof cell === 'string' && /^\s*[=+\-@]/.test(cell) ? `'${cell}` : cellText(cell),
      ),
    )
    .map((row) =>
      row
        .map((cell) => (/[\t\r\n"]/.test(cell) ? `"${cell.replace(/"/g, '""')}"` : cell))
        .join('\t'),
    )
    .join('\n');
}
function escapeHtml(value: string): string {
  return value
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;')
    .replace(/\r?\n/g, '<br>');
}
export function reportToHtml(report: ReportContent): string {
  const border = 'border:1px solid #d9d9d9;padding:6pt;vertical-align:middle;';
  const keepNext = 'page-break-after:avoid;';
  const paragraphStyle = 'font-size:11pt;line-height:1.5;';
  return `<html><body><h1 style="font-size:22pt;color:#000">${escapeHtml(report.title)}</h1>${report.sections
    .map((section) => {
      const compact = section.table ? compactReportTable(section.table) : false;
      const numbered = hasExplicitListNumbers(section.bullets);
      const numberedHeights = numbered
        ? section.bullets.map((bullet) => estimatedTextLines(bullet, (440 - 18) / 11) * 16.5 + 8)
        : [];
      const compactNumbered =
        numbered &&
        !section.table &&
        section.bullets.length <= 6 &&
        numberedHeights.reduce((total, height) => total + height, 0) <= 280;
      const bullets = section.bullets
        .map((bullet, index) => {
          const style =
            (compact && index === section.bullets.length - 1) ||
            (compactNumbered && index < section.bullets.length - 1)
              ? keepNext
              : '';
          return numbered
            ? `<p style="${paragraphStyle}margin:4pt 0 4pt 18pt;${numberedHeights[index] <= 280 ? 'page-break-inside:avoid;' : ''}${style}">${escapeHtml(bullet)}</p>`
            : `<li${style ? ` style="${style}"` : ''}>${escapeHtml(bullet)}</li>`;
        })
        .join('');
      let table = '';
      if (section.table) {
        const rowStyle = 'page-break-inside:avoid;';
        const cell = (text: string, header: boolean, keep: boolean) => {
          const tag = header ? 'th' : 'td';
          return `<${tag} style="${border}${header ? 'background:#e8eef2;color:#000;' : ''}"><p style="margin:0;font-size:10pt;${keep ? keepNext : ''}">${escapeHtml(text)}</p></${tag}>`;
        };
        table = `<table style="border-collapse:collapse;width:100%;font-size:10pt;${compact ? 'page-break-inside:avoid;' : ''}"><thead><tr style="${rowStyle}">${section.table.headers.map((header) => cell(header, true, compact)).join('')}</tr></thead><tbody>${section.table.rows.map((row, index) => `<tr style="${rowStyle}">${row.map((value) => cell(value, false, compact && index < section.table!.rows.length - 1)).join('')}</tr>`).join('')}</tbody></table><p></p>`;
      }
      return (
        `<h2 style="font-size:16pt;color:#000;${keepNext}">${escapeHtml(section.heading)}</h2>` +
        section.paragraphs
          .map(
            (paragraph, index) =>
              `<p style="${paragraphStyle}${((compact && !section.bullets.length) || compactNumbered) && index === section.paragraphs.length - 1 ? keepNext : ''}">${escapeHtml(paragraph)}</p>`,
          )
          .join('') +
        (numbered ? bullets : bullets ? `<ul>${bullets}</ul>` : '') +
        table
      );
    })
    .join('')}</body></html>`;
}

function compactReportTable(
  table: NonNullable<ReportContent['sections'][number]['table']>,
): boolean {
  if (table.rows.length > 6 || table.headers.length > 6) return false;
  // Conservative estimate for 10 pt text within a typical 440 pt Word content area.
  // Dense or tall tables must remain free to paginate; never chain an arbitrary report.
  const lineWidth = Math.max(1, (440 / table.headers.length - 12) / 10);
  const rowHeight = (row: string[]) =>
    Math.max(...row.map((cell) => estimatedTextLines(cell, lineWidth))) * 14 + 12;
  return [table.headers, ...table.rows].reduce((height, row) => height + rowHeight(row), 0) <= 280;
}

function estimatedTextLines(text: string, lineWidth: number): number {
  return text.split(/\r?\n/).reduce((lines, line) => {
    const width = [...line].reduce(
      (total, character) => total + (/^[\x00-\x7f]$/.test(character) ? 0.6 : 1),
      0,
    );
    return lines + Math.max(1, Math.ceil(width / lineWidth));
  }, 0);
}

/** Prefix text with an apostrophe so Excel never evaluates model-supplied string cells. */
export function tableMatrix(table: TableContent): (string | number | boolean)[][] {
  return [table.headers, ...table.rows].map((row) =>
    row.map((cell) => {
      if (cell !== null && typeof cell === 'object') return cell.formula;
      if (typeof cell === 'string') return cell ? `'${cell}` : '';
      return cell ?? '';
    }),
  );
}
export const NUMBER_FORMATS: Record<ColumnFormat, string> = {
  general: 'General',
  text: '@',
  integer: '#,##0',
  number: '#,##0.00',
  percent: '0.00%',
};

export function columnName(index: number): string {
  let value = index + 1;
  let name = '';
  while (value > 0) {
    name = String.fromCharCode(65 + ((value - 1) % 26)) + name;
    value = Math.floor((value - 1) / 26);
  }
  return name;
}
