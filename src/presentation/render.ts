import type PptxGenJS from 'pptxgenjs';
import type {
  ChartData,
  ChartKind,
  DeckContent,
  MaterialContext,
  PageSize,
  PresentationTheme,
} from './types';
import { allTables, parseDeck } from './content';
import { chartData } from './charts';
import { preserveChartData } from './chart-workbooks';

export const DEFAULT_THEME: PresentationTheme = {
  font: 'Aptos',
  accent: '#28604A',
  ink: '#25362D',
  background: '#FAFCF9',
};
export const WIDE: PageSize = { width: 960, height: 540 };
export const STANDARD: PageSize = { width: 720, height: 540 };
const color = (s: string) => s.replace(/^#/, '');
const textUnits = (value: string) =>
  [...value].reduce((n, c) => n + (/[^\x00-\x7F]/.test(c) ? 1 : /[MW@%]/.test(c) ? 0.85 : 0.56), 0);
function validateTheme(theme: PresentationTheme, size: PageSize) {
  if (
    !theme.font.trim() ||
    theme.font.length > 80 ||
    ![theme.accent, theme.ink, theme.background].every((s) => /^#[A-F\d]{6}$/i.test(s))
  )
    throw new Error('演示稿主题无效。');
  if (![size.width, size.height].every((n) => Number.isFinite(n) && n >= 200 && n <= 4000))
    throw new Error('页面尺寸不支持。');
}
export function addNativeChart(
  pptx: PptxGenJS,
  slide: PptxGenJS.Slide,
  data: ChartData,
  kind: ChartKind,
  area: { x: number; y: number; w: number; h: number },
  theme: PresentationTheme,
  unit = '',
) {
  const type = kind === 'column' || kind === 'bar' ? pptx.ChartType.bar : pptx.ChartType[kind];
  const numberFormat =
    unit === '%'
      ? '0.0%'
      : data.series.every((s) => s.values.every((v) => v === null || Number.isInteger(v)))
        ? '#,##0'
        : '#,##0.00';
  // The final export reconciles the chart cache and embedded workbook with these values.
  slide.addChart(
    type,
    data.series.map((s) => ({
      name: s.name,
      labels: data.categories,
      values: s.values as number[],
    })),
    {
      ...area,
      catAxisLabelFontFace: theme.font,
      catAxisLabelFontSize: 11,
      valAxisLabelFontFace: theme.font,
      valAxisLabelFontSize: 11,
      valAxisLabelFormatCode: numberFormat,
      chartColors: [color(theme.accent), '6F94A8', 'C59B53', '937A9B', '9AAB75', 'C88176'],
      showLegend: data.series.length > 1 || kind === 'pie' || kind === 'doughnut',
      legendPos: 'b',
      legendFontFace: theme.font,
      legendFontSize: 11,
      showTitle: false,
      showPercent: kind === 'pie' || kind === 'doughnut',
      showLabel: false,
      lineDataSymbol: 'circle',
      lineDataSymbolSize: 5,
      dataLabelFormatCode: numberFormat,
      dataLabelPosition: kind === 'line' ? 't' : undefined,
      showValue: kind === 'pie' || kind === 'doughnut' ? false : data.categories.length <= 8,
      barDir: kind === 'bar' ? 'bar' : 'col',
      barGrouping: 'clustered',
      displayBlanksAs: 'gap',
      catAxisLineColor: 'CFD9D1',
      valAxisLineColor: 'CFD9D1',
      valGridLine: { color: 'E2E8E1', size: 0.5 },
    },
  );
}
export async function renderDeck(
  input: DeckContent,
  context: MaterialContext,
  theme: PresentationTheme = DEFAULT_THEME,
  size: PageSize = WIDE,
): Promise<string> {
  validateTheme(theme, size);
  const deck = parseDeck(JSON.stringify(input), context);
  const { default: Constructor } = await import('pptxgenjs');
  const pptx = new Constructor();
  pptx.defineLayout({ name: 'CUSTOM', width: size.width / 72, height: size.height / 72 });
  pptx.layout = 'CUSTOM';
  pptx.author = 'TrioBridge AI';
  pptx.subject = '根据所选素材生成的可编辑演示稿';
  pptx.title = deck.title;
  pptx.company = '';
  pptx.theme = { headFontFace: theme.font, bodyFontFace: theme.font };
  const w = size.width / 72,
    h = size.height / 72,
    scale = Math.min(w / 13.333, h / 7.5);
  const tables = allTables(context);
  const chartSources: ChartData[] = [];
  for (const [index, page] of deck.slides.entries()) {
    const s = pptx.addSlide();
    s.background = { color: color(theme.background) };
    const tx = (
      value: string,
      x: number,
      y: number,
      width: number,
      height: number,
      fontSize = 22,
      options: PptxGenJS.TextPropsOptions = {},
    ) =>
      s.addText(value, {
        x: x * w,
        y: y * h,
        w: width * w,
        h: height * h,
        fontSize: fontSize * scale,
        fontFace: theme.font,
        lang: 'zh-CN',
        color: color(theme.ink),
        margin: 0,
        breakLine: false,
        valign: 'middle',
        fit: 'shrink',
        objectName:
          y >= 0.88 || fontSize <= 12
            ? 'TrioBridge Footer'
            : fontSize >= 28 && y < 0.5
              ? 'TrioBridge Title'
              : 'TrioBridge Content',
        ...options,
      });
    const rect = (x: number, y: number, width: number, height: number, fill: string) =>
      s.addShape(pptx.ShapeType.rect, {
        x: x * w,
        y: y * h,
        w: width * w,
        h: height * h,
        fill: { color: fill },
        line: { transparency: 100 },
      });
    const bullets = (
      values: string[],
      x: number,
      y: number,
      width: number,
      height: number,
      font = 22,
    ) => {
      const gap = height / Math.max(1, values.length);
      values.forEach((value, i) => {
        s.addShape(pptx.ShapeType.ellipse, {
          x: x * w,
          y: (y + gap * i + 0.015) * h,
          w: 0.065 * scale,
          h: 0.065 * scale,
          fill: { color: color(theme.accent) },
          line: { transparency: 100 },
        });
        tx(value, x + 0.02, y + gap * i, width - 0.02, gap - 0.02, font, { valign: 'top' });
      });
    };
    if (page.layout === 'cover' || page.layout === 'closing') {
      rect(0, 0, 0.36, 1, color(theme.accent));
      tx(
        page.layout === 'cover' ? String(index + 1).padStart(2, '0') : 'END',
        0.055,
        0.68,
        0.25,
        0.16,
        68,
        { color: 'FFFFFF', bold: true },
      );
      tx(page.title, 0.42, 0.24, 0.52, 0.25, 40, { bold: true });
      tx(page.subtitle || '', 0.42, 0.55, 0.5, 0.2, 22, { valign: 'top' });
    } else {
      tx(page.title, 0.055, 0.06, 0.89, 0.115, 30, { bold: true });
      s.addShape(pptx.ShapeType.line, {
        x: 0.055 * w,
        y: 0.2 * h,
        w: 0.89 * w,
        h: 0,
        line: { color: color(theme.accent), width: 1.2 },
      });
      switch (page.layout) {
        case 'bullets':
          bullets(page.bullets!, 0.065, 0.27, 0.86, 0.56, 25);
          break;
        case 'twoColumn':
          page.columns!.forEach((col, i) => {
            const x = 0.055 + i * 0.46;
            tx(col.heading, x, 0.27, 0.42, 0.09, 24, { color: color(theme.accent), bold: true });
            bullets(col.bullets, x, 0.4, 0.41, 0.43, 21);
          });
          break;
        case 'metrics': {
          const items = page.metrics!,
            n = items.length;
          items.forEach((item, i) => {
            const x = 0.055 + i * (0.89 / n),
              cw = 0.89 / n - 0.025;
            tx(
              item.value,
              x,
              0.3,
              cw,
              0.22,
              Math.min(n > 3 ? 40 : 50, (cw * w * 72) / scale / Math.max(1, textUnits(item.value))),
              { color: color(theme.accent), bold: true, wrap: false },
            );
            tx(item.label, x, 0.55, cw, 0.1, 23, { bold: true });
            tx(item.detail, x, 0.69, cw, 0.16, 17, { valign: 'top' });
          });
          break;
        }
        case 'process': {
          const items = page.steps!,
            n = items.length;
          items.forEach((item, i) => {
            const x = 0.055 + i * (0.89 / n),
              cw = 0.89 / n - 0.02;
            s.addShape(pptx.ShapeType.ellipse, {
              x: x * w,
              y: 0.29 * h,
              w: 0.44 * scale,
              h: 0.44 * scale,
              fill: { color: color(theme.accent) },
              line: { transparency: 100 },
            });
            tx(String(i + 1), x, 0.29, (0.44 * scale) / w, (0.44 * scale) / h, 18, {
              align: 'center',
              color: 'FFFFFF',
              bold: true,
            });
            tx(item.title, x, 0.43, cw, 0.12, 22, { bold: true });
            tx(item.detail, x, 0.6, cw, 0.22, 17, { valign: 'top' });
          });
          break;
        }
        case 'image': {
          const image = context.materials.find((m) => m.id === page.imageId)!.image!;
          const box = {
            x: 0.055 * w,
            y: 0.25 * h,
            w: (page.bullets?.length ? 0.52 : 0.89) * w,
            h: 0.59 * h,
          };
          const ratio = Math.min(box.w / image.width, box.h / image.height);
          const iw = image.width * ratio,
            ih = image.height * ratio;
          s.addImage({
            data: image.data,
            x: box.x + (box.w - iw) / 2,
            y: box.y + (box.h - ih) / 2,
            w: iw,
            h: ih,
            altText: context.materials.find((m) => m.id === page.imageId)!.name,
          });
          if (page.bullets?.length) bullets(page.bullets, 0.62, 0.27, 0.32, 0.56, 21);
          break;
        }
        case 'chart': {
          const c = page.chart!,
            table = tables.find((t) => t.id === c.tableId)!,
            data = chartData(table, c.category, c.series, c.type);
          chartSources.push(data);
          addNativeChart(
            pptx,
            s,
            data,
            c.type,
            { x: 0.065 * w, y: 0.25 * h, w: 0.87 * w, h: 0.59 * h },
            theme,
            c.unit,
          );
          if (c.unit) tx(`单位：${c.unit}`, 0.7, 0.205, 0.24, 0.035, 11, { align: 'right' });
          break;
        }
        case 'table': {
          const table = tables.find((t) => t.id === page.tableId)!;
          const cellWidth = (0.89 * size.width) / table.headers.length - 12;
          const rows = [table.headers, ...table.rows];
          const rowHeights = (font: number) =>
            rows.map(
              (row) =>
                Math.max(
                  ...row.map((v) =>
                    String(v ?? '')
                      .split('\n')
                      .reduce(
                        (n, line) =>
                          n + Math.max(1, Math.ceil((textUnits(line) * font) / cellWidth)),
                        0,
                      ),
                  ),
                ) *
                  font *
                  1.3 +
                12,
            );
          let tableFont = Math.max(12, 17 * scale),
            heights = rowHeights(tableFont);
          while (tableFont > 12 && heights.reduce((a, b) => a + b, 0) > 0.57 * size.height) {
            tableFont = Math.max(12, tableFont - 1);
            heights = rowHeights(tableFont);
          }
          if (heights.reduce((a, b) => a + b, 0) > 0.57 * size.height)
            throw new Error(
              `第 ${index + 1} 页表格文字过密，无法保持 12 磅以上字号。请精简素材表格或改为文字总结后重新生成。`,
            );
          s.addTable(
            [
              table.headers.map((t) => ({
                text: t,
                options: { bold: true, color: 'FFFFFF', fill: { color: color(theme.accent) } },
              })),
              ...table.rows.map((row) => row.map((v) => ({ text: String(v ?? '') }))),
            ],
            {
              x: 0.055 * w,
              y: 0.26 * h,
              w: 0.89 * w,
              h: 0.57 * h,
              colW: Array(table.headers.length).fill((0.89 * w) / table.headers.length),
              rowH: heights.map((value) => value / 72),
              fontFace: theme.font,
              fontSize: tableFont,
              color: color(theme.ink),
              fill: { color: 'FFFFFF' },
              border: { type: 'solid', color: 'DDE5DE', pt: 0.5 },
              margin: 5,
              autoPage: false,
              valign: 'middle',
            },
          );
          break;
        }
      }
    }
    const sourceNames = page.sources
      .map((id) => context.materials.find((m) => m.id === id)?.name)
      .filter(Boolean);
    tx(
      sourceNames.length ? `来源：${sourceNames.join(' · ')}` : '请核对内容与数据',
      page.layout === 'cover' || page.layout === 'closing' ? 0.42 : 0.055,
      0.925,
      page.layout === 'cover' || page.layout === 'closing' ? 0.43 : 0.79,
      0.035,
      9,
      { color: '647569' },
    );
    tx(`${index + 1} / ${deck.slides.length}`, 0.86, 0.925, 0.085, 0.035, 10, { align: 'right' });
    s.addNotes(
      `${page.notes}\n\n素材来源：${sourceNames.join('；') || '未引用素材，请核实事实。'}`,
    );
  }
  return preserveChartData(
    (await pptx.write({ outputType: 'base64', compression: true })) as string,
    chartSources,
  );
}
export function downloadBlob(blob: Blob, filename: string) {
  const url = URL.createObjectURL(blob),
    a = document.createElement('a');
  a.href = url;
  a.download = filename.replace(/[\\/:*?"<>|]/g, '-').slice(0, 160);
  a.click();
  setTimeout(() => URL.revokeObjectURL(url), 30_000);
}
export function downloadPresentation(base64: string, name: string) {
  const bytes = Uint8Array.from(atob(base64), (c) => c.charCodeAt(0));
  downloadBlob(
    new Blob([bytes], {
      type: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
    }),
    `${name}.pptx`,
  );
}
