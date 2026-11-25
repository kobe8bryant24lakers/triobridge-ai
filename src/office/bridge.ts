import { MAX_SOURCE_LENGTH } from '../core/tasks';
import { parseFormula, parseOutline } from '../core/output';
import { slideFontSize } from '../core/slide-layout';
import type { HostBridge, HostKind, SelectionSnapshot, WritePlan } from '../core/types';
import { wordSelectionWarning } from './word-safety';
import { wordMarkdownFormat } from '../../offline/app/word-markdown.mjs';
import {
  columnName,
  NUMBER_FORMATS,
  parseReport,
  parseTable,
  reportToHtml,
  tableMatrix,
} from '../core/structured-output';

const CHANGED = '目标内容或选区已经变化。请重新读取选区，再预览并确认写入。';
const MAX_CELLS = 2000;

function requireText(text: string): string {
  if (!text.trim()) throw new Error('没有读取到内容。请先在文档中选中文字或单元格。');
  if (text.length > MAX_SOURCE_LENGTH) throw new Error('选区内容超过 30,000 字符，请缩小选区。');
  return text;
}

export function supports(set: string, version = '1.1'): boolean {
  return Office.context.requirements.isSetSupported(set, version);
}

function selectedText(): Promise<string> {
  return new Promise((resolve, reject) =>
    Office.context.document.getSelectedDataAsync(Office.CoercionType.Text, (result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) resolve(String(result.value ?? ''));
      else reject(new Error('无法读取选区。请进入文本编辑状态，选中文字后重试。'));
    }),
  );
}

function setSelectedText(text: string): Promise<void> {
  return new Promise((resolve, reject) =>
    Office.context.document.setSelectedDataAsync(
      text,
      { coercionType: Office.CoercionType.Text },
      (result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) resolve();
        else reject(new Error('Office 无法写入当前选区。请检查文档是否只读，或复制结果手动粘贴。'));
      },
    ),
  );
}

type WordBlockBoundaries = {
  before: OfficeExtension.ClientResult<Word.LocationRelation> | null;
  after: OfficeExtension.ClientResult<Word.LocationRelation> | null;
  paragraphs: Word.ParagraphCollection;
  inserting: boolean;
};

function wordBlockBoundaries(range: Word.Range, inserting: boolean): WordBlockBoundaries {
  const precise = supports('WordApi', '1.3');
  const target = inserting && precise ? range.getRange('End') : range;
  const paragraphs = target.paragraphs;
  // Read only style names, never the neighboring paragraphs' text. Older Word
  // lacks exact boundary comparison but can still preserve both outside styles.
  paragraphs.load('items/style');
  return {
    before: precise
      ? target.getRange('Start').compareLocationWith(paragraphs.getFirst().getRange('Start'))
      : null,
    after: precise
      ? target.getRange('End').compareLocationWith(paragraphs.getLast().getRange('End'))
      : null,
    paragraphs,
    inserting,
  };
}

function wordBoundaryStyle(name: unknown): string {
  if (typeof name !== 'string' || !name.trim())
    throw new Error('无法确认原段落样式，请重新选择写入位置。');
  // Style names come from the document. Escape the CSS string and HTML raw-text
  // delimiters so a custom style cannot end this stylesheet or inject markup.
  return name.replace(
    /[\\"<>&\u0000-\u001f\u007f]/g,
    (char) => `\\${char.charCodeAt(0).toString(16)} `,
  );
}

function wordInsertionHtml(
  formatted: ReturnType<typeof wordMarkdownFormat>,
  boundaries: WordBlockBoundaries | null,
): string | null {
  if (!formatted) return null;
  if (!formatted.block) return formatted.html;
  if (!boundaries) throw new Error('无法确认段落边界，请重新选择写入位置。');
  const before = !boundaries.before || ['After', 'AdjacentAfter'].includes(boundaries.before.value);
  const after = !boundaries.after || ['Before', 'AdjacentBefore'].includes(boundaries.after.value);
  if (!before && !after) return formatted.html;
  const items = boundaries.paragraphs.items;
  const last = items[items.length - 1];
  const first = boundaries.inserting ? last : items[0];
  // Word drops empty <p> boundaries. A line break keeps them present, and an
  // explicit original style prevents the imported boundary restyling neighbors.
  // Keep both boundaries in the same write; never split or delete text first.
  const leading = before
    ? `p.trioWordBefore{mso-style-name:"${wordBoundaryStyle(first?.style)}";}`
    : '';
  const trailing = after
    ? `p.trioWordAfter{mso-style-name:"${wordBoundaryStyle(last?.style)}";}`
    : '';
  return `<style>${leading}${trailing}</style>${before ? '<p class="trioWordBefore"><br></p>' : ''}${formatted.html}${after ? '<p class="trioWordAfter"><br></p>' : ''}`;
}

async function readWord(): Promise<SelectionSnapshot> {
  return Word.run(async (context) => {
    const range = context.document.getSelection();
    range.load('text');
    const ooxml = range.getOoxml();
    await context.sync();
    const text = requireText(range.text);
    const warning = wordSelectionWarning(ooxml.value);
    if (warning) return { text, label: 'Word · 已读取的文字选区', warning };
    context.trackedObjects.add(range);
    await context.sync();
    return {
      text,
      label: 'Word · 已读取的文字选区',
      replace: async (value: string) =>
        Word.run(range, async (ctx) => {
          const formatted = wordMarkdownFormat(value);
          const boundaries = formatted?.block ? wordBlockBoundaries(range, false) : null;
          range.load('text');
          const currentOoxml = range.getOoxml();
          await ctx.sync();
          if (range.text !== text) throw new Error(CHANGED);
          const structureWarning = wordSelectionWarning(currentOoxml.value);
          if (structureWarning) throw new Error(structureWarning);
          const html = wordInsertionHtml(formatted, boundaries);
          if (html) range.insertHtml(html, Word.InsertLocation.replace);
          else range.insertText(value, Word.InsertLocation.replace);
          await ctx.sync();
        }),
      dispose: async () => {
        context.trackedObjects.remove(range);
        await context.sync();
      },
    };
  });
}

async function prepareWordInsertion(structured = false, markdownBlock = false): Promise<WritePlan> {
  return Word.run(async (context) => {
    const range = context.document.getSelection();
    range.load('text');
    await context.sync();
    const before = range.text;
    context.trackedObjects.add(range);
    await context.sync();
    return {
      label: structured
        ? '在已确认光标或选区后新起段落插入报告（保留原文）'
        : markdownBlock
          ? '在已确认光标或选区后新起段落插入排版内容（保留原文）'
          : before
            ? '在已确认选区末尾插入（保留原文）'
            : '在已确认光标处插入',
      before,
      commit: async (text: string) =>
        Word.run(range, async (ctx) => {
          const formatted = structured ? null : wordMarkdownFormat(text);
          const boundaries = formatted?.block ? wordBlockBoundaries(range, true) : null;
          const reportHtml = structured ? reportToHtml(parseReport(text)) : null;
          range.load('text');
          await ctx.sync();
          if (range.text !== before) throw new Error(CHANGED);
          const html = reportHtml || wordInsertionHtml(formatted, boundaries);
          // A heading inserted directly at a paragraph's end can merge into
          // that paragraph and restyle the user's existing text in Word.
          if (structured && html)
            range
              .insertParagraph('', Word.InsertLocation.after)
              .insertHtml(html, Word.InsertLocation.replace);
          // End is the tracked range's exact end, including a cursor in the
          // middle of a paragraph. Do not move the block after its parent paragraph.
          else if (html) range.insertHtml(html, Word.InsertLocation.end);
          else range.insertText(text, Word.InsertLocation.end);
          await ctx.sync();
        }),
      dispose: async () => {
        context.trackedObjects.remove(range);
        await context.sync();
      },
    };
  });
}

async function prepareExcelTable(raw: string): Promise<WritePlan> {
  const table = parseTable(raw);
  const signature = JSON.stringify(table);
  const address = `A1:${columnName(table.headers.length - 1)}${table.rows.length + 1}`;
  const sheetName = await Excel.run(async (context) => {
    const sheets = context.workbook.worksheets;
    sheets.load('items/name');
    await context.sync();
    const names = new Set(sheets.items.map((sheet) => sheet.name.toLocaleLowerCase()));
    let name = table.name;
    for (let i = 2; names.has(name.toLocaleLowerCase()); i += 1) {
      const suffix = ` (${i})`;
      name = `${table.name.slice(0, 31 - suffix.length)}${suffix}`;
    }
    return name;
  });
  return {
    label: `新建工作表「${sheetName}」 · ${address}`,
    before: `新工作表，${table.rows.length + 1} 行 × ${table.headers.length} 列（含表头）。现有工作表不会被覆盖。`,
    warning: '公式将在 Excel 中计算。创建失败时可能留下新工作表，请按错误提示检查后再试。',
    commit: async (value) => {
      if (JSON.stringify(parseTable(value)) !== signature)
        throw new Error('数据表与确认时的内容不同，请返回检查后重新确认。');
      let creationStarted = false;
      let created = false;
      try {
        return await Excel.run(async (context) => {
          const sheets = context.workbook.worksheets;
          sheets.load('items/name');
          await context.sync();
          if (
            sheets.items.some(
              (sheet) => sheet.name.toLocaleLowerCase() === sheetName.toLocaleLowerCase(),
            )
          )
            throw new Error('目标工作表名称已被使用，请返回检查后重新创建。');
          creationStarted = true;
          const sheet = sheets.add(sheetName);
          await context.sync();
          created = true;
          const range = sheet.getRange(address);
          range.formulas = tableMatrix(table);
          range.numberFormat = [
            table.headers.map(() => '@'),
            ...table.rows.map(() => table.formats.map((format) => NUMBER_FORMATS[format])),
          ];
          range.format.wrapText = true;
          range.format.columnWidth = 110;
          const nativeTable = sheet.tables.add(address, true);
          nativeTable.style = 'TableStyleMedium2';
          if (supports('ExcelApi', '1.2')) range.format.autofitRows();
          await context.sync();
          range.load('text');
          const application = context.workbook.application;
          const canReadCalculationState = supports('ExcelApi', '1.9');
          application.load(
            canReadCalculationState ? ['calculationMode', 'calculationState'] : ['calculationMode'],
          );
          sheet.activate();
          await context.sync();
          if (table.rows.some((row) => row.some((cell) => cell && typeof cell === 'object'))) {
            if (application.calculationMode === 'Manual')
              return {
                warning: `工作表「${sheetName}」已创建。Excel 当前为手动计算模式，公式结果尚未验证；请使用 Excel 的“立即计算”后检查结果。`,
              };
            if (canReadCalculationState && application.calculationState !== 'Done')
              return {
                warning: `工作表「${sheetName}」已创建，Excel 仍有待完成的计算；请等待计算结束后检查公式结果。`,
              };
          }
          const errors: string[] = [];
          table.rows.forEach((row, r) =>
            row.forEach((cell, c) => {
              if (
                cell &&
                typeof cell === 'object' &&
                /^#(?:REF!|DIV\/0!|VALUE!|NAME\?|N\/A|NUM!|NULL!|SPILL!|CALC!)/.test(
                  range.text[r + 1][c],
                )
              )
                errors.push(`${columnName(c)}${r + 2} ${range.text[r + 1][c]}`);
            }),
          );
          return errors.length
            ? {
                warning: `工作表「${sheetName}」已创建，但发现 ${errors.length} 个公式错误：${errors.slice(0, 8).join('、')}。请在 Excel 中检查引用和输入数据。`,
              }
            : undefined;
        });
      } catch (error) {
        if (creationStarted)
          throw new Error(
            `工作表「${sheetName}」${created ? '已创建，但写入或检查未完成' : '可能已创建，但未能确认结果'}。请检查该表，删除未完成的表后再试。${officeError(error)}`,
          );
        throw error;
      }
    },
  };
}

async function readExcel(): Promise<SelectionSnapshot> {
  return Excel.run(async (context) => {
    const range = context.workbook.getSelectedRange();
    range.load(['address', 'rowCount', 'columnCount']);
    await context.sync();
    if (range.rowCount * range.columnCount > MAX_CELLS)
      throw new Error('一次最多读取 2,000 个单元格，请缩小选区。');
    range.load(['text', 'values', 'formulas']);
    await context.sync();
    const displayValues = range.text;
    const values = range.values as unknown[][];
    const formulas = range.formulas as unknown[][];
    const hasFormulas = formulas.some((row) =>
      row.some((cell) => typeof cell === 'string' && cell.startsWith('=')),
    );
    if (!hasFormulas && !values.some((row) => row.some((cell) => cell !== '' && cell != null)))
      throw new Error('选区为空，请选中包含数据或公式的单元格。');
    // JSON retains tabs/newlines in cell values and avoids ambiguous TSV boundaries.
    // Compact serialization leaves more of the 30k budget available for actual cells.
    const text = JSON.stringify({
      address: range.address,
      values,
      displayValues,
      ...(hasFormulas ? { formulas } : {}),
    });
    return {
      text: requireText(text),
      label: `${range.address} · ${range.rowCount} 行 × ${range.columnCount} 列`,
    };
  });
}

async function prepareExcelFormula(): Promise<WritePlan> {
  return Excel.run(async (context) => {
    const range = context.workbook.getSelectedRange();
    range.load(['address', 'rowCount', 'columnCount']);
    await context.sync();
    if (range.rowCount !== 1 || range.columnCount !== 1)
      throw new Error('请在 Excel 中选中一个目标单元格，再点击写入公式。');
    range.load(['formulas', 'values', 'text']);
    range.worksheet.load(['id', 'name']);
    await context.sync();
    const worksheetId = range.worksheet.id;
    const address = range.address;
    const cellAddress = address.slice(address.lastIndexOf('!') + 1);
    const signature = JSON.stringify([range.formulas, range.values]);
    return {
      label: `写入 ${address}`,
      before: String(range.formulas[0][0] ?? range.text[0][0]) || '（空单元格）',
      commit: async (raw: string) => {
        const formula = parseFormula(raw);
        await Excel.run(async (ctx) => {
          const target = ctx.workbook.worksheets.getItem(worksheetId).getRange(cellAddress);
          target.load(['formulas', 'values']);
          await ctx.sync();
          if (JSON.stringify([target.formulas, target.values]) !== signature)
            throw new Error(CHANGED);
          target.formulas = [[formula]];
          await ctx.sync();
        });
      },
    };
  });
}

async function addSlides(raw: string): Promise<void> {
  const slides = parseOutline(raw);
  // PowerPointApi 1.4 is checked before reaching this function.
  await PowerPoint.run(async (context) => {
    const collection = context.presentation.slides;
    collection.load('items/id');
    const pageSetup = supports('PowerPointApi', '1.10')
      ? context.presentation.pageSetup
      : undefined;
    pageSetup?.load(['slideWidth', 'slideHeight']);
    await context.sync();
    const previousIds = collection.items.map((slide) => slide.id);
    const width = pageSetup?.slideWidth ?? 720;
    const height = pageSetup?.slideHeight ?? 405;
    if (!Number.isFinite(width) || !Number.isFinite(height) || width < 200 || height < 200)
      throw new Error('无法确认幻灯片尺寸，请复制大纲后手动排版。');
    for (let i = 0; i < slides.length; i += 1) collection.add();
    await context.sync();
    collection.load('items/id');
    await context.sync();
    const currentIds = collection.items.map((slide) => slide.id);
    if (
      currentIds.length !== previousIds.length + slides.length ||
      previousIds.some((id, i) => currentIds[i] !== id)
    )
      throw new Error('建页期间演示文稿页序发生变化，已停止填充。请检查新增的空白页并撤销后重试。');
    // Bind by ID after creation so later reordering cannot redirect writes to old slides.
    const created = currentIds.slice(previousIds.length).map((id) => collection.getItem(id));
    for (const slide of created) slide.shapes.load('items');
    await context.sync();
    for (const [i, content] of slides.entries()) {
      const slide = created[i];
      // Default layouts can contain title/body placeholders; clear only the new slides.
      for (const shape of slide.shapes.items) shape.delete();
      const title = slide.shapes.addTextBox(content.title, {
        left: width * 0.055,
        top: height * 0.075,
        width: width * 0.89,
        height: height * 0.18,
      });
      title.textFrame.textRange.font.size = slideFontSize(
        content.title,
        width * 0.89,
        height * 0.18,
        28,
        18,
      );
      title.textFrame.textRange.font.bold = true;
      title.textFrame.autoSizeSetting = 'AutoSizeTextToFitShape';
      // Older hosts use a 720×405 safe area; 1.10+ follows the actual page size.
      const bodyText = content.bullets.map((b) => `• ${b}`).join('\n\n');
      const body = slide.shapes.addTextBox(bodyText, {
        left: width * 0.055,
        top: height * 0.29,
        width: width * 0.89,
        height: height * 0.61,
      });
      body.textFrame.textRange.font.size = slideFontSize(
        bodyText,
        width * 0.89,
        height * 0.61,
        16,
        12,
      );
      body.textFrame.autoSizeSetting = 'AutoSizeTextToFitShape';
    }
    await context.sync();
  });
}

export async function createOfficeBridge(host: HostKind): Promise<HostBridge> {
  const nativeWord = host === 'word' && supports('WordApi');
  const nativeExcel = host === 'excel' && supports('ExcelApi');
  const nativeTable = nativeExcel && supports('ExcelApi', '1.2');
  const commonText = supports('TextCoercion');
  let selectionRevision = 0;
  let guarded = false;
  const onSelectionChanged = () => {
    selectionRevision += 1;
  };
  if (host === 'powerpoint' && commonText) {
    guarded = await new Promise<boolean>((resolve) => {
      const timer = setTimeout(() => resolve(false), 2500);
      Office.context.document.addHandlerAsync(
        Office.EventType.DocumentSelectionChanged,
        onSelectionChanged,
        (result) => {
          clearTimeout(timer);
          resolve(result.status === Office.AsyncResultStatus.Succeeded);
        },
      );
    });
  }
  const canReplace = nativeWord || (host === 'powerpoint' && guarded);
  return {
    ...(host === 'powerpoint' ? { presentation: createPresentationAccess() } : {}),
    host,
    mode: 'office',
    capabilities: {
      read: nativeWord || nativeExcel || commonText,
      replace: canReplace,
      insert: nativeWord,
      formula: nativeExcel,
      slides: host === 'powerpoint' && supports('PowerPointApi', '1.4'),
      document: nativeWord,
      table: nativeTable,
    },
    readSelection: async () => {
      if (nativeWord) return readWord();
      if (nativeExcel) return readExcel();
      if (!commonText) throw new Error('当前 Office 无法读取选区。可将内容粘贴到输入框。');
      const revision = selectionRevision;
      const text = requireText(await selectedText());
      if (revision !== selectionRevision) throw new Error(CHANGED);
      return {
        text,
        label: `${host === 'powerpoint' ? 'PowerPoint' : 'Office'} · 文本选区`,
        ...(host === 'word'
          ? {
              warning: '当前 Word 版本不支持自动写入排版结果，请更新 Office，或复制结果手动粘贴。',
            }
          : {}),
        ...(canReplace
          ? {
              replace: async (value: string) => {
                const currentText = await selectedText();
                if (selectionRevision !== revision || currentText !== text)
                  throw new Error(CHANGED);
                await setSelectedText(value);
              },
            }
          : {}),
      };
    },
    prepareWrite: async (kind, source, output) => {
      if (kind === 'replace' && source?.replace && canReplace)
        return {
          label: `替换 ${source.label}`,
          before: source.text,
          ...(nativeWord
            ? {
                warning: wordMarkdownFormat(output || '')
                  ? '将按 Markdown 排版替换选区，原选区的字体和段落样式可能改变。'
                  : '将以纯文本替换，选区内部的字体、加粗和段落样式可能改变。需要保留排版时，请返回并选择插入。',
              }
            : {}),
          commit: source.replace,
        };
      if (kind === 'insert' && nativeWord)
        return prepareWordInsertion(false, Boolean(wordMarkdownFormat(output || '')?.block));
      if (kind === 'document' && nativeWord) {
        parseReport(output || '');
        return prepareWordInsertion(true);
      }
      if (kind === 'table' && nativeTable) return prepareExcelTable(output || '');
      if (kind === 'formula' && nativeExcel) return prepareExcelFormula();
      if (kind === 'slides' && host === 'powerpoint' && supports('PowerPointApi', '1.4'))
        return {
          label: '在演示文稿末尾添加幻灯片',
          before: '保留所有现有幻灯片',
          commit: addSlides,
        };
      throw new Error('此版本 Office 不支持该写入操作。请复制结果并手动粘贴。');
    },
    dispose: () => {
      if (guarded)
        Office.context.document.removeHandlerAsync(
          Office.EventType.DocumentSelectionChanged,
          { handler: onSelectionChanged },
          () => {},
        );
    },
  };
}

export function officeError(error: unknown): string {
  if (error instanceof Error && error.name === 'RichApi.Error')
    return 'Office 操作失败，可能是目标已改变或文档只读。请重新选择内容后重试；结果仍可复制。';
  return error instanceof Error ? error.message : '操作未完成，请重试。';
}
import { createPresentationAccess } from './presentation';
