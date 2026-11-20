import { MAX_SOURCE_LENGTH } from '../core/tasks';
import { parseFormula, parseOutline } from '../core/output';
import type { HostBridge, HostKind, SelectionSnapshot, WritePlan } from '../core/types';

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

async function readWord(): Promise<SelectionSnapshot> {
  return Word.run(async (context) => {
    const range = context.document.getSelection();
    range.load('text');
    await context.sync();
    const text = requireText(range.text);
    context.trackedObjects.add(range);
    await context.sync();
    return {
      text,
      label: 'Word · 已读取的文字选区',
      replace: async (value: string) =>
        Word.run(range, async (ctx) => {
          range.load('text');
          await ctx.sync();
          if (range.text !== text) throw new Error(CHANGED);
          range.insertText(value, Word.InsertLocation.replace);
          await ctx.sync();
        }),
      dispose: async () => {
        context.trackedObjects.remove(range);
        await context.sync();
      },
    };
  });
}

async function prepareWordInsertion(): Promise<WritePlan> {
  return Word.run(async (context) => {
    const range = context.document.getSelection();
    range.load('text');
    await context.sync();
    const before = range.text;
    context.trackedObjects.add(range);
    await context.sync();
    return {
      label: before ? '在当前选区后插入（保留原文）' : '在当前光标处插入',
      before,
      commit: async (text: string) =>
        Word.run(range, async (ctx) => {
          range.load('text');
          await ctx.sync();
          if (range.text !== before) throw new Error(CHANGED);
          range.insertText(text, Word.InsertLocation.end);
          await ctx.sync();
        }),
      dispose: async () => {
        context.trackedObjects.remove(range);
        await context.sync();
      },
    };
  });
}

async function readExcel(): Promise<SelectionSnapshot> {
  return Excel.run(async (context) => {
    const range = context.workbook.getSelectedRange();
    range.load(['address', 'rowCount', 'columnCount']);
    await context.sync();
    if (range.rowCount * range.columnCount > MAX_CELLS)
      throw new Error('一次最多读取 2,000 个单元格，请缩小选区。');
    range.load(['text', 'formulas']);
    await context.sync();
    const values = range.text;
    if (!values.some((row) => row.some((cell) => cell.trim())))
      throw new Error('选区为空，请选中包含数据或公式的单元格。');
    const formulas = range.formulas as unknown[][];
    const hasFormulas = formulas.some((row) =>
      row.some((cell) => typeof cell === 'string' && cell.startsWith('=')),
    );
    // JSON retains tabs/newlines in cell values and avoids ambiguous TSV boundaries.
    const text = JSON.stringify(
      { address: range.address, values, ...(hasFormulas ? { formulas } : {}) },
      null,
      2,
    );
    return { text: requireText(text), label: range.address };
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
      before: String(range.formulas[0][0] || range.text[0][0] || '（空单元格）'),
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
    const count = context.presentation.slides.getCount();
    await context.sync();
    const start = count.value;
    for (let i = 0; i < slides.length; i += 1) context.presentation.slides.add();
    await context.sync();
    const created = slides.map((_, i) => context.presentation.slides.getItemAt(start + i));
    for (const slide of created) slide.shapes.load('items');
    await context.sync();
    for (const [i, content] of slides.entries()) {
      const slide = created[i];
      // Default layouts can contain title/body placeholders; clear only the new slides.
      for (const shape of slide.shapes.items) shape.delete();
      const title = slide.shapes.addTextBox(content.title, {
        left: 40,
        top: 32,
        width: 620,
        height: 72,
      });
      title.textFrame.textRange.font.size = 28;
      title.textFrame.textRange.font.bold = true;
      title.textFrame.autoSizeSetting = 'AutoSizeTextToFitShape';
      // Fit both common 720×405 widescreen and 720×540 standard slide dimensions.
      const body = slide.shapes.addTextBox(content.bullets.map((b) => `• ${b}`).join('\n\n'), {
        left: 40,
        top: 118,
        width: 620,
        height: 245,
      });
      body.textFrame.textRange.font.size = 16;
      body.textFrame.autoSizeSetting = 'AutoSizeTextToFitShape';
    }
    await context.sync();
  });
}

export async function createOfficeBridge(host: HostKind): Promise<HostBridge> {
  const nativeWord = host === 'word' && supports('WordApi');
  const nativeExcel = host === 'excel' && supports('ExcelApi');
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
    host,
    mode: 'office',
    capabilities: {
      read: nativeWord || nativeExcel || commonText,
      replace: canReplace,
      insert: nativeWord,
      formula: nativeExcel,
      slides: host === 'powerpoint' && supports('PowerPointApi', '1.4'),
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
    prepareWrite: async (kind, source) => {
      if (kind === 'replace' && source?.replace && canReplace)
        return {
          label: `替换 ${source.label}`,
          before: source.text,
          commit: source.replace,
        };
      if (kind === 'insert' && nativeWord) return prepareWordInsertion();
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
