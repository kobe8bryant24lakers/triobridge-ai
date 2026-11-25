ObjC.import('Foundation');

function fail(message) {
  throw new Error(message);
}
function optional(read, fallback) {
  try {
    const v = read();
    return v == null ? fallback : v;
  } catch (_) {
    return fallback;
  }
}
function same(before, now) {
  if (JSON.stringify(before) !== JSON.stringify(now))
    fail('文档、选区或内容已发生变化。请重新读取并确认写入位置。');
}
function limited(text) {
  if (text.length > 48000) fail('选区文字过长，请缩小选区再读取。');
  return text;
}
function documentKey(doc) {
  return [String(doc.name()), String(optional(() => doc.fullName(), ''))];
}
function rangeValue(v) {
  return Array.isArray(v)
    ? v.map((row) => (Array.isArray(row) ? row.join('\t') : String(row))).join('\n')
    : String(v == null ? '' : v);
}
function col(n) {
  let s = '';
  for (; n; n = Math.floor((n - 1) / 26)) s = String.fromCharCode(65 + ((n - 1) % 26)) + s;
  return s;
}

function word(a, q) {
  const doc = a.activeDocument,
    sel = a.selection;
  function current() {
    const text = limited(String(sel.content() || ''));
    const guard = {
      doc: documentKey(doc),
      start: sel.selectionStart(),
      end: sel.selectionEnd(),
      story: String(sel.storyType()),
      text,
    };
    // Fail closed if a version of Word cannot inspect a potentially structured range.
    const protectedRange =
      [
        'tables',
        'fields',
        'inlineShapes',
        'mathObjects',
        'formFields',
        'footnotes',
        'endnotes',
        'hyperlinkObjects',
      ].some((k) => optional(() => sel[k].length, 1) > 0) || /[\u0001\u0007]/.test(text);
    const writable = text.trim().length > 0 && !protectedRange;
    return { guard, text, label: 'Word · 当前选区', writable };
  }
  const target = current();
  if (q.op === 'selection') {
    if (!target.text.trim()) fail('请在 Word 中先选中需要处理的文字。');
    if (!target.writable)
      target.warning = '选区包含表格、域或其他文档对象，可以读取并在选区后插入结果。';
    return target;
  }
  if (q.op === 'target')
    return Object.assign(target, { label: 'Word · 在当前选区后插入', writable: true });
  same(q.guard, target.guard);
  if (q.kind === 'replace') {
    if (!target.writable) fail('这段选区包含特殊对象，请使用插入。');
    sel.content = String(q.value);
  } else if (q.kind === 'document' || q.kind === 'insert') {
    const range = a.collapseRange(sel.textObject, { direction: 'collapse end' });
    if (range.startOfContent() !== target.guard.end || range.endOfContent() !== target.guard.end)
      fail('未能确认选区末尾的插入位置，请重新选择。');
    if (q.kind === 'document')
      a.insertFile({ at: range, fileName: q.file, confirmConversions: false, link: false });
    else range.content = String(q.value);
  } else fail('未知 Word 操作。');
}
function excel(a, q) {
  const book = a.activeWorkbook,
    sheet = a.activeSheet;
  const bookKey = documentKey(book);
  if (q.kind === 'table') {
    const guard = { doc: bookKey, sheets: book.worksheets.name() };
    if (q.op === 'target')
      return {
        guard,
        text: '新建工作表，保留已有数据',
        label: 'Excel · 新建工作表',
        writable: true,
      };
    same(q.guard, guard);
    const data = q.value,
      m = data.matrix;
    if (!m.length || !m[0].length || m.length > 1001 || m[0].length > 40)
      fail('数据表尺寸超出限制。');
    let base =
        String(data.name)
          .replace(/[:\\/?*\[\]]/g, ' ')
          .trim()
          .slice(0, 24) || 'AI 数据',
      name = base,
      i = 2;
    const names = guard.sheets.map((x) => x.toLowerCase());
    while (names.includes(name.toLowerCase())) name = base + ' ' + i++;
    const ws = a.Worksheet({ name });
    book.worksheets.push(ws);
    const added = book.worksheets.byName(name);
    try {
      const r = added.ranges.byName('A1:' + col(m[0].length) + m.length);
      r.formula = m;
      added.ranges.byName('A1:' + col(m[0].length) + '1').fontObject.bold = true;
      for (let j = 0; j < data.formats.length && m.length > 1; j++)
        added.ranges.byName(col(j + 1) + '2:' + col(j + 1) + m.length).numberFormat =
          data.formats[j];
    } catch (e) {
      const alerts = a.displayAlerts();
      try {
        a.displayAlerts = false;
        a.delete(added);
      } finally {
        a.displayAlerts = alerts;
      }
      throw e;
    }
    return;
  }
  const r = q.kind === 'formula' ? a.activeCell : a.selection;
  if (Number(r.countLarge()) > 2000) fail('请把选区缩小到 2,000 个单元格以内。');
  const address = String(a.getAddress(r)),
    values = r.value(),
    formula = r.formula();
  const guard = { doc: bookKey, sheet: sheet.name(), address, values, formula };
  const text = limited(
    rangeValue(values) +
      (JSON.stringify(formula) !== JSON.stringify(values) ? '\n公式\n' + rangeValue(formula) : ''),
  );
  if (q.op === 'selection' || q.op === 'target')
    return {
      guard,
      text,
      label: 'Excel · ' + sheet.name() + '!' + address,
      writable: q.kind === 'formula',
    };
  same(q.guard, guard);
  if (q.kind !== 'formula' || r.countLarge() != 1) fail('请选择一个公式目标单元格。');
  if (optional(() => r.mergeCells(), false) || optional(() => r.hasArray(), false))
    fail('请选择普通单元格，不能覆盖合并区域或数组公式。');
  r.formula = String(q.value);
}
function pptSnapshot(a, scope) {
  const doc = a.activePresentation,
    window = a.activeWindow;
  const slideIds = doc.slides.slideID().map(String);
  if (slideIds.length > 500) fail('演示稿超过 500 页，请拆分后再编辑。');
  let ids = slideIds,
    selectedShapeIds = [];
  if (scope === 'selected') {
    const range = window.selection.slideRange;
    ids = optional(() => range.slides.slideID().map(String), []);
    if (!ids.length) ids = [String(window.view.slide.slideID())];
  }
  if (ids.length > 60) fail('每次最多处理 60 页，请在 PowerPoint 中缩小选区。');
  let skipped = 0,
    count = 0;
  const slides = ids.map((id) => {
    const index = slideIds.indexOf(id);
    if (index < 0) fail('无法识别所选页面。');
    const slide = doc.slides[index],
      n = slide.shapes.length;
    count += n;
    if (count > 350) fail('选区对象过多，请减少选中的页面。');
    const selectedNames =
      scope === 'selected' && ids.length === 1
        ? optional(() => window.selection.shapeRange.shapes.name(), [])
        : [];
    const shapes = [];
    for (let k = 0; k < n; k++) {
      const s = slide.shapes[k],
        name = String(s.name()),
        shape = {
          id: String(k + 1),
          name,
          type: String(s.shapeType()),
          left: s.leftPosition(),
          top: s.top(),
          width: s.width(),
          height: s.height(),
        };
      if (selectedNames.includes(name)) selectedShapeIds.push(shape.id);
      const value = optional(() => s.textFrame.textRange.content(), null);
      if (typeof value === 'string') {
        const r = s.textFrame.textRange,
          colors = optional(() => r.font.fontColor(), null);
        shape.text = limited(value);
        shape.role = /title|标题/i.test(name) ? 'title' : 'body';
        shape.font = {
          name: optional(() => r.font.fontName(), null),
          size: optional(() => r.font.fontSize(), null),
          bold: optional(() => r.font.bold(), null),
          color: Array.isArray(colors)
            ? '#' +
              colors
                .slice(0, 3)
                .map((c) =>
                  Math.round(c / 257)
                    .toString(16)
                    .padStart(2, '0'),
                )
                .join('')
                .toUpperCase()
            : null,
        };
        shape.alignment = optional(
          () =>
            String(r.paragraphFormat.alignment())
              .replace(/^paragraph align /, '')
              .replace(/^./, (c) => c.toUpperCase()),
          null,
        );
      } else skipped++;
      shapes.push(shape);
    }
    return {
      id,
      index,
      title:
        (shapes.find((s) => s.role === 'title' && s.text) || shapes.find((s) => s.text) || {})
          .text || '第 ' + (index + 1) + ' 页',
      shapes,
    };
  });
  return {
    guard: documentKey(doc),
    snapshot: { scope, slides, slideIds, selectedShapeIds, skipped, size: null },
  };
}
function powerpoint(a, q) {
  if (q.op === 'openDeck') {
    a.open(Path(q.file));
    a.activate();
    return;
  }
  if (q.op === 'presentation') return pptSnapshot(a, q.scope);
  const doc = a.activePresentation;
  if (q.op === 'selection') {
    const data = pptSnapshot(a, 'selected');
    const text = data.snapshot.slides
      .map((s) =>
        s.shapes
          .map((sh) => sh.text)
          .filter(Boolean)
          .join('\n'),
      )
      .join('\n');
    if (!text) fail('选中的页面没有可读取的文字。');
    return {
      guard: data.guard,
      label: 'PowerPoint · 选中页面',
      text: limited(text),
      writable: false,
    };
  }
  if (q.op === 'target' || q.op === 'write') {
    const guard = { doc: documentKey(doc), ids: doc.slides.slideID().map(String) };
    if (q.op === 'target')
      return { guard, text: '保留现有页面', label: 'PowerPoint · 末尾添加页面', writable: true };
    same(q.guard, guard);
    const added = [];
    try {
      for (const spec of q.value) {
        const slide = a.Slide({ layout: 'slide layout text slide' });
        doc.slides.push(slide);
        const last = doc.slides[doc.slides.length - 1];
        added.push(last.slideID());
        last.shapes[0].textFrame.textRange.content = spec.title;
        last.shapes[1].textFrame.textRange.content = spec.bullets.join('\r');
      }
    } catch (e) {
      for (const id of added.reverse()) {
        const i = doc.slides.slideID().indexOf(id);
        if (i >= 0) a.delete(doc.slides[i]);
      }
      throw e;
    }
    return;
  }
  if (q.op !== 'edit') fail('该操作不适用于此 PowerPoint 入口。');
  const now = pptSnapshot(a, q.snapshot.scope);
  same(q.guard, now.guard);
  same(q.snapshot, now.snapshot);
  if (q.proposal.slideOrder) fail('请在 PowerPoint 中调整页面顺序。');
  const rollback = [];
  function set(obj, key, v) {
    const old = obj[key]();
    obj[key] = v;
    rollback.push(() => {
      obj[key] = old;
    });
  }
  try {
    for (const p of q.proposal.changes) {
      const slide = now.snapshot.slides.find((s) => s.id === p.slideId),
        shape = slide && slide.shapes.find((s) => s.id === p.shapeId);
      if (!shape) fail('页面对象已变化。');
      const obj = doc.slides[slide.index].shapes[Number(shape.id) - 1],
        r = obj.textFrame.textRange;
      if (p.text !== undefined) set(r, 'content', p.text);
      if (p.fontName !== undefined) set(r.font, 'fontName', p.fontName);
      if (p.fontSize !== undefined) set(r.font, 'fontSize', p.fontSize);
      if (p.bold !== undefined) set(r.font, 'bold', p.bold);
      if (p.fontColor !== undefined)
        set(
          r.font,
          'fontColor',
          [1, 3, 5].map((i) => parseInt(p.fontColor.slice(i, i + 2), 16) * 257),
        );
      if (p.alignment !== undefined)
        set(r.paragraphFormat, 'alignment', 'paragraph align ' + p.alignment.toLowerCase());
      for (const key of ['left', 'top', 'width', 'height'])
        if (p[key] !== undefined) set(obj, key === 'left' ? 'leftPosition' : key, p[key]);
    }
  } catch (e) {
    let failed = false;
    for (const undo of rollback.reverse()) {
      try {
        undo();
      } catch (_) {
        failed = true;
      }
    }
    if (failed) fail('部分改动需要在 PowerPoint 中撤销，请检查目标页面。');
    throw e;
  }
}
function run() {
  try {
    const data = $.NSFileHandle.fileHandleWithStandardInput.readDataToEndOfFile;
    const q = JSON.parse(
      ObjC.unwrap($.NSString.alloc.initWithDataEncoding(data, $.NSUTF8StringEncoding)),
    );
    const names = {
      word: 'Microsoft Word',
      excel: 'Microsoft Excel',
      powerpoint: 'Microsoft PowerPoint',
    };
    if (!Object.prototype.hasOwnProperty.call(names, q.host)) fail('未知的 Office 应用。');
    const app = Application(names[q.host]);
    if (!app.running() && q.op !== 'openDeck') fail('请先在 ' + names[q.host] + ' 中打开文档。');
    const value = { word, excel, powerpoint }[q.host](app, q);
    return JSON.stringify({ ok: true, value: value == null ? null : value });
  } catch (e) {
    const msg = String(e.message || '');
    const hint = /(-1743|not authorized|not permitted)/i.test(msg)
      ? '请在“系统设置 → 隐私与安全性 → 自动化”允许工作台控制目标 Office 应用。'
      : /[\u3400-\u9fff]/.test(msg)
        ? msg
        : 'Office 未完成操作。请检查当前文档、选区及编辑权限。';
    return JSON.stringify({ ok: false, message: hint });
  }
}
