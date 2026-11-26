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
function wordDocumentKeys(a) {
  // Word can return null entries for documents.name() after Save As, even
  // though the corresponding indexed document exposes both name and fullName.
  const keys = [];
  for (let i = 0; i < a.documents.length; i++) keys.push(documentKey(a.documents[i]));
  return keys;
}
function wordDocument(a, key) {
  const keys = wordDocumentKeys(a);
  const matches = keys.filter((entry) => entry[1] === key[1]);
  if (!key[1] || matches.length !== 1)
    fail('无法唯一定位目标 Word 文档，请关闭重复副本后重新读取。');
  same(key, matches[0]);
  // A short byName() key can fail after Save As, and an index moves when Word
  // opens the import. Word accepts the complete HFS filename as its document key.
  // Unlike a whose() predicate, that key also works for activation/write commands.
  let doc;
  try {
    doc = a.documents.byName(key[1]);
    documentKey(doc);
  } catch (error) {
    // Word reports /private/var for an import opened through /var; its complete
    // path lookup can then fail. Only a unique, fully verified short name may
    // serve as the alternate key. Never use an ambiguous same-name document.
    if (Number(error.errorNumber) !== -1728) throw error;
    if (keys.filter((entry) => entry[0] === key[0]).length !== 1)
      fail('无法唯一定位目标 Word 文档，请关闭同名副本后重新读取。');
    doc = a.documents.byName(key[0]);
  }
  same(key, documentKey(doc));
  return doc;
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
    const start = sel.selectionStart(),
      end = sel.selectionEnd();
    // Some Word versions return the next character for a collapsed selection.
    const text = start === end ? '' : limited(String(sel.content() || ''));
    const guard = {
      doc: documentKey(doc),
      start,
      end,
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
  if (optional(() => doc.readOnly(), false)) fail('文档为只读，请打开可编辑副本。');
  if (q.file && (q.kind === 'replace' || q.kind === 'insert')) {
    if (q.kind === 'replace' && !target.writable) fail('这段选区包含特殊对象，请使用插入。');
    // activeDocument and selection are deferred JXA specifiers. Bind the original
    // by full identity before opening the private HTML file.
    const original = wordDocument(a, target.guard.doc);
    same(target.guard.doc, documentKey(original));
    // A range rooted in selection produces unusable specifiers in current Word.
    // Keep every main-story range rooted in its named document. Other stories
    // must be matched to the actual selected story, not just its story type.
    let story;
    if (target.guard.story !== 'main text story') {
      story = a.getStoryRange(original, { storyType: target.guard.story });
      for (let i = 0; i < 128; i++) {
        if (a.inStory(story, { textRange: sel.textObject })) break;
        story = story.nextStoryRange;
        if (i === 127) fail('无法确认此处的文字范围，请在正文中插入格式化内容。');
      }
    }
    const originalRange = (start, end) => {
      try {
        const range = story
          ? a.setRange(story, { start, end })
          : a.createRange(original, { start, end });
        // Word may return an unresolved range after Save As even when document
        // properties remain readable. Resolve it before any original-text write.
        range.startOfContent();
        return range;
      } catch (error) {
        if (
          Number(error.errorNumber) === -1728 ||
          /(?:^|\D)-1728(?:$|\D)/.test(String(error.message))
        )
          fail('Word 无法定位这份文档的文字范围。请保存并重新打开该文档，再重新读取选区。');
        throw error;
      }
    };
    const selected = originalRange(target.guard.start, target.guard.end);
    if (
      selected.startOfContent() !== target.guard.start ||
      selected.endOfContent() !== target.guard.end ||
      String(selected.content() || '') !== target.text ||
      (story && !a.isEquivalent(selected, { textRange: sel.textObject }))
    )
      fail('无法确认选区的文字范围，请重新读取并确认写入位置。');
    const destinationStart = q.kind === 'insert' ? target.guard.end : target.guard.start;
    let destinationEnd = target.guard.end;
    let prefixMark;
    if (q.block) {
      const paragraph = originalRange(destinationStart, destinationStart).paragraphs[0].textObject;
      if (destinationStart > paragraph.startOfContent()) {
        const end = paragraph.endOfContent();
        prefixMark = originalRange(end - 1, end);
        if (String(prefixMark.content()) !== '\r') fail('请在普通文字段落中插入格式化内容。');
      }
      const ending = originalRange(destinationEnd, destinationEnd).paragraphs[0].textObject;
      if (
        !(q.kind === 'replace' && target.text.endsWith('\r')) &&
        destinationEnd === ending.endOfContent() - 1 &&
        String(ending.content()).endsWith('\r') &&
        optional(() => ending.tables.length, 1) === 0 &&
        String(originalRange(destinationEnd, destinationEnd + 1).content()) === '\r'
      )
        destinationEnd++;
    }
    same(q.guard, current().guard);
    const sourceName = String(q.file).split('/').pop();
    if (wordDocumentKeys(a).some((key) => key[0] === sourceName))
      fail('格式化临时文档已打开，请重新生成。');
    let source;
    try {
      a.open(Path(q.file), { readOnly: false, addToRecentFiles: false, confirmConversions: false });
      const sourceKeys = wordDocumentKeys(a).filter((key) => key[0] === sourceName);
      if (sourceKeys.length !== 1) fail('无法唯一定位格式化临时文档，请重新生成。');
      source = wordDocument(a, sourceKeys[0]);
      const openedKey = documentKey(a.activeDocument);
      if (
        JSON.stringify(openedKey) !== JSON.stringify(documentKey(source)) &&
        JSON.stringify(openedKey) !== JSON.stringify(target.guard.doc)
      )
        fail('目标文档已改变，请重新读取并确认写入位置。');
      // Word's range commands can move the active selection even when passed a
      // document-rooted range. Prepare paragraphs only while the source is active.
      a.activateObject(source);
      same(documentKey(source), documentKey(a.activeDocument));
      if (prefixMark) {
        const start = source.textObject.startOfContent();
        a.insertParagraph({ at: a.createRange(source, { start, end: start }) });
        const leading = a.createRange(source, { start, end: start + 1 });
        leading.formattedText = prefixMark.formattedText;
      }
      const sourceStart = source.textObject.startOfContent();
      let sourceEnd = source.textObject.endOfContent();
      const selectedParagraphEnd = q.kind === 'replace' && target.text.endsWith('\r');
      if (!q.block && !selectedParagraphEnd) {
        if (!String(source.textObject.content()).endsWith('\r'))
          fail('无法确认格式化内容的段落边界。');
        sourceEnd--;
      }
      if (sourceEnd <= sourceStart) fail('格式化内容为空，请重新生成。');
      const content = a.createRange(source, { start: sourceStart, end: sourceEnd });
      const activeKey = documentKey(a.activeDocument);
      if (JSON.stringify(activeKey) === JSON.stringify(documentKey(source)))
        a.activateObject(original);
      else same(target.guard.doc, activeKey);
      same(q.guard, current().guard);
      if (optional(() => original.readOnly(), false)) fail('文档为只读，请打开可编辑副本。');
      if (q.kind === 'replace' && !current().writable) fail('这段选区包含特殊对象，请使用插入。');
      const destination = originalRange(destinationStart, destinationEnd);
      same(q.guard, current().guard);
      const expected =
        (q.kind === 'insert' ? '' : target.text) + (destinationEnd > target.guard.end ? '\r' : '');
      if (
        String(destination.content() || '') !== expected ||
        (story && !a.inStory(destination, { textRange: sel.textObject }))
      )
        fail('写入位置的段落边界已改变，请重新读取并确认写入位置。');
      wordDocument(a, target.guard.doc);
      destination.formattedText = content.formattedText;
    } finally {
      if (source) {
        const activeKey = optional(() => documentKey(a.activeDocument), null);
        const restore =
          JSON.stringify(activeKey) === JSON.stringify(target.guard.doc) ||
          JSON.stringify(activeKey) === JSON.stringify(documentKey(source));
        a.close(source, { saving: 'no' });
        if (
          restore &&
          JSON.stringify(documentKey(a.activeDocument)) !== JSON.stringify(target.guard.doc)
        )
          a.activateObject(original);
      }
    }
    return;
  }
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
  // Excel's macOS dictionary declares countLarge as null. Collection counts
  // work through JXA; check dimensions too because cells.length can overflow.
  const rows = Number(r.rows.length),
    columns = Number(r.columns.length),
    count = Number(r.cells.length);
  if (rows * columns > 2000 || count > 2000) fail('请把选区缩小到 2,000 个单元格以内。');
  if (
    !Number.isSafeInteger(rows) ||
    !Number.isSafeInteger(columns) ||
    !Number.isSafeInteger(count) ||
    rows < 1 ||
    columns < 1 ||
    count < 1 ||
    rows * columns !== count
  )
    fail('无法确认选区大小，请选择一块连续的单元格区域。');
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
  if (q.kind !== 'formula' || count !== 1) fail('请选择一个公式目标单元格。');
  if (optional(() => r.mergeCells(), false) || optional(() => r.hasArray(), false))
    fail('请选择普通单元格，不能覆盖合并区域或数组公式。');
  r.formula = String(q.value);
}
function pptHex(rgb) {
  if (!Array.isArray(rgb) || rgb.length < 3) return null;
  return (
    '#' +
    rgb
      .slice(0, 3)
      .map((c) =>
        Math.round(c > 255 ? c / 257 : c)
          .toString(16)
          .padStart(2, '0'),
      )
      .join('')
      .toUpperCase()
  );
}
function pptBackground(slide) {
  if (slide.followMasterBackground()) return { kind: 'master' };
  const fill = slide.background.fillFormat,
    type = String(fill.fillFormatType());
  if (type === 'fill solid' && !fill.transparency())
    return { kind: 'solid', color: pptHex(fill.foreColor()) };
  if (
    type === 'fill gradient' &&
    String(fill.gradientColorType()) === 'two colors gradient type' &&
    String(fill.gradientStyle()) === 'vertical gradient' &&
    fill.gradientVariant() === 1
  )
    return { kind: 'gradient', from: pptHex(fill.foreColor()), to: pptHex(fill.backColor()) };
  return { kind: 'other', label: type };
}
function pptSetBackground(app, slide, style) {
  if (style.kind === 'master') {
    slide.followMasterBackground = true;
    return;
  }
  const rgb = (hex) => [1, 3, 5].map((i) => parseInt(hex.slice(i, i + 2), 16));
  const fill = slide.background.fillFormat;
  slide.followMasterBackground = false;
  if (style.kind === 'solid') {
    app.solid(fill);
    fill.foreColor = rgb(style.color);
    fill.transparency = 0;
  } else if (style.kind === 'gradient') {
    app.twoColorGradient(fill, { style: 'vertical gradient', variant: 1 });
    fill.foreColor = rgb(style.from);
    fill.backColor = rgb(style.to);
  } else fail('请在 PPTX 工作副本中使用此背景。');
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
        shape.role = /footer|页脚/i.test(name)
          ? 'footer'
          : /title|标题/i.test(name)
            ? 'title'
            : 'body';
        shape.font = {
          name: optional(() => r.font.fontName(), null),
          size: optional(() => r.font.fontSize(), null),
          bold: optional(() => r.font.bold(), null),
          color: pptHex(colors),
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
      background: optional(() => pptBackground(slide), undefined),
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
  if (q.op === 'presentationCopy') {
    const current = a.activePresentation,
      path = String(current.fullName());
    if (!current.saved() || !/\.pptx$/i.test(path) || path[0] !== '/')
      fail('请先在 PowerPoint 中将当前演示稿保存为本地 PPTX，再读取工作副本。');
    if (!$.NSFileManager.defaultManager.copyItemAtPathToPathError(path, q.file, null))
      fail('无法读取当前文件，请通过文件选择器打开 PPTX。');
    return;
  }
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
  const namedDoc = a.presentations.byName(String(doc.name()));
  same(documentKey(doc), documentKey(namedDoc));
  const backgrounds = (q.proposal.actions || []).map((action) => {
    const target = now.snapshot.slides.find((s) => s.id === action.slideId);
    if (
      action.kind !== 'background' ||
      !target ||
      !target.background ||
      target.background.kind === 'other'
    )
      fail('使用 PPTX 工作副本处理此页面操作。');
    const fill =
      action.fill.kind === 'reference'
        ? (now.snapshot.slides.find((s) => s.id === action.fill.sourceId) || {}).background
        : action.fill;
    if (!fill || !['master', 'solid', 'gradient'].includes(fill.kind))
      fail('请通过 PPTX 工作副本处理图片或模板背景。');
    return { target, fill };
  });
  const rollback = [];
  function set(obj, key, v) {
    const old = obj[key]();
    obj[key] = v;
    rollback.push(() => {
      obj[key] = old;
    });
  }
  try {
    for (const task of backgrounds) {
      const slide = namedDoc.slides[task.target.index];
      rollback.push(() => pptSetBackground(a, slide, task.target.background));
      pptSetBackground(a, slide, task.fill);
      same(task.fill, pptBackground(slide));
    }
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
          [1, 3, 5].map((i) => parseInt(p.fontColor.slice(i, i + 2), 16)),
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
