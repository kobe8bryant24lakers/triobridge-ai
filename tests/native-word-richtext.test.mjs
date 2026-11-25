import { test } from 'node:test';
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import { runInNewContext } from 'node:vm';

const script = readFileSync(new URL('../offline/mac/office-mac.js', import.meta.url), 'utf8');
const word = runInNewContext(script + '\nword', { ObjC: { import() {} }, Path: (path) => path });

function fixture({
  text = 'prefixOLDsuffix\r',
  start = 6,
  end = 9,
  sourceText = 'new\r',
  story = 'main text story',
} = {}) {
  const writes = [],
    preparations = [];
  const docs = new Map();
  class Range {
    constructor(doc, start, end) {
      this.doc = doc;
      this.start = start;
      this.end = end;
    }
    startOfContent() {
      return this.start;
    }
    endOfContent() {
      return this.end ?? this.doc.text.length;
    }
    content() {
      return this.start === this.endOfContent()
        ? null
        : this.doc.text.slice(this.start, this.endOfContent());
    }
    get tables() {
      return { length: this.doc.text.includes('\u0007') ? 1 : 0 };
    }
    get paragraphs() {
      const begin = this.start === 0 ? 0 : this.doc.text.lastIndexOf('\r', this.start - 1) + 1;
      const next = this.doc.text.indexOf('\r', this.start);
      return [
        { textObject: new Range(this.doc, begin, next < 0 ? this.doc.text.length : next + 1) },
      ];
    }
    get formattedText() {
      return { text: this.content(), style: this.doc.prefixStyle || this.doc.style };
    }
    set formattedText(value) {
      const log = this.doc === original ? writes : preparations;
      log.push({ start: this.start, end: this.endOfContent(), ...value });
      this.doc.text =
        this.doc.text.slice(0, this.start) + value.text + this.doc.text.slice(this.endOfContent());
      if (this.doc !== original) this.doc.prefixStyle = value.style;
    }
  }
  function document(name, text, style) {
    const result = {
      text,
      style,
      name: () => name,
      fullName: () => '/tmp/' + name,
      readOnly: () => false,
    };
    Object.defineProperty(result, 'textObject', { get: () => new Range(result, 0, null) });
    docs.set(name, result);
    return result;
  }
  const original = document('original.docx', text, 'original paragraph style');
  const selection = { start, end, protected: false };
  const selected = {
    content: () => original.text.slice(selection.start, selection.end),
    selectionStart: () => selection.start,
    selectionEnd: () => selection.end,
    storyType: () => story,
    get textObject() {
      return new Range(original, selection.start, selection.end);
    },
  };
  for (const key of [
    'tables',
    'fields',
    'inlineShapes',
    'mathObjects',
    'formFields',
    'footnotes',
    'endnotes',
    'hyperlinkObjects',
  ])
    Object.defineProperty(selected, key, { get: () => ({ length: selection.protected ? 1 : 0 }) });
  const app = {
    active: original,
    // Deliberately delayed: opening a temporary document retargets this specifier.
    activeDocument: {
      name: () => app.active.name(),
      fullName: () => app.active.fullName(),
      readOnly: () => app.active.readOnly(),
    },
    selection: selected,
    documents: new Proxy(
      {},
      {
        get(_target, key) {
          if (key === 'length') return docs.size;
          if (key === 'byName')
            return (fullName) => {
              if (fullName === '/tmp/source.htm')
                throw Object.assign(Error('Word cannot resolve the /private/var import alias'), {
                  errorNumber: -1728,
                });
              if (fullName === 'source.htm') return docs.get('source.htm');
              const doc = [...docs.values()].find((doc) => doc.fullName() === fullName);
              if (!doc) throw Error('Word short byName keys fail after Save As');
              return doc;
            };
          if (key === 'name') return () => [...docs.values()].map(() => null);
          if (key === 'whose')
            return () => {
              throw Error('Word predicate specifiers cannot activate the document');
            };
          // Newly opened documents precede the original, as in the real Word app.
          return [...docs.values()].reverse()[Number(key)];
        },
      },
    ),
    open(file, options) {
      assert.equal(options.addToRecentFiles, false);
      assert.equal(options.confirmConversions, false);
      assert.equal(options.visible, undefined);
      app.source = document(file.split('/').pop(), sourceText, 'generated style');
      app.active = app.source;
      app.onOpen?.();
    },
    activateObject(doc) {
      app.active = doc;
      app.onActivate?.(doc);
    },
    createRange(doc, position) {
      return new Range(doc, position.start, position.end);
    },
    getStoryRange(doc, { storyType }) {
      assert.equal(storyType, story);
      return new Range(doc, 0, null);
    },
    inStory(range, { textRange }) {
      return range.doc === textRange.doc;
    },
    isEquivalent(range, { textRange }) {
      return (
        range.doc === textRange.doc &&
        range.start === textRange.start &&
        range.end === textRange.end
      );
    },
    setRange(range, position) {
      return new Range(range.doc, position.start, position.end);
    },
    collapseRange(range, { direction }) {
      const at = direction === 'collapse end' ? range.endOfContent() : range.startOfContent();
      return new Range(range.doc, at, at);
    },
    insertParagraph({ at }) {
      assert.notEqual(at.doc, original, 'paragraph preparation must only edit the source');
      assert.equal(app.active, at.doc, 'Word range preparation requires the source to be active');
      at.formattedText = { text: '\r', style: 'generated style' };
    },
    close(doc, options) {
      assert.equal(options.saving, 'no');
      docs.delete(doc.name());
      app.closed = true;
      app.onClose?.();
    },
  };
  const guard = word(app, { op: 'target' }).guard;
  return { app, original, selection, writes, preparations, guard, docs };
}

test('Mac Word saved documents tolerate broken byName and null collection names', () => {
  const f = fixture();
  f.docs.set('same-name-other-file', {
    name: f.original.name,
    fullName: () => '/other/original.docx',
  });
  word(f.app, {
    op: 'write',
    kind: 'replace',
    guard: f.guard,
    file: '/tmp/source.htm',
    block: false,
  });
  assert.equal(f.original.text, 'prefixnewsuffix\r');
  assert.equal(f.writes.length, 1);
  assert.equal(f.app.active, f.original);
});

test('Mac Word refuses ambiguous full file identities before writing', () => {
  for (const duringImport of [false, true]) {
    const f = fixture();
    const duplicate = () =>
      f.docs.set('duplicate', {
        name: f.original.name,
        fullName: f.original.fullName,
      });
    if (duringImport) f.app.onOpen = duplicate;
    else duplicate();
    assert.throws(
      () =>
        word(f.app, {
          op: 'write',
          kind: 'replace',
          guard: f.guard,
          file: '/tmp/source.htm',
          block: false,
        }),
      /无法唯一定位目标 Word 文档/,
    );
    assert.equal(f.original.text, 'prefixOLDsuffix\r');
    assert.equal(f.writes.length, 0);
  }
});

test('Mac Word unresolved range errors explain the Save As recovery without masking other failures', () => {
  for (const [error, expected] of [
    [Object.assign(Error('The object does not exist'), { errorNumber: -1728 }), /保存并重新打开/],
    [Error('Error -1728'), /保存并重新打开/],
    [Error('automation permission denied'), /automation permission denied/],
  ]) {
    const f = fixture();
    f.app.createRange = () => ({
      startOfContent() {
        throw error;
      },
    });
    assert.throws(
      () =>
        word(f.app, {
          op: 'write',
          kind: 'replace',
          guard: f.guard,
          file: '/tmp/source.htm',
          block: false,
        }),
      expected,
    );
    assert.equal(f.writes.length, 0);
  }
});

test('Mac Word inline formatting replaces exactly once and never imports the source final paragraph', () => {
  const f = fixture();
  word(f.app, {
    op: 'write',
    kind: 'replace',
    guard: f.guard,
    file: '/tmp/source.htm',
    block: false,
  });
  assert.equal(f.original.text, 'prefixnewsuffix\r');
  assert.equal(f.writes.length, 1);
  assert.equal(f.preparations.length, 0);
  assert.equal(f.app.active, f.original);
  assert.equal(f.app.closed, true);
});

test('Mac Word inline formatting preserves a selected paragraph break', () => {
  const f = fixture({ text: 'OLD\rnext\r', start: 0, end: 4 });
  word(f.app, {
    op: 'write',
    kind: 'replace',
    guard: f.guard,
    file: '/tmp/source.htm',
    block: false,
  });
  assert.equal(f.original.text, 'new\rnext\r');
});

test('Mac Word formatted insertion stays at the exact selection end', () => {
  const f = fixture();
  word(f.app, {
    op: 'write',
    kind: 'insert',
    guard: f.guard,
    file: '/tmp/source.htm',
    block: false,
  });
  assert.equal(f.original.text, 'prefixOLDnewsuffix\r');
  assert.equal(f.writes[0].start, 9);
  assert.equal(f.writes[0].end, 9);
});

test('Mac Word block preparation isolates the prefix in the source and preserves its paragraph format', () => {
  const f = fixture({ sourceText: 'Heading\rBody\r' });
  word(f.app, {
    op: 'write',
    kind: 'replace',
    guard: f.guard,
    file: '/tmp/source.htm',
    block: true,
  });
  assert.equal(f.original.text, 'prefix\rHeading\rBody\rsuffix\r');
  assert.equal(f.writes.length, 1);
  assert.equal(f.preparations.at(-1).text, '\r');
  assert.equal(f.preparations.at(-1).style, 'original paragraph style');
});

test('Mac Word whole paragraph block replacement does not add an empty leading paragraph', () => {
  const f = fixture({ text: 'OLD\rnext\r', start: 0, end: 4, sourceText: 'Heading\rBody\r' });
  word(f.app, {
    op: 'write',
    kind: 'replace',
    guard: f.guard,
    file: '/tmp/source.htm',
    block: true,
  });
  assert.equal(f.original.text, 'Heading\rBody\rnext\r');
  assert.equal(f.preparations.length, 0);
});

test('Mac Word blocks consume only an immediately adjacent ordinary paragraph mark', () => {
  for (const [text, start, end, kind, expected] of [
    ['OLD\rnext\r', 0, 3, 'replace', 'Heading\rBody\rnext\r'],
    ['\rnext\r', 0, 0, 'insert', 'Heading\rBody\rnext\r'],
    ['prefix\rnext\r', 6, 6, 'insert', 'prefix\rHeading\rBody\rnext\r'],
    ['OLD\r\rnext\r', 0, 4, 'replace', 'Heading\rBody\r\rnext\r'],
  ]) {
    const f = fixture({ text, start, end, sourceText: 'Heading\rBody\r' });
    word(f.app, { op: 'write', kind, guard: f.guard, file: '/tmp/source.htm', block: true });
    assert.equal(f.original.text, expected);
    assert.equal(f.writes.length, 1);
  }
});

test('Mac Word rich insertion does not use selection-derived or collapsing range commands', () => {
  const f = fixture({ start: 9, end: 9 });
  f.app.setRange = f.app.collapseRange = () => {
    throw Error('Word selection-derived ranges are not stable');
  };
  word(f.app, {
    op: 'write',
    kind: 'insert',
    guard: f.guard,
    file: '/tmp/source.htm',
    block: false,
  });
  assert.equal(f.original.text, 'prefixOLDnewsuffix\r');
  assert.equal(f.selection.start, 9);
  assert.equal(f.selection.end, 9);
});

test('Mac Word normalizes a cursor even when selection content exposes the next character', () => {
  const f = fixture({ start: 9, end: 9 });
  f.app.selection.content = () => 's';
  const target = word(f.app, { op: 'target' });
  assert.equal(target.guard.text, '');
  assert.equal(target.text, '');
  word(f.app, {
    op: 'write',
    kind: 'insert',
    guard: target.guard,
    file: '/tmp/source.htm',
    block: false,
  });
  assert.equal(f.original.text, 'prefixOLDnewsuffix\r');
});

test('Mac Word non-main-story range must match the selected story before writing', () => {
  const f = fixture({ story: 'footnotes story' });
  f.app.createRange = (doc, position) => {
    assert.notEqual(doc, f.original, 'never use a main-story range for footnote positions');
    return f.app.setRange(doc.textObject, position);
  };
  word(f.app, {
    op: 'write',
    kind: 'replace',
    guard: f.guard,
    file: '/tmp/source.htm',
    block: false,
  });
  assert.equal(f.original.text, 'prefixnewsuffix\r');
  const mismatch = fixture({ story: 'footnotes story' });
  mismatch.app.isEquivalent = () => false;
  assert.throws(
    () =>
      word(mismatch.app, {
        op: 'write',
        kind: 'replace',
        guard: mismatch.guard,
        file: '/tmp/source.htm',
        block: false,
      }),
    /无法确认选区/,
  );
  assert.equal(mismatch.writes.length, 0);
});

test('Mac Word restores the target only when closing its own source changed focus', () => {
  const f = fixture();
  f.app.onClose = () => {
    f.app.active = { name: () => 'other.docx', fullName: () => '/tmp/other.docx' };
  };
  word(f.app, {
    op: 'write',
    kind: 'replace',
    guard: f.guard,
    file: '/tmp/source.htm',
    block: false,
  });
  assert.equal(f.app.active, f.original);
});

test('Mac Word does not write or steal focus after the user activates another document', () => {
  const f = fixture();
  const other = { name: () => 'other.docx', fullName: () => '/tmp/other.docx' };
  f.app.onActivate = (doc) => {
    if (doc !== f.original) f.app.active = other;
  };
  assert.throws(() =>
    word(f.app, {
      op: 'write',
      kind: 'replace',
      guard: f.guard,
      file: '/tmp/source.htm',
      block: true,
    }),
  );
  assert.equal(f.app.active, other);
  assert.equal(f.writes.length, 0);
  assert.equal(f.preparations.length, 0);
  assert.equal(f.app.closed, true);
});

test('Mac Word failed HTML conversion leaves the selected original text untouched', () => {
  const f = fixture();
  f.app.open = () => {
    throw Error('conversion failed');
  };
  assert.throws(
    () =>
      word(f.app, {
        op: 'write',
        kind: 'replace',
        guard: f.guard,
        file: '/tmp/source.htm',
        block: true,
      }),
    /conversion failed/,
  );
  assert.equal(f.original.text, 'prefixOLDsuffix\r');
  assert.equal(f.writes.length, 0);
});

test('Mac Word rechecks selection and protected objects after opening generated HTML', () => {
  for (const mutate of [
    (f) => {
      f.selection.start++;
    },
    (f) => {
      f.original.text = 'changed original\r';
    },
    (f) => {
      f.selection.protected = true;
    },
  ]) {
    const f = fixture();
    f.app.onOpen = () => mutate(f);
    assert.throws(() =>
      word(f.app, {
        op: 'write',
        kind: 'replace',
        guard: f.guard,
        file: '/tmp/source.htm',
        block: false,
      }),
    );
    assert.equal(f.writes.length, 0);
    assert.equal(f.app.closed, true);
  }
});

test('Mac Word rejects a changed adjacent paragraph mark before the single formatted write', () => {
  const f = fixture({ text: 'OLD\rnext\r', start: 0, end: 3, sourceText: 'Heading\rBody\r' });
  f.app.onOpen = () => {
    f.original.text = 'OLDXnext\r';
  };
  assert.throws(
    () =>
      word(f.app, {
        op: 'write',
        kind: 'replace',
        guard: f.guard,
        file: '/tmp/source.htm',
        block: true,
      }),
    /段落边界已改变/,
  );
  assert.equal(f.original.text, 'OLDXnext\r');
  assert.equal(f.writes.length, 0);
  assert.equal(f.app.closed, true);
});
