import { describe, expect, it, vi } from 'vitest';
import MarkdownIt from 'markdown-it/browser';
import { wordMarkdownFormat, wordMarkdownHtml } from '../offline/app/word-markdown.mjs';

function fragment(markdown: string) {
  const html = wordMarkdownHtml(markdown);
  expect(html).not.toBeNull();
  expect(html).not.toMatch(/<\/?(?:html|head|body|style)(?:\s|>)/i);
  const template = document.createElement('template');
  template.innerHTML = html!;
  return template.content;
}

describe('Word Markdown HTML', () => {
  it.each([
    '',
    ' \n\t ',
    '普通正文𠮷<纯文本>\n保留换行。\n\n第二段。',
    'model_id、snake_case 与 2026-11-25 不表示格式。',
    'https://example.com/path?a=1&b=2',
    '一处未闭合的 ** 标记',
    'a | b\n只有正文，没有表格分隔行。',
    '<script>alert("literal HTML is plain text")</script>',
  ])('preserves the exact plain-text path for %j', (text) => {
    expect(wordMarkdownHtml(text)).toBeNull();
  });

  it('bypasses the parser for large Unicode plain text rather than slowing or changing its write path', () => {
    const parse = vi.spyOn(MarkdownIt.prototype, 'parse');
    for (const length of [100_001, 2_000_000]) {
      const unit = '正文𠮷<纯文本>\n第二行。\n';
      const text = unit.repeat(Math.ceil(length / unit.length)).slice(0, length);
      expect(text.length).toBe(length);
      expect(wordMarkdownHtml(text)).toBeNull();
    }
    expect(wordMarkdownHtml(' '.repeat(2_000_000))).toBeNull();
    expect(parse).not.toHaveBeenCalled();
  });

  it('emits semantic heading levels with Word heading and outline style hints', () => {
    const root = fragment(
      [1, 2, 3, 4, 5, 6].map((level) => `${'#'.repeat(level)} 标题 ${level}`).join('\n\n'),
    );
    for (const level of [1, 2, 3, 4, 5, 6]) {
      const heading = root.querySelector(`h${level}`)!;
      expect(heading.textContent).toBe(`标题 ${level}`);
      expect(heading.className).toBe(`MsoHeading${level}`);
      expect(heading.getAttribute('style')).toContain(`mso-style-name:'Heading ${level}'`);
      expect(heading.getAttribute('style')).toContain(`mso-outline-level:${level}`);
    }
  });

  it('keeps paragraphs, nested inline emphasis, entities, Unicode and line breaks editable', () => {
    const root = fragment(
      '普通 **粗体与 *斜体***，~~删去~~，𠮷 &amp; &lt;字&gt;。\n下一行。\n\n第二段 `a < b && c`。',
    );
    expect(root.querySelectorAll('p')).toHaveLength(2);
    expect(root.querySelector('strong em')?.textContent).toBe('斜体');
    expect(root.querySelector('s')?.textContent).toBe('删去');
    expect(root.querySelectorAll('br')).toHaveLength(1);
    expect(root.textContent).toContain('𠮷 & <字>');
    expect(root.querySelector('code')?.textContent).toBe('a < b && c');
    expect(root.textContent).not.toContain('**');
  });

  it('creates real numbered and nested bullet lists and preserves non-one starting numbers', () => {
    const root = fragment(
      '3. 第三项\n4. 第四项\n   - 子项 **强调**\n   - 子项二\n\n- 新列表\n\n  同一项的另一段。',
    );
    const numbered = root.querySelector('ol')!;
    expect(numbered.getAttribute('start')).toBe('3');
    expect(numbered.querySelectorAll(':scope > li')).toHaveLength(2);
    expect(numbered.querySelectorAll('ul > li')).toHaveLength(2);
    expect(numbered.querySelector('ul strong')?.textContent).toBe('强调');
    expect(root.textContent).toContain('同一项的另一段。');
    expect(root.textContent).not.toContain('3.');
  });

  it('creates editable pipe tables with alignment, escaped pipes and formatted cells', () => {
    const root = fragment(
      '| 名称 | 数据 |\n| :--- | ---: |\n| A\\|B | **3** |\n| 第二项 | `x\\|y` |',
    );
    expect(root.querySelectorAll('table')).toHaveLength(1);
    expect(root.querySelectorAll('thead th')).toHaveLength(2);
    expect(root.querySelectorAll('tbody tr')).toHaveLength(2);
    expect(root.querySelector('tbody td')?.textContent).toBe('A|B');
    expect(root.querySelector('td strong')?.textContent).toBe('3');
    expect(root.querySelector('td code')?.textContent).toBe('x|y');
    expect(root.querySelector('th')?.style.textAlign).toBe('left');
    expect(root.querySelectorAll('th')[1].style.textAlign).toBe('right');
  });

  it('preserves fenced code spacing and escapes its text without using its language as attributes', () => {
    const root = fragment(
      '```html" onmouseover="evil\nconst s = "<img src=x>";\n  a & b\n```\n\n内联 `&amp; <script>`。',
    );
    expect(root.querySelector('pre code')?.textContent).toBe('const s = "<img src=x>";\n  a & b\n');
    expect(root.querySelector('p code')?.textContent).toBe('&amp; <script>');
    expect(root.querySelector('[onmouseover], img, script')).toBeNull();
    expect(root.querySelector('pre')?.getAttribute('style')).toContain('white-space:pre-wrap');
  });

  it('preserves safe citation links, prevents preview navigation, and escapes raw HTML and encoded attacks', () => {
    const root = fragment(
      [
        '# 安全 **标题**',
        '[可读链接](https://example.com/track?secret=1)',
        '![图片 **说明**](https://example.com/pixel.png)',
        '[危险](javascript:alert%281%29)',
        '[编码](jav&#x61;script:alert%281%29)',
        '[文件](file:///private/secret)',
        '[相对地址](/relative)',
        '[账号](https://user:password@example.com)',
        '![嵌入](data:image/svg+xml;base64,PHN2Zz4=)',
        '<img src="https://example.com/pixel" onerror="evil()">',
        '<iframe src="file:///private/secret"></iframe>',
        '<style>body{background:url(https://example.com/track)}</style>',
        '&#60;script&#62;alert(1)&#60;/script&#62;',
      ].join('\n\n'),
    );
    expect(root.textContent).toContain('可读链接');
    expect(root.textContent).toContain('图片 说明');
    expect(root.querySelectorAll('a')).toHaveLength(1);
    expect(root.querySelector('a')?.getAttribute('href')).toBe(
      'https://example.com/track?secret=1',
    );
    expect(root.querySelector('a')?.getAttribute('target')).toBe('_blank');
    expect(root.querySelector('a')?.getAttribute('rel')).toBe('noopener noreferrer');
    expect(
      root.querySelector('img, iframe, script, style, link, object, embed, form, input, meta'),
    ).toBeNull();
    for (const element of root.querySelectorAll('*')) {
      for (const attribute of element.attributes) {
        expect(attribute.name).not.toMatch(/^on|^(?:src|srcset|action)$/i);
        if (attribute.name === 'style')
          expect(attribute.value).not.toMatch(/url\(|expression\(|javascript:/i);
      }
    }
  });

  it('distinguishes inline replacements from block insertions without guessing from the HTML string', () => {
    expect(wordMarkdownFormat('一段 **强调** 和 `代码`')).toEqual({
      html: expect.stringContaining('<strong>强调</strong>'),
      block: false,
    });
    expect(wordMarkdownHtml('一段 **强调**')).not.toMatch(/<p\b/);
    for (const text of ['# 标题', '- 列表', '**第一段**\n\n第二段', '```js\ncode\n```'])
      expect(wordMarkdownFormat(text)?.block).toBe(true);
    expect(wordMarkdownFormat('原样正文')).toBeNull();
  });

  it.each(['markdown', 'md', 'MD'])(
    'unwraps a single explicit %s fence containing the entire reply',
    (language) => {
      const root = fragment(`\n\n\`\`\`${language}\n# 标题\n\n**正文**\n\`\`\`\n`);
      expect(root.querySelector('h1')?.textContent).toBe('标题');
      expect(root.querySelector('strong')?.textContent).toBe('正文');
      expect(root.querySelector('pre')).toBeNull();
      expect(wordMarkdownFormat(`\`\`\`${language}\n只有正文\n\`\`\``)).toEqual({
        html: '只有正文',
        block: false,
      });
    },
  );

  it.each([
    '```text\n# sample\n```',
    '```\n# sample\n```',
    '说明\n\n```md\n# sample\n```',
    '```markdown\n# sample',
    '```md\n# sample\n```\n\n其他正文',
  ])('preserves ordinary code examples and incomplete or mixed wrappers: %j', (text) => {
    const root = fragment(text);
    expect(root.querySelector('pre')).not.toBeNull();
    expect(root.querySelector('h1')).toBeNull();
  });

  it('recognizes no-space quotes and indented code without treating paragraph continuation as code', () => {
    expect(fragment('>引用 **原文**').querySelector('blockquote strong')?.textContent).toBe('原文');
    expect(fragment('    code <tag>\n    second').querySelector('pre code')?.textContent).toBe(
      'code <tag>\nsecond\n',
    );
    expect(wordMarkdownHtml('普通正文\n    段落中的缩进续行')).toBeNull();
  });

  it('renders block quotes and rules while keeping ordinary content around them', () => {
    const root = fragment('前言\n\n> 引用 **原文**\n> 第二行\n\n---\n\n结尾');
    expect(root.querySelector('blockquote strong')?.textContent).toBe('原文');
    expect(root.querySelector('hr')).not.toBeNull();
    expect(root.textContent).toContain('前言');
    expect(root.textContent).toContain('结尾');
  });

  it('does not truncate complete long rich text at the former Word character limit', () => {
    const body = '中文𠮷<纯文本> & '.repeat(10_000) + '完整结尾';
    const root = fragment(`# 标题\n\n${body}`);
    expect(body.length).toBeGreaterThan(100_000);
    expect(root.querySelector('p')?.textContent).toBe(body);
    expect(root.textContent?.endsWith('完整结尾')).toBe(true);
  });
});
