// The browser export is a self-contained ESM build and also runs in native Node.
// It avoids relying on browser APIs added after the supported Office WebViews.
import MarkdownIt from 'markdown-it/browser';

const markdown = new MarkdownIt({ html: false, linkify: false, typographer: false });
const inlineMarkers = /[`*_~\[|]/;
const blockMarkers =
  /^(?: {0,3}(?:#{1,6}(?:[ \t]|$)|>|(?:[+-]|\d{1,9}[.)])(?:[ \t]|$)|(?:={2,}|-{3,})[ \t]*$)|(?: {4}|\t)[ \t]*\S)/m;
const hardBreak = / {2}\r?\n|\\\r?\n/;
const explicitAutolink = /<https?:\/\//i;

const paragraphStyle = 'margin:0 0 8pt 0;line-height:115%;';
const codeStyle = 'font-family:Consolas,Menlo,monospace;font-size:10pt;mso-no-proof:yes;';
const cellStyle = 'border:1px solid #b7b7b7;padding:5pt 7pt;vertical-align:top;';
const headingSizes = [20, 16, 14, 12, 11, 11];
const richTokens = new Set([
  'heading_open',
  'bullet_list_open',
  'ordered_list_open',
  'table_open',
  'blockquote_open',
  'fence',
  'code_block',
  'hr',
  'strong_open',
  'em_open',
  's_open',
  'code_inline',
  'link_open',
  'image',
  'hardbreak',
]);

function escapeHtml(text) {
  return text.replace(
    /[&<>"']/g,
    (value) =>
      ({
        '&': '&amp;',
        '<': '&lt;',
        '>': '&gt;',
        '"': '&quot;',
        "'": '&#39;',
      })[value],
  );
}

function hasFormatting(tokens) {
  return tokens.some(
    (token) => richTokens.has(token.type) || (token.children && hasFormatting(token.children)),
  );
}

function imageDescription(tokens) {
  return tokens
    .map((token) => {
      if (token.type === 'softbreak' || token.type === 'hardbreak') return ' ';
      if (token.children) return imageDescription(token.children);
      return token.content || '';
    })
    .join('');
}

// Link destinations are validated and normalized before attribute escaping.
function safeLink(value) {
  if (typeof value !== 'string' || !/^https?:\/\//i.test(value)) return null;
  try {
    const url = new URL(value);
    if (!['http:', 'https:'].includes(url.protocol) || url.username || url.password) return null;
    return url.href;
  } catch {
    return null;
  }
}

// Only allowlisted semantic tags, fixed styles and validated links are emitted.
// Never copy raw HTML, arbitrary attributes, image URLs or fence languages.
function renderTokens(tokens) {
  const html = [];
  const links = [];
  for (const token of tokens) {
    switch (token.type) {
      case 'inline':
        html.push(renderTokens(token.children || []));
        break;
      case 'text':
        html.push(escapeHtml(token.content));
        break;
      case 'softbreak':
      case 'hardbreak':
        html.push('<br>');
        break;
      case 'paragraph_open':
        html.push(
          `<p class="MsoNormal" style="${token.hidden ? 'margin:0 0 3pt 0;' : paragraphStyle}">`,
        );
        break;
      case 'paragraph_close':
        html.push('</p>');
        break;
      case 'heading_open': {
        const level = Number(token.tag.slice(1));
        if (!Number.isInteger(level) || level < 1 || level > 6) throw new Error('Invalid heading');
        html.push(
          `<h${level} class="MsoHeading${level}" style="mso-style-name:'Heading ${level}';mso-outline-level:${level};font-size:${headingSizes[level - 1]}pt;font-weight:bold;color:#000;margin:12pt 0 6pt 0;page-break-after:avoid;">`,
        );
        break;
      }
      case 'heading_close':
        html.push(`</${token.tag}>`);
        break;
      case 'strong_open':
        html.push('<strong>');
        break;
      case 'strong_close':
        html.push('</strong>');
        break;
      case 'em_open':
        html.push('<em>');
        break;
      case 'em_close':
        html.push('</em>');
        break;
      case 's_open':
        html.push('<s>');
        break;
      case 's_close':
        html.push('</s>');
        break;
      case 'code_inline':
        html.push(
          `<code style="${codeStyle}background-color:#f2f2f2;">${escapeHtml(token.content)}</code>`,
        );
        break;
      case 'fence':
      case 'code_block':
        html.push(
          `<pre style="${codeStyle}white-space:pre-wrap;margin:6pt 0 8pt 0;padding:8pt;background-color:#f2f2f2;"><code>${escapeHtml(token.content)}</code></pre>`,
        );
        break;
      case 'bullet_list_open':
        html.push('<ul style="margin:0 0 8pt 18pt;padding-left:18pt;">');
        break;
      case 'bullet_list_close':
        html.push('</ul>');
        break;
      case 'ordered_list_open': {
        const start = Number(token.attrGet('start') ?? 1);
        const attribute =
          Number.isSafeInteger(start) && start >= 0 && start !== 1 ? ` start="${start}"` : '';
        html.push(`<ol${attribute} style="margin:0 0 8pt 18pt;padding-left:18pt;">`);
        break;
      }
      case 'ordered_list_close':
        html.push('</ol>');
        break;
      case 'list_item_open':
        html.push('<li style="margin:0 0 3pt 0;">');
        break;
      case 'list_item_close':
        html.push('</li>');
        break;
      case 'blockquote_open':
        html.push(
          '<blockquote class="MsoQuote" style="margin:6pt 0 8pt 18pt;border-left:2pt solid #b7b7b7;padding-left:10pt;">',
        );
        break;
      case 'blockquote_close':
        html.push('</blockquote>');
        break;
      case 'hr':
        html.push('<hr style="border:0;border-top:1px solid #b7b7b7;">');
        break;
      case 'table_open':
        html.push('<table style="border-collapse:collapse;width:100%;margin:6pt 0 8pt 0;">');
        break;
      case 'table_close':
        html.push('</table>');
        break;
      case 'thead_open':
        html.push('<thead>');
        break;
      case 'thead_close':
        html.push('</thead>');
        break;
      case 'tbody_open':
        html.push('<tbody>');
        break;
      case 'tbody_close':
        html.push('</tbody>');
        break;
      case 'tr_open':
        html.push('<tr>');
        break;
      case 'tr_close':
        html.push('</tr>');
        break;
      case 'th_open':
      case 'td_open': {
        const alignment = /^text-align:(left|center|right)$/.exec(
          token.attrGet('style') || '',
        )?.[1];
        const header = token.type === 'th_open';
        html.push(
          `<${header ? 'th' : 'td'}${header ? ' scope="col"' : ''} style="${cellStyle}${header ? 'font-weight:bold;background-color:#f2f2f2;' : ''}${alignment ? `text-align:${alignment};` : ''}">`,
        );
        break;
      }
      case 'th_close':
        html.push('</th>');
        break;
      case 'td_close':
        html.push('</td>');
        break;
      // Explicit HTTP(S) citations remain usable; no other scheme, relative
      // reference, embedded credentials or remote image enters the document.
      case 'link_open': {
        const href = safeLink(token.attrGet('href'));
        links.push(Boolean(href));
        if (href)
          html.push(`<a href="${escapeHtml(href)}" target="_blank" rel="noopener noreferrer">`);
        break;
      }
      case 'link_close':
        if (links.pop()) html.push('</a>');
        break;
      case 'image':
        html.push(escapeHtml(imageDescription(token.children || []) || token.content || '图片'));
        break;
      default:
        html.push(escapeHtml(token.content || ''));
    }
  }
  return html.join('');
}

function markdownWrapper(tokens, text) {
  if (
    tokens.length !== 1 ||
    tokens[0].type !== 'fence' ||
    !/^(?:markdown|md)$/i.test(tokens[0].info.trim())
  )
    return false;
  const trimmed = text.trimEnd();
  const lastLine = trimmed.slice(trimmed.lastIndexOf('\n') + 1);
  const close = /^ {0,3}(`{3,}|~{3,})[ \t]*$/.exec(lastLine)?.[1];
  return Boolean(
    close && close[0] === tokens[0].markup[0] && close.length >= tokens[0].markup.length,
  );
}

/** Parse once for insertion: single-paragraph inline markup must not split a Word sentence. */
export function wordMarkdownFormat(text) {
  if (typeof text !== 'string' || !text.length) return null;
  // The common 100k–2m character Word response should not enter a Markdown
  // parser merely because it contains newlines, Chinese text or literal <...>.
  if (
    !inlineMarkers.test(text) &&
    !blockMarkers.test(text) &&
    !hardBreak.test(text) &&
    !explicitAutolink.test(text)
  )
    return null;
  try {
    let tokens = markdown.parse(text, {});
    const wrapped = markdownWrapper(tokens, text);
    if (wrapped) tokens = markdown.parse(tokens[0].content, {});
    if (!tokens.length || (!wrapped && !hasFormatting(tokens))) return null;
    const inline =
      tokens.length === 3 &&
      tokens[0].type === 'paragraph_open' &&
      tokens[1].type === 'inline' &&
      tokens[2].type === 'paragraph_close';
    return { html: renderTokens(inline ? tokens[1].children || [] : tokens), block: !inline };
  } catch {
    // Unsupported/pathological syntax remains intact through insertText.
    return null;
  }
}

/** Safe HTML fragment for preview/import; null means keep the exact plain text. */
export function wordMarkdownHtml(text) {
  return wordMarkdownFormat(text)?.html ?? null;
}
