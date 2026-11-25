// All document content is escaped as RTF text. Only fixed formatting controls are emitted.
const escape = (value) =>
  String(value)
    .split('')
    .map((c) => {
      if ('\\{}'.includes(c)) return '\\' + c;
      if (c === '\n' || c === '\r') return '\\line ';
      const n = c.charCodeAt(0);
      return n > 126 ? '\\u' + (n > 32767 ? n - 65536 : n) + '?' : n < 32 ? ' ' : c;
    })
    .join('');
export function reportRtf(report) {
  if (!report || !Array.isArray(report.sections) || !report.sections.length)
    throw Error('报告结构无效。');
  const paragraph = (text, size, bold = false) =>
    '{\\pard\\sa100\\fs' + size + (bold ? '\\b ' : ' ') + escape(text) + '\\par}';
  const rows = (r) => {
    const width = Math.floor(8500 / r.length);
    return (
      '{\\trowd\\trgaph80' +
      r.map((_, i) => '\\cellx' + (i + 1) * width).join('') +
      ' ' +
      r.map((cell) => '\\pard\\intbl ' + escape(cell) + '\\cell ').join('') +
      '\\row}'
    );
  };
  const out = ['{\\rtf1\\ansi\\uc1 ', paragraph(report.title, 36, true)];
  for (const section of report.sections) {
    out.push(paragraph(section.heading, 28, true));
    for (const p of section.paragraphs) out.push(paragraph(p, 22));
    for (const b of section.bullets) out.push(paragraph('• ' + b, 22));
    if (section.table)
      for (const row of [section.table.headers, ...section.table.rows]) out.push(rows(row));
  }
  return out.join('') + '}';
}
