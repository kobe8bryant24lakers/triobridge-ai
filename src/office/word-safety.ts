const WORD_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const MATH_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/math';

/** Inspect only the selected document body, never style definitions or package metadata. */
export function wordSelectionWarning(ooxml: string): string | undefined {
  const xml = new DOMParser().parseFromString(ooxml, 'application/xml');
  const body = xml.getElementsByTagNameNS(WORD_NS, 'body')[0];
  if (xml.getElementsByTagName('parsererror').length || !body)
    return '无法确认选区结构。可生成和插入结果，请重新选择纯文字段落后再替换。';
  const protectedElements = [
    'tbl',
    'drawing',
    'pict',
    'object',
    'sdt',
    'fldChar',
    'fldSimple',
    'hyperlink',
    'footnoteReference',
    'endnoteReference',
  ];
  if (
    protectedElements.some((name) => body.getElementsByTagNameNS(WORD_NS, name).length) ||
    body.getElementsByTagNameNS(MATH_NS, 'oMath').length ||
    body.getElementsByTagNameNS(MATH_NS, 'oMathPara').length
  )
    return '选区含表格、图片、公式或其他文档对象。为保留结构，本次仅支持插入或复制；如需替换，请只选中纯文字段落。';
}
