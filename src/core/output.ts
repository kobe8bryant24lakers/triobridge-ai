import type { SlideContent } from './types';

function unwrapFence(value: string): string {
  return value
    .trim()
    .replace(/^```(?:json|excel|formula)?\s*\n([\s\S]*?)\n```$/i, '$1')
    .trim();
}

export function parseFormula(raw: string): string {
  const formula = unwrapFence(raw);
  if (
    !formula.startsWith('=') ||
    /[\r\n\t]/.test(formula) ||
    formula.length > 8192 ||
    formula.length < 2
  ) {
    throw new Error('结果不是有效的单行 Excel 公式。请编辑结果为以 = 开头的公式后再写入。');
  }
  // This is a structural preflight, not an Excel calculation engine. Quoted strings,
  // sheet names and structured references can contain literal parentheses.
  let code = '';
  let quote = '';
  let brackets = 0;
  let parentheses = 0;
  for (let i = 0; i < formula.length; i += 1) {
    const character = formula[i];
    if (quote) {
      if (character === quote) {
        if (formula[i + 1] === quote) i += 1;
        else quote = '';
      }
      code += ' ';
      continue;
    }
    if (!brackets && (character === '"' || character === "'")) {
      quote = character;
      code += ' ';
      continue;
    }
    // Apostrophe escapes special characters inside Excel table column names.
    if (brackets && character === "'") {
      i += 1;
      continue;
    }
    if (character === '[') brackets += 1;
    if (character === ']') brackets -= 1;
    if (brackets < 0) throw new Error('公式的方括号不匹配，请检查表格列引用。');
    if (!brackets) {
      if (character === '(') parentheses += 1;
      if (character === ')') parentheses -= 1;
      if (parentheses < 0) throw new Error('公式的括号不匹配，请检查后再写入。');
      code += character;
    }
  }
  if (quote || brackets || parentheses) throw new Error('公式的引号或括号未闭合，请检查后再写入。');
  if (
    /\b(?:WEBSERVICE|HYPERLINK|RTD|CALL|REGISTER(?:\.ID)?|EXEC)\s*\(/i.test(code) ||
    /[|]/.test(code) ||
    /\[[^\]]+\][^\[\]+*/(),=<>:&^%\-]*!|https?:|file:|\\\\/i.test(formula)
  ) {
    throw new Error('公式含外部链接、工作簿引用或外部调用，请移除后重试。');
  }
  return formula;
}

export function parseOutline(raw: string): SlideContent[] {
  let data: unknown;
  try {
    data = JSON.parse(unwrapFence(raw));
  } catch {
    throw new Error(
      '大纲格式无效，需要 {"slides":[{"title":"标题","bullets":["要点"]}]}。可编辑结果或重新生成。',
    );
  }
  const slides = data && typeof data === 'object' && 'slides' in data ? data.slides : undefined;
  if (data && typeof data === 'object' && Object.keys(data).some((key) => key !== 'slides'))
    throw new Error('大纲包含暂不支持的字段，请仅保留 slides，并将文字放入页面标题或要点。');
  if (!Array.isArray(slides) || !slides.length || slides.length > 12)
    throw new Error('大纲需要包含 1–12 页。');
  return slides.map((slide: unknown, index: number) => {
    if (
      slide &&
      typeof slide === 'object' &&
      Object.keys(slide).some((key) => !['title', 'bullets'].includes(key))
    )
      throw new Error(`第 ${index + 1} 页包含暂不支持的字段，请将内容移入标题或要点。`);
    if (
      !slide ||
      typeof slide !== 'object' ||
      !('title' in slide) ||
      !('bullets' in slide) ||
      typeof slide.title !== 'string' ||
      !slide.title.trim() ||
      slide.title.length > 60 ||
      !Array.isArray(slide.bullets) ||
      !slide.bullets.length ||
      slide.bullets.length > 5 ||
      !slide.bullets.every(
        (item: unknown) => typeof item === 'string' && item.trim().length > 0 && item.length <= 160,
      )
    ) {
      throw new Error(
        `第 ${index + 1} 页格式有误：需要一个不超过 60 字符的标题，以及 1–5 条不超过 160 字符的要点。`,
      );
    }
    return {
      title: slide.title.trim(),
      bullets: slide.bullets.map((item: string) => item.trim()),
    };
  });
}

export function outlineToText(raw: string): string {
  return parseOutline(raw)
    .map((slide, i) => `${i + 1}. ${slide.title}\n${slide.bullets.map((b) => `• ${b}`).join('\n')}`)
    .join('\n\n');
}
