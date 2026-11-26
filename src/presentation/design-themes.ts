import { t } from '../i18n';
import { json, object, text } from './validation';
import type { PresentationTheme } from './types';

export interface DeckTheme extends PresentationTheme {
  name: string;
  headingFont: string;
  secondary: string;
}
export const themeFonts = ['Aptos', 'Arial', 'Georgia', '等线', '宋体'];
const rgb = (c: string) => [1, 3, 5].map((i) => parseInt(c.slice(i, i + 2), 16));
export function contrast(a: string, b: string) {
  const lum = (c: string) =>
    rgb(c)
      .map((v) => v / 255)
      .map((v) => (v <= 0.04045 ? v / 12.92 : ((v + 0.055) / 1.055) ** 2.4))
      .reduce((s, v, i) => s + v * [0.2126, 0.7152, 0.0722][i], 0);
  const l = lum(a),
    r = lum(b);
  return (Math.max(l, r) + 0.05) / (Math.min(l, r) + 0.05);
}
export function readableOn(background: string, preferred: string) {
  if (contrast(background, preferred) >= 4.5) return preferred;
  return contrast(background, '#FFFFFF') > contrast(background, '#111111') ? '#FFFFFF' : '#111111';
}
export function chartPalette(theme?: PresentationTheme) {
  return [
    theme?.accent || '#28604A',
    theme?.secondary || '#6F94A8',
    '#C59B53',
    '#937A9B',
    '#9AAB75',
    '#C88176',
  ];
}
export function builtInThemes(brand?: string): DeckTheme[] {
  return [
    {
      name: t('蓝色简报', 'Blue briefing'),
      headingFont: 'Arial',
      font: 'Aptos',
      background: '#EEF4FA',
      ink: '#172F49',
      accent: '#245887',
      secondary: '#728DA2',
    },
    {
      name: t('夜间讲堂', 'After dark'),
      headingFont: '等线',
      font: 'Arial',
      background: '#18262D',
      ink: '#F2F5ED',
      accent: '#93CEC0',
      secondary: '#D9AD74',
    },
    {
      name: t('纸本阅读', 'Paper and ink'),
      headingFont: '宋体',
      font: 'Georgia',
      background: '#FAF3E6',
      ink: '#372A23',
      accent: '#8C4636',
      secondary: '#727C61',
    },
  ].map((v) => ({ ...v, accent: brand || v.accent }));
}
export function validatedTheme(value: unknown): DeckTheme {
  const v = object(value, t('主题', 'Theme'));
  const color = (key: string) => {
    if (typeof v[key] !== 'string' || !/^#[a-f\d]{6}$/i.test(v[key] as string))
      throw Error(t('主题颜色需要六位 HEX 色值。', 'Use six-digit HEX theme colors.'));
    return (v[key] as string).toUpperCase();
  };
  const font = (key: string) => {
    if (!themeFonts.includes(v[key] as string))
      throw Error(t('主题使用了不支持的字体。', 'The theme uses an unsupported font.'));
    return v[key] as string;
  };
  const background = color('background');
  return {
    name: text(v.name, 30, t('主题名称', 'Theme name')),
    font: font('font'),
    headingFont: font('headingFont'),
    accent: color('accent'),
    secondary: color('secondary'),
    background,
    ink: readableOn(background, color('ink')),
  };
}
export function parseThemes(raw: string, brand?: string) {
  const result = json(raw).themes;
  if (!Array.isArray(result) || result.length !== 3)
    throw Error(t('请重新生成三套主题。', 'Generate three themes again.'));
  const themes = result.map(validatedTheme).map((s) => ({ ...s, accent: brand || s.accent }));
  if (
    new Set(themes.map((s) => s.name)).size !== 3 ||
    new Set(
      themes.map((s) =>
        JSON.stringify([s.font, s.headingFont, s.accent, s.secondary, s.ink, s.background]),
      ),
    ).size !== 3
  )
    throw Error(t('主题过于相似，请重新生成。', 'These themes are too similar. Generate again.'));
  return themes;
}
