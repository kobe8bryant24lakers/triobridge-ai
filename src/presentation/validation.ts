export function object(value: unknown, label: string): Record<string, unknown> {
  if (!value || typeof value !== 'object' || Array.isArray(value))
    throw new Error(`${label}格式无效。`);
  return value as Record<string, unknown>;
}
export function fields(value: Record<string, unknown>, names: string[], label: string) {
  if (Object.keys(value).some((k) => !names.includes(k)))
    throw new Error(`${label}含不支持的字段，请重新生成。`);
}
export function text(value: unknown, max: number, label: string, empty = false): string {
  if (
    typeof value !== 'string' ||
    value.length > max ||
    (!empty && !value.trim()) ||
    /[\u0000-\u0008\u000B\u000C\u000E-\u001F]/.test(value)
  )
    throw new Error(`${label}需要${empty ? '不超过' : '1–'}${max}字符的文字。`);
  return value;
}
export function list(value: unknown, max: number, label: string): unknown[] {
  if (!Array.isArray(value) || value.length > max) throw new Error(`${label}最多 ${max} 项。`);
  return value;
}
export function json(raw: string): Record<string, unknown> {
  if (raw.length > 100_000) throw new Error('结果过长，请缩小范围后重新生成。');
  try {
    return object(
      JSON.parse(raw.trim().replace(/^```(?:json)?\s*\n([\s\S]*?)\n```$/i, '$1')),
      '结果',
    );
  } catch {
    throw new Error('模型未返回有效的 JSON，请重新生成或更换模型。');
  }
}
