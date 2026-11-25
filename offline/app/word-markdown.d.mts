/** Safe, self-contained HTML fragment for Word; null means keep the exact plain text. */
export function wordMarkdownHtml(text: string): string | null;
/** block=false is inline HTML safe to replace text inside a paragraph. */
export function wordMarkdownFormat(text: string): { html: string; block: boolean } | null;
