import { useMemo, type HTMLAttributes } from 'react';
import { wordMarkdownHtml } from '../../offline/app/word-markdown.mjs';

export function WordTextPreview({
  text,
  className = '',
  ...props
}: { text: string } & HTMLAttributes<HTMLElement>) {
  const html = useMemo(() => wordMarkdownHtml(text), [text]);
  if (!html)
    return (
      <pre className={className} {...props}>
        {text}
      </pre>
    );
  // Only the Word converter's allowlisted markup enters this preview.
  return (
    <div
      {...props}
      className={`word-rich-text ${className}`}
      dangerouslySetInnerHTML={{ __html: html }}
    />
  );
}
