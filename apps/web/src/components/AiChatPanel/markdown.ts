import DOMPurify from 'dompurify';
import { marked } from 'marked';

export function renderMd(text: string) {
  const raw = marked.parse(text || '', { async: false }) as string;
  const withCopy = raw.replace(
    /<pre><code([^>]*)>([\s\S]*?)<\/code><\/pre>/g,
    (_match, attrs: string, code: string) =>
      `<div class="code-block"><button type="button" class="copy-code" data-copy="1">Копировать</button><pre><code${attrs}>${code}</code></pre></div>`,
  );
  return DOMPurify.sanitize(withCopy, {
    ADD_ATTR: ['data-copy'],
    ADD_TAGS: ['button'],
  });
}
