import { ref } from 'vue';
import { renderMd } from '../markdown';
import type { UiMsg } from '../types';

const STREAM_HTML_MS = 80;

export function useStreamHtml() {
  const streamHtmlKey = ref('');
  const streamHtml = ref('');
  let streamHtmlTimer: ReturnType<typeof setTimeout> | null = null;

  function clearStreamHtml() {
    if (streamHtmlTimer) {
      clearTimeout(streamHtmlTimer);
      streamHtmlTimer = null;
    }
    streamHtmlKey.value = '';
    streamHtml.value = '';
  }

  function flushStreamHtml(key: string, content: string) {
    streamHtmlKey.value = key;
    streamHtml.value = renderMd(content);
  }

  function scheduleStreamHtml(key: string, content: string) {
    if (!streamHtmlTimer) {
      streamHtmlTimer = setTimeout(() => {
        streamHtmlTimer = null;
        flushStreamHtml(key, content);
      }, STREAM_HTML_MS);
      return;
    }
    const pendingKey = key;
    const pendingContent = content;
    clearTimeout(streamHtmlTimer);
    streamHtmlTimer = setTimeout(() => {
      streamHtmlTimer = null;
      flushStreamHtml(pendingKey, pendingContent);
    }, STREAM_HTML_MS);
  }

  function assistantHtml(message: Extract<UiMsg, { role: 'assistant' }>) {
    if (message.key === streamHtmlKey.value && streamHtml.value) {
      return streamHtml.value;
    }
    return renderMd(message.content);
  }

  return {
    streamHtmlKey,
    streamHtml,
    clearStreamHtml,
    flushStreamHtml,
    scheduleStreamHtml,
    assistantHtml,
  };
}
