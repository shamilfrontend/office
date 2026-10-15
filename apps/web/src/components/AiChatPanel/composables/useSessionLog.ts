import { computed, ref, type Ref } from 'vue';
import type { SessionLogEntry, Tab } from '../types';

export function useSessionLog(options: {
  documentId: () => string | undefined;
  docName: () => string | undefined;
  docType: () => string | null | undefined;
  attachError: Ref<string>;
}) {
  const { documentId, docName, docType, attachError } = options;

  const sessionLog = ref<SessionLogEntry[]>([]);
  const logCopied = ref(false);
  let logCopiedTimer: ReturnType<typeof setTimeout> | null = null;

  const hasSessionLog = computed(() => sessionLog.value.length > 0);

  function pushSessionLog(
    tab: Tab | null | undefined,
    entry: Omit<SessionLogEntry, 'ts' | 'chatId' | 'chatTitle'> & {
      chatId?: string;
      chatTitle?: string;
    },
  ) {
    sessionLog.value.push({
      ts: new Date().toISOString(),
      chatId: entry.chatId ?? tab?.id ?? '',
      chatTitle: entry.chatTitle ?? tab?.title ?? '',
      kind: entry.kind,
      name: entry.name,
      content: entry.content,
      status: entry.status,
    });
  }

  function formatSessionLog() {
    const header = [
      '=== Office session log ===',
      `exportedAt: ${new Date().toISOString()}`,
      `documentId: ${documentId() || '-'}`,
      `docName: ${docName() || '-'}`,
      `docType: ${docType() || '-'}`,
      `entries: ${sessionLog.value.length}`,
      '',
    ].join('\n');
    const body = sessionLog.value
      .map((entry) => {
        const meta = [
          entry.ts,
          entry.kind,
          entry.name || '',
          entry.status || '',
          `chat=${entry.chatTitle || entry.chatId || '-'}`,
        ]
          .filter(Boolean)
          .join(' | ');
        return `[${meta}]\n${entry.content}`;
      })
      .join('\n\n---\n\n');
    return `${header}${body}`;
  }

  async function copySessionLog() {
    if (!sessionLog.value.length) return;
    try {
      await navigator.clipboard.writeText(formatSessionLog());
      logCopied.value = true;
      if (logCopiedTimer) clearTimeout(logCopiedTimer);
      logCopiedTimer = setTimeout(() => {
        logCopied.value = false;
        logCopiedTimer = null;
      }, 1500);
    } catch {
      attachError.value = 'Не удалось скопировать лог';
    }
  }

  function resetSessionLog() {
    sessionLog.value = [];
    logCopied.value = false;
    if (logCopiedTimer) {
      clearTimeout(logCopiedTimer);
      logCopiedTimer = null;
    }
  }

  return {
    sessionLog,
    logCopied,
    hasSessionLog,
    pushSessionLog,
    copySessionLog,
    resetSessionLog,
  };
}
