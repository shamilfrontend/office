import {
  computed,
  nextTick,
  ref,
  watch,
  type ComputedRef,
  type Ref,
} from 'vue';
import type { Tab, UiMsg } from '../types';

export function useTranscriptUi(options: {
  active: ComputedRef<Tab | null>;
  docType: () => 'SHEET' | 'DOC' | 'SLIDE' | null | undefined;
  attachError: Ref<string>;
  transcriptEl: () => HTMLElement | null | undefined;
}) {
  const { active, docType, attachError, transcriptEl } = options;

  const expandedTools = ref<Set<string>>(new Set());
  const copiedKey = ref('');
  let copiedTimer: ReturnType<typeof setTimeout> | null = null;

  const composerPlaceholder = computed(() => {
    const mode = active.value?.mode;
    if (mode === 'plan') return 'Опишите задачу для плана…';
    if (mode === 'chat') return 'Спросите что угодно…';
    return 'Что сделать с документом?';
  });

  const emptyHints = computed(() => {
    const mode = active.value?.mode || 'agent';
    const kind = docType();
    if (kind === 'SLIDE') {
      if (mode === 'plan') {
        return [
          'Составь план презентации на 5 слайдов',
          'Разбей задачу на шаги без правок файла',
          'Какие слайды стоит сократить?',
        ];
      }
      if (mode === 'chat') {
        return [
          'О чём эта презентация?',
          'Что сейчас выделено?',
          'Перескажи заголовки слайдов',
        ];
      }
      return [
        'Смени заголовок на первом слайде',
        'Добавь слайд с тремя тезисами',
        'Сделай выделенный текст короче',
      ];
    }
    const isDoc = kind === 'DOC';
    if (mode === 'plan') {
      return isDoc
        ? [
            'Составь план правки выделенного фрагмента',
            'Разбей задачу на шаги без правок файла',
            'Какие риски у глобальной замены текста?',
          ]
        : [
            'Составь план очистки и нормализации таблицы',
            'Разбей задачу на шаги без правок файла',
            'Какие риски у массовой замены значений?',
          ];
    }
    if (mode === 'chat') {
      return isDoc
        ? [
            'Объясни структуру текущего документа',
            'Что сейчас выделено?',
            'Кратко перескажи начало текста',
          ]
        : [
            'Объясни структуру текущего файла',
            'Что в текущем выделении?',
            'Найди пустые ячейки на активном листе',
          ];
    }
    return isDoc
      ? [
          'Сделай выделенный заголовок жирным',
          'Вставь абзац в место курсора',
          'Замени выделенный текст на более короткий',
        ]
      : [
          'Сделай первую строку жирной и подгони ширину колонок',
          'Объедини A1:C1 и добавь внешнюю рамку',
          'Отформатируй заголовки: жирный, заливка, по центру',
        ];
  });

  const approveSummary = computed(() => {
    const pending = active.value?.pendingApprove;
    if (!pending) return '';
    const { title, detail } = pending.summary;
    return detail ? `${title}: ${detail}` : title;
  });

  watch(
    () => active.value?.messages.length,
    async () => {
      await nextTick();
      const el = transcriptEl();
      if (el) el.scrollTop = el.scrollHeight;
    },
  );

  async function copyText(text: string, key: string) {
    try {
      await navigator.clipboard.writeText(text);
      copiedKey.value = key;
      if (copiedTimer) clearTimeout(copiedTimer);
      copiedTimer = setTimeout(() => {
        copiedKey.value = '';
        copiedTimer = null;
      }, 1500);
    } catch {
      attachError.value = 'Не удалось скопировать';
    }
  }

  function onTranscriptClick(event: MouseEvent) {
    const target = event.target as HTMLElement | null;
    const btn = target?.closest?.('.copy-code') as HTMLElement | null;
    if (!btn) return;
    const pre = btn.parentElement?.querySelector('code');
    const text = pre?.textContent || '';
    void copyText(text, `code:${text.slice(0, 24)}`);
    btn.textContent = 'Скопировано';
    setTimeout(() => {
      btn.textContent = 'Копировать';
    }, 1200);
  }

  function isToolOpen(message: Extract<UiMsg, { role: 'tool' }>) {
    if (message.status === 'error' || message.status === 'pending') return true;
    return expandedTools.value.has(message.key);
  }

  function toggleTool(key: string) {
    const next = new Set(expandedTools.value);
    if (next.has(key)) next.delete(key);
    else next.add(key);
    expandedTools.value = next;
  }

  function applyHint(text: string) {
    const tab = active.value;
    if (!tab || tab.streaming) return;
    tab.draft = text;
  }

  function resetTranscriptUi() {
    expandedTools.value = new Set();
    copiedKey.value = '';
    if (copiedTimer) {
      clearTimeout(copiedTimer);
      copiedTimer = null;
    }
  }

  return {
    expandedTools,
    copiedKey,
    composerPlaceholder,
    emptyHints,
    approveSummary,
    copyText,
    onTranscriptClick,
    isToolOpen,
    toggleTool,
    applyHint,
    resetTranscriptUi,
  };
}
