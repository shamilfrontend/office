import { ref, type Ref } from 'vue';
import { api, type ChatSummary } from '@/api/client';
import { uiConfirm } from '@/composables/useConfirm';
import { tabFromRecord } from '../tabMappers';
import type { Tab } from '../types';

type ProviderInfo = { id: string; model?: string | null };

export function useChatHistory(options: {
  documentId: () => string | undefined;
  tabs: Ref<Tab[]>;
  activeId: Ref<string>;
  providers: () => ProviderInfo[];
  ensureDefaultModel: () => Promise<void>;
  modeMenu: Ref<boolean>;
  modelMenu: Ref<boolean>;
  historyOpen: Ref<boolean>;
}) {
  const {
    documentId,
    tabs,
    activeId,
    providers,
    ensureDefaultModel,
    modeMenu,
    modelMenu,
    historyOpen,
  } = options;

  const historyItems = ref<ChatSummary[]>([]);
  const historyLoading = ref(false);
  const historyError = ref('');

  async function loadHistory() {
    const id = documentId();
    if (!id) return;
    historyLoading.value = true;
    historyError.value = '';
    try {
      historyItems.value = await api.listClosedChats(id);
    } catch (e) {
      historyError.value =
        e instanceof Error ? e.message : 'Не удалось загрузить историю';
    } finally {
      historyLoading.value = false;
    }
  }

  async function toggleHistory() {
    historyOpen.value = !historyOpen.value;
    modeMenu.value = false;
    modelMenu.value = false;
    if (historyOpen.value) await loadHistory();
  }

  async function reopenChat(id: string) {
    try {
      const record = await api.updateChat(id, { pinnedOpen: true });
      if (!tabs.value.some((tab) => tab.id === record.id)) {
        tabs.value.push(tabFromRecord(record, providers()));
      }
      activeId.value = record.id;
      historyOpen.value = false;
      historyItems.value = historyItems.value.filter((item) => item.id !== id);
      await ensureDefaultModel();
    } catch (e) {
      historyError.value =
        e instanceof Error ? e.message : 'Не удалось открыть чат';
    }
  }

  async function deleteFromHistory(id: string) {
    const item = historyItems.value.find((row) => row.id === id);
    const ok = await uiConfirm({
      title: 'Удалить чат?',
      message: item?.title
        ? `Чат «${item.title}» будет удалён безвозвратно.`
        : 'Чат будет удалён безвозвратно.',
      confirmLabel: 'Удалить',
    });
    if (!ok) return;
    try {
      await api.deleteChat(id);
      historyItems.value = historyItems.value.filter((row) => row.id !== id);
    } catch (e) {
      historyError.value =
        e instanceof Error ? e.message : 'Не удалось удалить чат';
    }
  }

  function resetHistory() {
    historyOpen.value = false;
    historyItems.value = [];
    historyError.value = '';
    historyLoading.value = false;
  }

  return {
    historyItems,
    historyLoading,
    historyError,
    loadHistory,
    toggleHistory,
    reopenChat,
    deleteFromHistory,
    resetHistory,
  };
}
