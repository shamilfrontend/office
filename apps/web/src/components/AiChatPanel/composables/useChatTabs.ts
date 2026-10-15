import { computed, nextTick, ref, type ComputedRef, type Ref } from 'vue';
import { api } from '@/api/client';
import { tabFromRecord } from '../tabMappers';
import type { Tab } from '../types';

type ProviderInfo = { id: string; model?: string | null };

export function useChatTabs(options: {
  documentId: () => string | undefined;
  hasProvider: ComputedRef<boolean>;
  providers: ComputedRef<ProviderInfo[]>;
  attachError: Ref<string>;
  ensureDefaultModel: () => Promise<void>;
  loadHistoryIfOpen: () => void;
}) {
  const {
    documentId,
    hasProvider,
    providers,
    attachError,
    ensureDefaultModel,
    loadHistoryIfOpen,
  } = options;

  const tabs = ref<Tab[]>([]);
  const activeId = ref('');
  const renamingId = ref('');
  const renameDraft = ref('');

  const active = computed(
    () => tabs.value.find((t) => t.id === activeId.value) ?? null,
  );

  function findTab(id: string) {
    return tabs.value.find((tab) => tab.id === id);
  }

  async function patchTab(tab: Tab, body: Record<string, unknown>) {
    try {
      await api.updateChat(tab.id, body);
    } catch (e) {
      console.error(e);
    }
  }

  async function loadTabs() {
    const forDoc = documentId();
    if (!hasProvider.value || !forDoc) return;
    try {
      const rows = await api.listOpenChats(forDoc);
      if (documentId() !== forDoc) return;
      tabs.value = rows.map((record) => {
        const tab = tabFromRecord(record, providers.value);
        if (record.planExecuted == null && tab.planRunState === 'done') {
          void patchTab(tab, { planExecuted: true });
        }
        return tab;
      });
      activeId.value = tabs.value[0]?.id || '';
      await ensureDefaultModel();
    } catch (e) {
      if (documentId() !== forDoc) return;
      attachError.value =
        e instanceof Error ? e.message : 'Не удалось открыть чаты';
    }
  }

  async function addTab() {
    const docId = documentId();
    if (!docId) return;
    const created = await api.createChat({
      mode: active.value?.mode || 'agent',
      providerId: active.value?.providerId || providers.value[0]?.id,
      model: active.value?.model || undefined,
      documentId: docId,
    });
    tabs.value.push(tabFromRecord(created, providers.value));
    activeId.value = created.id;
    await ensureDefaultModel();
  }

  async function closeTab(id: string) {
    const tab = findTab(id);
    if (!tab) return;
    tab.abort?.abort();
    tab.pendingApprove?.resolve(false);
    if (renamingId.value === id) {
      renamingId.value = '';
      renameDraft.value = '';
    }
    try {
      const result = await api.closeChat(id);
      tabs.value = tabs.value.filter((item) => item.id !== id);
      if (result.chat)
        tabs.value.push(tabFromRecord(result.chat, providers.value));
      if (!tabs.value.some((item) => item.id === activeId.value)) {
        activeId.value = tabs.value[0]?.id || '';
      }
      loadHistoryIfOpen();
      await ensureDefaultModel();
    } catch (e) {
      attachError.value =
        e instanceof Error ? e.message : 'Не удалось закрыть чат';
    }
  }

  function startRename(tab: { id: string; title: string }) {
    renamingId.value = tab.id;
    renameDraft.value = tab.title;
    void nextTick(() => {
      const input = document.querySelector(
        `.tab-rename[data-id="${tab.id}"]`,
      ) as HTMLInputElement | null;
      input?.focus();
      input?.select();
    });
  }

  function cancelRename() {
    renamingId.value = '';
    renameDraft.value = '';
  }

  async function commitRename(tab: { id: string }) {
    if (renamingId.value !== tab.id) return;
    const next = renameDraft.value.replace(/\s+/g, ' ').trim() || 'Новый чат';
    renamingId.value = '';
    renameDraft.value = '';
    const target = findTab(tab.id);
    if (!target || next === target.title) return;
    target.title = next;
    await patchTab(target, { title: next });
  }

  function onRenameKey(event: KeyboardEvent, tab: { id: string }) {
    if (event.key === 'Enter') {
      event.preventDefault();
      void commitRename(tab);
    }
    if (event.key === 'Escape') {
      event.preventDefault();
      cancelRename();
    }
  }

  function resetTabsState() {
    tabs.value.forEach((tab) => {
      tab.abort?.abort();
      tab.pendingApprove?.resolve(false);
    });
    tabs.value = [];
    activeId.value = '';
    renamingId.value = '';
    renameDraft.value = '';
  }

  return {
    tabs,
    activeId,
    active,
    renamingId,
    renameDraft,
    findTab,
    patchTab,
    loadTabs,
    addTab,
    closeTab,
    startRename,
    cancelRename,
    commitRename,
    onRenameKey,
    resetTabsState,
  };
}
