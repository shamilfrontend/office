import { defineStore } from 'pinia';
import { ref, watch } from 'vue';
import type { DocItem } from '@/api/client';
import { useAuthStore } from '@/stores/auth';

type FileTab = {
  id: string;
  name: string;
  type: DocItem['type'];
};

type Persisted = {
  tabs: FileTab[];
  activeId: string;
};

function storageKey(userId: string) {
  return `office_file_tabs:${userId}`;
}

export const useFileTabsStore = defineStore('fileTabs', () => {
  const tabs = ref<FileTab[]>([]);
  const activeId = ref('');
  let hydratedUserId: string | null = null;

  function persist() {
    const auth = useAuthStore();
    const userId = auth.me?.id;
    if (!userId) return;
    const payload: Persisted = {
      tabs: tabs.value,
      activeId: activeId.value,
    };
    sessionStorage.setItem(storageKey(userId), JSON.stringify(payload));
  }

  function hydrateFromSession() {
    const auth = useAuthStore();
    const userId = auth.me?.id;
    if (!userId) return;
    if (hydratedUserId === userId && tabs.value.length) return;
    hydratedUserId = userId;
    try {
      const raw = sessionStorage.getItem(storageKey(userId));
      if (!raw) return;
      const data = JSON.parse(raw) as Persisted;
      if (!Array.isArray(data.tabs)) return;
      tabs.value = data.tabs.filter(
        (t) => t && typeof t.id === 'string' && typeof t.name === 'string',
      );
      activeId.value =
        data.activeId && tabs.value.some((t) => t.id === data.activeId)
          ? data.activeId
          : tabs.value[0]?.id || '';
    } catch {
      /* ignore */
    }
  }

  function ensureHydrated() {
    hydrateFromSession();
  }

  function open(doc: Pick<DocItem, 'id' | 'name' | 'type'>) {
    ensureHydrated();
    const existing = tabs.value.find((t) => t.id === doc.id);
    if (existing) {
      existing.name = doc.name;
      existing.type = doc.type;
    } else {
      tabs.value.push({ id: doc.id, name: doc.name, type: doc.type });
    }
    activeId.value = doc.id;
    persist();
  }

  function focus(id: string) {
    if (!tabs.value.some((t) => t.id === id)) return false;
    activeId.value = id;
    persist();
    return true;
  }

  function close(id: string): string | null {
    const idx = tabs.value.findIndex((t) => t.id === id);
    if (idx < 0) return activeId.value || null;
    tabs.value.splice(idx, 1);
    if (!tabs.value.length) {
      activeId.value = '';
      persist();
      return null;
    }
    if (activeId.value === id) {
      const next = tabs.value[Math.min(idx, tabs.value.length - 1)];
      activeId.value = next.id;
    }
    persist();
    return activeId.value;
  }

  function rename(id: string, name: string) {
    const tab = tabs.value.find((t) => t.id === id);
    if (!tab) return;
    tab.name = name;
    persist();
  }

  function syncActive(id: string, meta?: Pick<DocItem, 'name' | 'type'>) {
    ensureHydrated();
    const existing = tabs.value.find((t) => t.id === id);
    if (existing) {
      if (meta) {
        existing.name = meta.name;
        existing.type = meta.type;
      }
    } else if (meta) {
      tabs.value.push({ id, name: meta.name, type: meta.type });
    } else {
      return;
    }
    activeId.value = id;
    persist();
  }

  function remove(id: string) {
    close(id);
  }

  watch(
    () => useAuthStore().me?.id,
    (userId, prev) => {
      if (userId !== prev) {
        tabs.value = [];
        activeId.value = '';
        hydratedUserId = null;
        if (userId) hydrateFromSession();
      }
    },
  );

  return {
    tabs,
    hydrateFromSession,
    open,
    focus,
    close,
    rename,
    syncActive,
    remove,
  };
});
