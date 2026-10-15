import { ref, type ComputedRef, type Ref } from 'vue';
import { api } from '@/api/client';
import type { ChatMode } from '@/univer/officeTools';
import type { ModelOption, Tab } from '../types';

type ProviderInfo = {
  id: string;
  provider: string;
  model?: string | null;
};

export function useModelMenu(options: {
  active: ComputedRef<Tab | null>;
  providers: ComputedRef<ProviderInfo[]>;
  patchTab: (tab: Tab, body: Record<string, unknown>) => Promise<void>;
  modeMenu: Ref<boolean>;
  modelMenu: Ref<boolean>;
  historyOpen: Ref<boolean>;
}) {
  const { active, providers, patchTab, modeMenu, modelMenu, historyOpen } =
    options;

  const modelOptions = ref<ModelOption[]>([]);
  const modelMenuLoading = ref(false);
  const modelMenuError = ref('');

  function setMode(mode: ChatMode) {
    const tab = active.value;
    modeMenu.value = false;
    if (!tab || tab.mode === mode) return;
    tab.mode = mode;
    void patchTab(tab, { mode });
  }

  async function loadModelOptions() {
    modelMenuLoading.value = true;
    modelMenuError.value = '';
    try {
      const errors: string[] = [];
      const rows = await Promise.all(
        providers.value.map(async (provider) => {
          try {
            const res = await api.listProviderModels(provider.id);
            return res.models.map((model) => ({
              providerId: provider.id,
              provider: provider.provider,
              model,
            }));
          } catch (e) {
            const msg =
              e instanceof Error ? e.message : 'Ошибка загрузки моделей';
            errors.push(
              providers.value.length > 1 ? `${provider.provider}: ${msg}` : msg,
            );
            return [] as ModelOption[];
          }
        }),
      );
      modelOptions.value = rows.flat();
      if (!modelOptions.value.length) {
        modelMenuError.value = errors[0] || 'Не удалось загрузить модели';
      }
    } finally {
      modelMenuLoading.value = false;
    }
  }

  async function toggleModelMenu() {
    modeMenu.value = false;
    historyOpen.value = false;
    modelMenu.value = !modelMenu.value;
    if (!modelMenu.value) return;
    await loadModelOptions();
  }

  function setModel(option: ModelOption) {
    const tab = active.value;
    modelMenu.value = false;
    if (!tab) return;
    tab.providerId = option.providerId;
    tab.model = option.model;
    void patchTab(tab, { providerId: option.providerId, model: option.model });
  }

  async function ensureDefaultModel() {
    const tab = active.value;
    if (!tab || tab.model || !providers.value.length) return;
    if (!modelOptions.value.length) await loadModelOptions();
    const first = modelOptions.value[0];
    if (first) setModel(first);
  }

  function modelLabel(tab: Tab) {
    if (tab.model) return tab.model;
    return 'модель';
  }

  function resetModelMenu() {
    modelMenu.value = false;
    modelOptions.value = [];
    modelMenuLoading.value = false;
    modelMenuError.value = '';
  }

  return {
    modelOptions,
    modelMenuLoading,
    modelMenuError,
    setMode,
    toggleModelMenu,
    loadModelOptions,
    setModel,
    ensureDefaultModel,
    modelLabel,
    resetModelMenu,
  };
}
