import type { ComputedRef, Ref } from 'vue';
import { api } from '@/api/client';
import type { Tab } from '../types';

export function useComposerAttachments(options: {
  active: ComputedRef<Tab | null>;
  attachError: Ref<string>;
  fileInput: () => HTMLInputElement | null | undefined;
}) {
  const { active, attachError, fileInput } = options;

  function pickAttach() {
    fileInput()?.click();
  }

  async function onAttach(e: Event) {
    const inputEl = e.target as HTMLInputElement;
    const file = inputEl.files?.[0];
    inputEl.value = '';
    const tab = active.value;
    if (!file || !tab) return;
    attachError.value = '';
    try {
      const extracted = await api.extractAttachment(file);
      tab.attachments.push(extracted);
    } catch (err) {
      attachError.value =
        err instanceof Error ? err.message : 'Не удалось прочитать файл';
    }
  }

  function removeAttach(index: number) {
    active.value?.attachments.splice(index, 1);
  }

  return {
    pickAttach,
    onAttach,
    removeAttach,
  };
}
