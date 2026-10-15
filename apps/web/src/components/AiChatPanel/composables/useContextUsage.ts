import { ref, watch, type ComputedRef } from 'vue';
import {
  estimateChatContextUsage,
  type ChatContextUsage,
} from '@/ai/estimateChatContextUsage';
import { chatTokenCalibrator } from '@/ai/estimateTokens';
import { toWire, trimWireForProvider } from '../chatWire';
import type { Tab } from '../types';

const CONTEXT_USAGE_MS = 400;

export function useContextUsage(options: {
  active: ComputedRef<Tab | null>;
  docName: () => string | undefined;
  docType: () => 'SHEET' | 'DOC' | null | undefined;
  documentId: () => string | undefined;
  getSelectionContext: () => (() => string) | undefined;
}) {
  const { active, docName, docType, documentId, getSelectionContext } = options;

  const contextUsage = ref<ChatContextUsage | null>(null);
  let contextUsageTimer: ReturnType<typeof setTimeout> | null = null;

  function refreshContextUsage() {
    const tab = active.value;
    if (!tab || !tab.model) {
      contextUsage.value = null;
      return;
    }
    const hasDocument = Boolean(documentId() || docName());
    const wireMessages = trimWireForProvider(toWire(tab));
    contextUsage.value = estimateChatContextUsage({
      mode: tab.mode,
      model: tab.model,
      wireMessages,
      draft: tab.draft,
      attachments: tab.attachments,
      includeDoc: tab.includeDoc,
      docName: docName(),
      docType: docType(),
      selectionContext: tab.includeDoc ? getSelectionContext()?.() : undefined,
      buildingPlan: tab.executingPlan,
      hasDocument,
      calibrator: chatTokenCalibrator,
    });
  }

  watch(
    () => {
      const tab = active.value;
      if (!tab) return null;
      return [
        tab.id,
        tab.mode,
        tab.model,
        tab.draft,
        tab.includeDoc,
        tab.executingPlan,
        tab.attachments.length,
        tab.attachments.reduce((n, f) => n + f.text.length, 0),
        tab.messages.length,
        tab.messages.map((m) => m.key).join(','),
        docName(),
        docType(),
        documentId(),
      ] as const;
    },
    () => {
      if (contextUsageTimer) clearTimeout(contextUsageTimer);
      contextUsageTimer = setTimeout(() => {
        contextUsageTimer = null;
        refreshContextUsage();
      }, CONTEXT_USAGE_MS);
    },
    { immediate: true },
  );

  function resetContextUsage() {
    contextUsage.value = null;
    if (contextUsageTimer) {
      clearTimeout(contextUsageTimer);
      contextUsageTimer = null;
    }
  }

  return {
    contextUsage,
    refreshContextUsage,
    resetContextUsage,
  };
}
