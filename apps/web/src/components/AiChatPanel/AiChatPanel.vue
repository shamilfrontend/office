<script setup lang="ts">
import {
  computed,
  nextTick,
  onBeforeUnmount,
  onMounted,
  ref,
  watch,
} from 'vue';
import { useRouter } from 'vue-router';
import type { OfficeToolDef } from '@office/shared';
import { useAuthStore } from '@/stores/auth';
import type { UniverApi } from '@/univer/useUniver';
import { useAgentChat } from './composables/useAgentChat';
import { useChatHistory } from './composables/useChatHistory';
import { useChatPlan } from './composables/useChatPlan';
import { useChatTabs } from './composables/useChatTabs';
import { useComposerAttachments } from './composables/useComposerAttachments';
import { useContextUsage } from './composables/useContextUsage';
import { useFloatingMenus } from './composables/useFloatingMenus';
import { useModelMenu } from './composables/useModelMenu';
import { usePanelResize } from './composables/usePanelResize';
import { useSessionLog } from './composables/useSessionLog';
import { useStreamHtml } from './composables/useStreamHtml';
import { useTranscriptUi } from './composables/useTranscriptUi';
import ChatComposer from './parts/ChatComposer.vue';
import ChatEmptyState from './parts/ChatEmptyState.vue';
import ChatHistoryPanel from './parts/ChatHistoryPanel.vue';
import ChatPlanCard from './parts/ChatPlanCard.vue';
import ChatPlanModal from './parts/ChatPlanModal.vue';
import ChatProviderCta from './parts/ChatProviderCta.vue';
import ChatTabsBar from './parts/ChatTabsBar.vue';
import ChatTranscript from './parts/ChatTranscript.vue';

const props = defineProps<{
  collapsed: boolean;
  docType?: 'SHEET' | 'DOC' | 'SLIDE' | null;
  docName?: string;
  documentId?: string;
  univerApi?: UniverApi | null;
  slideReady?: boolean;
  getSlideTools?: () => OfficeToolDef[];
  runSlideTool?: (name: string, input: unknown) => Promise<unknown>;
  getSelectionContext?: () => string;
}>();

const emit = defineEmits<{
  toggle: [];
  checkpoint: [];
  restoreCheckpoint: [];
  documentChanged: [];
}>();

const auth = useAuthStore();
const router = useRouter();
const { panelStyle, startResize } = usePanelResize(() => props.collapsed);
const modeMenu = ref(false);
const modelMenu = ref(false);
const historyOpen = ref(false);
const tabsBarRef = ref<{ historyBtn: HTMLElement | null } | null>(null);
const historyPanelRef = ref<{ panelEl: HTMLElement | null } | null>(null);
const transcriptRef = ref<{ el: HTMLElement | null } | null>(null);
const composerRef = ref<{
  composerEl: HTMLTextAreaElement | null;
  modeWrap: HTMLElement | null;
  modelWrap: HTMLElement | null;
  fileInput: HTMLInputElement | null;
} | null>(null);
const planModalRef = ref<{ modalEl: HTMLElement | null } | null>(null);
const attachError = ref('');
const approveExpanded = ref(false);
let clearLastTokenRateImpl: () => void = () => {};
let ensureDefaultModelImpl: () => Promise<void> = async () => {};
let loadHistoryIfOpenImpl: () => void = () => {};

const { clearStreamHtml, flushStreamHtml, scheduleStreamHtml, assistantHtml } =
  useStreamHtml();

const {
  sessionLog,
  logCopied,
  hasSessionLog,
  pushSessionLog,
  copySessionLog,
  resetSessionLog,
} = useSessionLog({
  documentId: () => props.documentId,
  docName: () => props.docName,
  docType: () => props.docType,
  attachError,
});

const attachHelpText =
  'Можно прикрепить файл (скрепка): Excel, Word, PowerPoint, PDF, TXT, MD, CSV, JSON. Текст файла попадёт в контекст чата (до ~20 тыс. символов).';

const hasProvider = computed(() => Boolean(auth.me?.hasAiProvider));
const providers = computed(
  () => auth.me?.providers?.filter((p) => p.enabled) ?? [],
);

const {
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
  commitRename,
  onRenameKey,
  resetTabsState,
} = useChatTabs({
  documentId: () => props.documentId,
  hasProvider,
  providers,
  attachError,
  ensureDefaultModel: () => ensureDefaultModelImpl(),
  loadHistoryIfOpen: () => loadHistoryIfOpenImpl(),
});

const {
  historyItems,
  historyLoading,
  historyError,
  loadHistory,
  toggleHistory,
  reopenChat,
  deleteFromHistory,
  resetHistory,
} = useChatHistory({
  documentId: () => props.documentId,
  tabs,
  activeId,
  providers: () => providers.value,
  ensureDefaultModel: () => ensureDefaultModelImpl(),
  modeMenu,
  modelMenu,
  historyOpen,
});
loadHistoryIfOpenImpl = () => {
  if (historyOpen.value) void loadHistory();
};

const {
  modelOptions,
  modelMenuLoading,
  modelMenuError,
  setMode,
  toggleModelMenu,
  setModel,
  ensureDefaultModel,
  modelLabel,
  resetModelMenu,
} = useModelMenu({
  active,
  providers,
  patchTab,
  modeMenu,
  modelMenu,
  historyOpen,
});
ensureDefaultModelImpl = ensureDefaultModel;

const { contextUsage, refreshContextUsage, resetContextUsage } =
  useContextUsage({
    active,
    docName: () => props.docName,
    docType: () => props.docType,
    documentId: () => props.documentId,
    getSelectionContext: () => props.getSelectionContext,
    getSlideTools: () => props.getSlideTools?.() ?? [],
  });

const {
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
} = useTranscriptUi({
  active,
  docType: () => props.docType,
  attachError,
  transcriptEl: () => transcriptRef.value?.el,
});

const hasMessages = computed(() => Boolean(active.value?.messages.length));

function resetTabs() {
  resetTabsState();
  resetHistory();
  resetModelMenu();
  modeMenu.value = false;
  resetSessionLog();
  resetTranscriptUi();
  approveExpanded.value = false;
  clearStreamHtml();
  clearLastTokenRateImpl();
  resetContextUsage();
}

watch(
  () => active.value?.pendingApprove,
  () => {
    approveExpanded.value = false;
  },
);

watch(hasProvider, (ok) => {
  if (ok && props.documentId && !tabs.value.length) void loadTabs();
});

watch(
  () => props.documentId,
  (id, prev) => {
    if (id === prev) return;
    resetTabs();
    if (id) void loadTabs();
  },
);

onMounted(() => {
  if (props.documentId) void loadTabs();
});
onBeforeUnmount(() => {
  resetTabs();
});

const { send, retryTurn, stop, resolveApprove, clearLastTokenRate } =
  useAgentChat({
    props,
    emit,
    active,
    hasProvider,
    providers,
    attachError,
    approveExpanded,
    findTab,
    patchTab,
    pushSessionLog,
    toggleModelMenu,
    resizeComposer,
    refreshContextUsage,
    clearStreamHtml,
    flushStreamHtml,
    scheduleStreamHtml,
  });
clearLastTokenRateImpl = clearLastTokenRate;

const { planOpen, planDraft, planOverview, openPlan, savePlan, runPlan } =
  useChatPlan({
    active,
    patchTab,
    send,
  });

const { pickAttach, onAttach, removeAttach } = useComposerAttachments({
  active,
  attachError,
  fileInput: () => composerRef.value?.fileInput,
});

useFloatingMenus({
  modeMenu,
  modelMenu,
  historyOpen,
  planOpen,
  modeWrap: () => composerRef.value?.modeWrap,
  modelWrap: () => composerRef.value?.modelWrap,
  historyPanel: () => historyPanelRef.value?.panelEl,
  historyBtn: () => tabsBarRef.value?.historyBtn,
  planModal: () => planModalRef.value?.modalEl,
});

function goSettings() {
  void router.push({ name: 'home' });
  auth.openSettings();
}

function resizeComposer() {
  const el = composerRef.value?.composerEl;
  if (!el) return;
  el.style.height = 'auto';
  el.style.height = `${el.scrollHeight}px`;
}

function onComposerKey(event: KeyboardEvent) {
  if (event.key !== 'Enter' || event.isComposing) return;
  if (event.shiftKey) return;
  event.preventDefault();
  void send();
}

watch(
  () => [activeId.value, active.value?.draft] as const,
  async () => {
    await nextTick();
    resizeComposer();
  },
);
</script>

<template>
  <aside class="panel" :class="{ collapsed }" :style="panelStyle">
    <div v-if="!collapsed" class="resize" @pointerdown="startResize" />
    <button
      v-if="collapsed"
      type="button"
      class="expand"
      @click="emit('toggle')"
    >
      <span class="expand-label">AI</span>
    </button>
    <template v-else>
      <ChatTabsBar
        ref="tabsBarRef"
        :tabs="tabs"
        :active-id="activeId"
        :renaming-id="renamingId"
        :rename-draft="renameDraft"
        :history-open="historyOpen"
        :has-session-log="hasSessionLog"
        :log-copied="logCopied"
        @update:active-id="activeId = $event"
        @update:rename-draft="renameDraft = $event"
        @add="addTab"
        @close="closeTab"
        @start-rename="startRename"
        @commit-rename="commitRename"
        @rename-key="onRenameKey"
        @toggle-history="toggleHistory"
        @copy-log="copySessionLog"
        @collapse="emit('toggle')"
      />

      <ChatHistoryPanel
        v-if="historyOpen"
        ref="historyPanelRef"
        :items="historyItems"
        :loading="historyLoading"
        :error="historyError"
        @close="historyOpen = false"
        @reopen="reopenChat"
        @delete="deleteFromHistory"
      />

      <div v-if="docName" class="doc-context" :title="docName">
        Файл: {{ docName }}
      </div>

      <ChatProviderCta v-if="!hasProvider" @settings="goSettings" />
      <div v-else-if="!active" class="cta">
        <p>Открываю чат…</p>
      </div>
      <div v-else class="stack" :class="{ empty: !hasMessages }">
        <ChatEmptyState
          v-if="!hasMessages"
          :mode="active.mode"
          :hints="emptyHints"
          :streaming="active.streaming"
          :attach-help-text="attachHelpText"
          @apply-hint="applyHint"
        />

        <ChatTranscript
          v-else
          ref="transcriptRef"
          :messages="active.messages"
          :streaming="active.streaming"
          :token-rate="active.tokenRate"
          :token-meta="active.tokenMeta"
          :show-checkpoint="active.showCheckpoint"
          :can-retry="active.canRetry"
          :retry-reason="active.retryReason"
          :pending-approve="
            active.pendingApprove
              ? {
                  name: active.pendingApprove.name,
                  args: active.pendingApprove.args,
                }
              : null
          "
          :approve-summary="approveSummary"
          :approve-expanded="approveExpanded"
          :copied-key="copiedKey"
          :assistant-html="assistantHtml"
          :is-tool-open="isToolOpen"
          @copy-message="copyText"
          @toggle-tool="toggleTool"
          @transcript-click="onTranscriptClick"
          @restore-checkpoint="emit('restoreCheckpoint')"
          @retry="retryTurn"
          @update:approve-expanded="approveExpanded = $event"
          @resolve-approve="resolveApprove"
        />

        <ChatPlanCard
          v-if="active.planMarkdown && active.planRunState !== 'done'"
          :title="active.planTitle"
          :overview="planOverview"
          :run-state="active.planRunState"
          :streaming="active.streaming"
          @open="openPlan"
          @run="runPlan"
        />

        <ChatComposer
          ref="composerRef"
          :draft="active.draft"
          :mode="active.mode"
          :model-label-text="modelLabel(active)"
          :placeholder="composerPlaceholder"
          :streaming="active.streaming"
          :attachments="active.attachments"
          :attach-error="attachError"
          :attach-help-text="attachHelpText"
          :context-usage="contextUsage"
          :mode-menu="modeMenu"
          :model-menu="modelMenu"
          :model-options="modelOptions"
          :model-menu-loading="modelMenuLoading"
          :model-menu-error="modelMenuError"
          :providers-count="providers.length"
          :active-provider-id="active.providerId"
          :active-model="active.model"
          @update:draft="active.draft = $event"
          @update:mode-menu="
            modeMenu = $event;
            historyOpen = false;
          "
          @update:model-menu="modelMenu = $event"
          @set-mode="setMode"
          @toggle-model-menu="toggleModelMenu"
          @set-model="setModel"
          @pick-attach="pickAttach"
          @attach-change="onAttach"
          @remove-attach="removeAttach"
          @send="send()"
          @stop="stop"
          @composer-key="onComposerKey"
          @composer-input="resizeComposer"
        />
      </div>
    </template>

    <ChatPlanModal
      ref="planModalRef"
      :open="planOpen"
      :draft="planDraft"
      @update:open="planOpen = $event"
      @update:draft="planDraft = $event"
      @save="savePlan"
    />
  </aside>
</template>

<style scoped lang="scss">
.panel {
  position: relative;
  width: var(--chat-width);
  border-left: 1px solid var(--border);
  background: var(--bg-1);
  display: flex;
  flex-direction: column;
  min-height: 0;
  min-width: 0;
}
.panel.collapsed {
  width: 40px;
}
.resize {
  position: absolute;
  left: -3px;
  top: 0;
  bottom: 0;
  width: 6px;
  cursor: col-resize;
  z-index: 3;
}
.expand {
  height: 100%;
  width: 100%;
  border: 0;
  background: var(--bg-2);
  color: var(--fg-2);
  cursor: pointer;
  display: grid;
  place-items: center;
}
.expand:hover {
  color: var(--fg-0);
  background: var(--bg-hover);
}
.expand-label {
  writing-mode: vertical-rl;
  letter-spacing: 0.08em;
  font-size: 0.78rem;
  font-weight: 650;
}
.doc-context {
  padding: 5px 12px;
  font-size: 0.72rem;
  color: var(--fg-3);
  border-bottom: 1px solid var(--border);
  overflow: hidden;
  text-overflow: ellipsis;
  white-space: nowrap;
}
.cta,
.stack {
  flex: 1;
  min-height: 0;
  display: flex;
  flex-direction: column;
}
.cta {
  align-items: center;
  justify-content: center;
  gap: 10px;
  padding: 28px 24px;
  color: var(--fg-2);
  text-align: center;
}
.stack.empty :deep(.transcript) {
  display: none;
}
.stack :deep(.plan-card) {
  order: 2;
}
.stack :deep(.composer) {
  order: 3;
}
.stack.empty :deep(.empty-state) {
  order: 1;
}
.stack.empty :deep(.plan-card) {
  order: 2;
}
.stack.empty :deep(.composer) {
  order: 3;
}
</style>
