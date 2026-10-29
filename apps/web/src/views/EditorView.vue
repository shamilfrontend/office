<script setup lang="ts">
import { computed, nextTick, ref, watch } from 'vue';
import { onBeforeRouteLeave, useRoute, useRouter } from 'vue-router';
import { api } from '@/api/client';
import AiChatPanel from '@/components/AiChatPanel';
import { uiConfirm } from '@/composables/useConfirm';
import { useAuthStore } from '@/stores/auth';
import { useFileTabsStore } from '@/stores/fileTabs';
import { base64ToBytes, bytesToBase64 } from '@/slides/pptxBytes';
import { useSlideEditor } from '@/slides/useSlideEditor';
import { formatSelectionContext } from '@/univer/selection';
import { useUniverHost } from '@/univer/useUniver';
import EditorTopBar from './editor/EditorTopBar.vue';
import EditorFileTabs from './editor/EditorFileTabs.vue';

const route = useRoute();
const router = useRouter();
const auth = useAuthStore();
const fileTabs = useFileTabsStore();
const {
  container,
  api: univerApi,
  mountSheet,
  mountDoc,
  getSnapshot,
  destroy: destroyUniver,
} = useUniverHost();
const {
  ready: slideReady,
  toolDefs: slideToolDefs,
  mount: mountSlide,
  destroy: destroySlide,
  save: saveSlide,
  open: openSlide,
  snapshot: snapshotSlide,
  run: runSlideTool,
  selectionText: slideSelectionText,
} = useSlideEditor(container);

const docId = computed(() => String(route.params.id));
const docName = ref('…');
const docType = ref<'SHEET' | 'DOC' | 'SLIDE'>('SHEET');
const saveState = ref<'saved' | 'saving' | 'dirty'>('saved');
const saveFeedback = ref<'idle' | 'saving' | 'saved'>('idle');
const chatCollapsed = ref(
  sessionStorage.getItem('office_chat_collapsed') === '1',
);
const checkpoint = ref<unknown>(null);
const newMenuOpen = ref(false);
const creating = ref(false);
let saveTimer: ReturnType<typeof setTimeout> | null = null;
let saveFeedbackTimer: ReturnType<typeof setTimeout> | null = null;
let loadedId = '';
let loadSeq = 0;
let dirtyListenerBound = false;

const titleWidth = computed(() => {
  const len = Math.max(docName.value.length, 8);
  return `min(${Math.min(len, 40)}ch, 40vw)`;
});

async function flushSave() {
  if (saveTimer) {
    clearTimeout(saveTimer);
    saveTimer = null;
  }
  if (saveState.value === 'dirty' || saveState.value === 'saving') {
    await save();
  }
}

async function load(id: string) {
  if (!id) return;
  loadSeq += 1;
  const seq = loadSeq;
  if (loadedId && loadedId !== id) {
    await flushSave();
  }
  if (seq !== loadSeq) return;
  checkpoint.value = null;
  clearSaveFeedback();
  try {
    const doc = await api.getFile(id);
    if (seq !== loadSeq) return;
    docName.value = doc.name;
    docType.value = doc.type;
    fileTabs.syncActive(doc.id, { name: doc.name, type: doc.type });
    await nextTick();
    if (seq !== loadSeq) return;
    await destroySlide();
    if (doc.type === 'SLIDE') {
      await destroyUniver();
      const raw = doc.snapshot;
      const bytes = base64ToBytes(typeof raw === 'string' ? raw : '');
      await mountSlide(bytes, doc.name, {
        onDirty: () => {
          if (docType.value === 'SLIDE') markDirty();
        },
        onSave: persistSlideBytes,
      });
    } else if (doc.type === 'SHEET') {
      await mountSheet(doc.snapshot as object);
    } else {
      await mountDoc(doc.snapshot as object);
    }
    if (seq !== loadSeq) return;
    loadedId = id;
    saveState.value = 'saved';
    bindDirtyListeners();
  } catch (e) {
    if (seq !== loadSeq) return;
    console.error(e);
    const next = fileTabs.close(id);
    if (next) await router.replace({ name: 'edit', params: { id: next } });
    else await router.replace({ name: 'home' });
  }
}

function clearSaveFeedback() {
  if (saveFeedbackTimer) {
    clearTimeout(saveFeedbackTimer);
    saveFeedbackTimer = null;
  }
  saveFeedback.value = 'idle';
}

function showSavedFeedback() {
  saveFeedback.value = 'saved';
  if (saveFeedbackTimer) clearTimeout(saveFeedbackTimer);
  saveFeedbackTimer = setTimeout(() => {
    saveFeedback.value = 'idle';
    saveFeedbackTimer = null;
  }, 2000);
}

function markDirty() {
  saveState.value = 'dirty';
  scheduleSave();
}

function onHostEdit() {
  if (docType.value === 'SLIDE') return;
  markDirty();
}

function bindDirtyListeners() {
  const el = container.value;
  if (!el || dirtyListenerBound) return;
  dirtyListenerBound = true;
  el.addEventListener('keydown', onHostEdit, true);
  el.addEventListener('pointerup', onHostEdit, true);
}

function scheduleSave() {
  if (saveTimer) clearTimeout(saveTimer);
  saveTimer = setTimeout(() => {
    void save();
  }, 1200);
}

function onDocumentChanged() {
  markDirty();
}

async function confirmLeaveIfDirty(): Promise<boolean> {
  if (saveState.value !== 'dirty' && saveState.value !== 'saving') return true;
  const ok = await uiConfirm({
    title: 'Сохранить изменения?',
    message: 'Документ ещё не сохранён. Без сохранения правки пропадут.',
    confirmLabel: 'Сохранить',
    cancelLabel: 'Отмена',
    danger: false,
  });
  if (!ok) return false;
  await flushSave();
  return saveState.value !== 'dirty';
}

async function persistSlideBytes(pptx: Uint8Array) {
  const id = loadedId || docId.value;
  try {
    await api.saveFile(id, bytesToBase64(pptx), docName.value);
    saveState.value = 'saved';
    fileTabs.rename(id, docName.value);
    showSavedFeedback();
  } catch (e) {
    saveState.value = 'dirty';
    saveFeedback.value = 'idle';
    console.error(e);
    throw e;
  }
}

async function save() {
  const id = loadedId || docId.value;
  if (docType.value === 'SLIDE') {
    if (!slideReady.value) return;
    saveState.value = 'saving';
    if (saveFeedbackTimer) {
      clearTimeout(saveFeedbackTimer);
      saveFeedbackTimer = null;
    }
    saveFeedback.value = 'saving';
    try {
      await saveSlide();
      if (saveState.value === 'saving') {
        saveState.value = 'saved';
        fileTabs.rename(id, docName.value);
        showSavedFeedback();
      }
    } catch {
      saveState.value = 'dirty';
      saveFeedback.value = 'idle';
    }
    return;
  }
  const snapshot = getSnapshot(docType.value);
  if (!snapshot) return;
  saveState.value = 'saving';
  if (saveFeedbackTimer) {
    clearTimeout(saveFeedbackTimer);
    saveFeedbackTimer = null;
  }
  saveFeedback.value = 'saving';
  try {
    await api.saveFile(id, snapshot, docName.value);
    saveState.value = 'saved';
    fileTabs.rename(id, docName.value);
    showSavedFeedback();
  } catch (e) {
    saveState.value = 'dirty';
    saveFeedback.value = 'idle';
    console.error(e);
  }
}

function onTitleChange() {
  fileTabs.rename(docId.value, docName.value);
  markDirty();
}

async function download(format: 'json' | 'csv' | 'xlsx' | 'docx' | 'pptx') {
  await save();
  if (format === 'xlsx' && docType.value === 'SHEET') {
    const snapshot = getSnapshot('SHEET');
    if (!snapshot) {
      alert('Нет данных таблицы для скачивания');
      return;
    }
    try {
      const { downloadWorkbookXlsx } = await import('@/univer/excelImport');
      await downloadWorkbookXlsx(snapshot, docName.value || 'workbook');
    } catch (err) {
      alert(err instanceof Error ? err.message : 'Ошибка экспорта Excel');
    }
    return;
  }
  window.open(api.downloadUrl(docId.value, format), '_blank');
}

function toggleChat() {
  chatCollapsed.value = !chatCollapsed.value;
  sessionStorage.setItem(
    'office_chat_collapsed',
    chatCollapsed.value ? '1' : '0',
  );
}

async function takeCheckpoint() {
  if (docType.value === 'SLIDE') {
    checkpoint.value = await snapshotSlide();
    return;
  }
  checkpoint.value = getSnapshot(docType.value);
}

async function restoreCheckpoint() {
  if (!checkpoint.value) return;
  const ok = await uiConfirm({
    title: 'Откатить документ?',
    message:
      'Документ вернётся к контрольной точке. Текущие изменения после неё будут потеряны.',
    confirmLabel: 'Откатить',
  });
  if (!ok) return;
  if (docType.value === 'SLIDE' && checkpoint.value instanceof Uint8Array) {
    await openSlide(checkpoint.value);
  } else if (docType.value === 'SHEET') {
    await mountSheet(checkpoint.value as object);
  } else if (docType.value === 'DOC') {
    await mountDoc(checkpoint.value as object);
  }
  markDirty();
}

function selectionContext() {
  if (docType.value === 'SLIDE') return slideSelectionText();
  return formatSelectionContext(univerApi.value, docType.value);
}

function slideTools() {
  return slideToolDefs.value;
}

async function switchTab(id: string) {
  if (id === docId.value) return;
  await flushSave();
  fileTabs.focus(id);
  await router.push({ name: 'edit', params: { id } });
}

async function closeTab(id: string, e?: Event) {
  e?.stopPropagation();
  if (id === docId.value) await flushSave();
  const next = fileTabs.close(id);
  if (!next) {
    await router.push({ name: 'home' });
    return;
  }
  if (id === docId.value) {
    await router.push({ name: 'edit', params: { id: next } });
  }
}

function toggleNewMenu() {
  newMenuOpen.value = !newMenuOpen.value;
}

function closeNewMenu() {
  newMenuOpen.value = false;
}

async function createAndOpen(type: 'SHEET' | 'DOC' | 'SLIDE') {
  if (creating.value) return;
  creating.value = true;
  newMenuOpen.value = false;
  try {
    await flushSave();
    const doc = await api.createFile(type);
    fileTabs.open(doc);
    await router.push({ name: 'edit', params: { id: doc.id } });
  } catch (e) {
    console.error(e);
    alert(e instanceof Error ? e.message : 'Не удалось создать файл');
  } finally {
    creating.value = false;
  }
}

async function openFromFiles() {
  newMenuOpen.value = false;
  await router.push({ name: 'home' });
}

onBeforeRouteLeave(async (to) => {
  if (String(to.path).startsWith('/edit/')) {
    await flushSave();
    return true;
  }
  return confirmLeaveIfDirty();
});

watch(
  docId,
  (id) => {
    if (!id || id === 'undefined') return;
    fileTabs.hydrateFromSession();
    void load(id);
  },
  { immediate: true },
);
</script>

<template>
  <div class="editor-page">
    <EditorTopBar
      :doc-name="docName"
      :title-width="titleWidth"
      :save-feedback="saveFeedback"
      :doc-type="docType"
      @update:doc-name="docName = $event"
      @title-change="onTitleChange"
      @home="router.push({ name: 'home' })"
      @save="save"
      @download="download"
      @settings="auth.openSettings()"
    />

    <EditorFileTabs
      :tabs="fileTabs.tabs"
      :active-id="docId"
      :creating="creating"
      :new-menu-open="newMenuOpen"
      @switch="switchTab"
      @close="closeTab"
      @toggle-new="toggleNewMenu"
      @close-new="closeNewMenu"
      @create="createAndOpen"
      @open-files="openFromFiles"
    />

    <div class="main">
      <div ref="container" class="univer-host" />
      <AiChatPanel
        :collapsed="chatCollapsed"
        :doc-type="docType"
        :doc-name="docName"
        :document-id="docId"
        :univer-api="univerApi"
        :slide-ready="slideReady"
        :get-slide-tools="slideTools"
        :run-slide-tool="runSlideTool"
        :get-selection-context="selectionContext"
        @toggle="toggleChat"
        @checkpoint="takeCheckpoint"
        @restore-checkpoint="restoreCheckpoint"
        @document-changed="onDocumentChanged"
      />
    </div>
  </div>
</template>

<style scoped lang="scss">
.editor-page {
  height: 100%;
  display: flex;
  flex-direction: column;
  background: var(--bg-0);
}
.main {
  flex: 1;
  min-height: 0;
  display: flex;
}
.univer-host {
  flex: 1;
  min-width: 0;
  min-height: 0;
  position: relative;
  background: var(--canvas);
}
</style>
