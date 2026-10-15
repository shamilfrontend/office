<script setup lang="ts">
import { computed, nextTick, ref, watch } from 'vue';
import { onBeforeRouteLeave, useRoute, useRouter } from 'vue-router';
import { api } from '@/api/client';
import AiChatPanel from '@/components/AiChatPanel';
import { uiConfirm } from '@/composables/useConfirm';
import { useAuthStore } from '@/stores/auth';
import { useFileTabsStore } from '@/stores/fileTabs';
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
} = useUniverHost();

const docId = computed(() => String(route.params.id));
const docName = ref('…');
const docType = ref<'SHEET' | 'DOC'>('SHEET');
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
    if (doc.type === 'SHEET') await mountSheet(doc.snapshot as object);
    else await mountDoc(doc.snapshot as object);
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

function bindDirtyListeners() {
  const el = container.value;
  if (!el || dirtyListenerBound) return;
  dirtyListenerBound = true;
  el.addEventListener('keydown', markDirty, true);
  el.addEventListener('pointerup', markDirty, true);
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

async function save() {
  const id = loadedId || docId.value;
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

async function download(format: 'json' | 'csv' | 'xlsx' | 'docx') {
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

function takeCheckpoint() {
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
  if (docType.value === 'SHEET') await mountSheet(checkpoint.value as object);
  else await mountDoc(checkpoint.value as object);
  markDirty();
}

function selectionContext() {
  return formatSelectionContext(univerApi.value, docType.value);
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

async function createAndOpen(type: 'SHEET' | 'DOC') {
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
