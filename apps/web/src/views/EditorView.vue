<script setup lang="ts">
import { computed, nextTick, ref, watch } from 'vue'
import { onBeforeRouteLeave, useRoute, useRouter } from 'vue-router'
import { api } from '@/api/client'
import AiChatPanel from '@/components/AiChatPanel.vue'
import AppLogo from '@/components/AppLogo.vue'
import { uiConfirm } from '@/composables/useConfirm'
import { useAuthStore } from '@/stores/auth'
import { useFileTabsStore } from '@/stores/fileTabs'
import { formatSelectionContext } from '@/univer/selection'
import { useUniverHost } from '@/univer/useUniver'

const route = useRoute()
const router = useRouter()
const auth = useAuthStore()
const fileTabs = useFileTabsStore()
const { container, api: univerApi, mountSheet, mountDoc, getSnapshot } =
  useUniverHost()

const docId = computed(() => String(route.params.id))
const docName = ref('…')
const docType = ref<'SHEET' | 'DOC'>('SHEET')
const saveState = ref<'saved' | 'saving' | 'dirty'>('saved')
const saveFeedback = ref<'idle' | 'saving' | 'saved'>('idle')
const chatCollapsed = ref(sessionStorage.getItem('office_chat_collapsed') === '1')
const checkpoint = ref<unknown>(null)
const newMenuOpen = ref(false)
const creating = ref(false)
let saveTimer: ReturnType<typeof setTimeout> | null = null
let saveFeedbackTimer: ReturnType<typeof setTimeout> | null = null
let loadedId = ''
let loadSeq = 0
let dirtyListenerBound = false

const titleWidth = computed(() => {
  const len = Math.max(docName.value.length, 8)
  return `min(${Math.min(len, 40)}ch, 40vw)`
})

async function flushSave() {
  if (saveTimer) {
    clearTimeout(saveTimer)
    saveTimer = null
  }
  if (saveState.value === 'dirty' || saveState.value === 'saving') {
    await save()
  }
}

async function load(id: string) {
  if (!id) return
  const seq = ++loadSeq
  if (loadedId && loadedId !== id) {
    await flushSave()
  }
  if (seq !== loadSeq) return
  checkpoint.value = null
  clearSaveFeedback()
  try {
    const doc = await api.getFile(id)
    if (seq !== loadSeq) return
    docName.value = doc.name
    docType.value = doc.type
    fileTabs.syncActive(doc.id, { name: doc.name, type: doc.type })
    await nextTick()
    if (seq !== loadSeq) return
    if (doc.type === 'SHEET') await mountSheet(doc.snapshot as object)
    else await mountDoc(doc.snapshot as object)
    if (seq !== loadSeq) return
    loadedId = id
    saveState.value = 'saved'
    bindDirtyListeners()
  } catch (e) {
    if (seq !== loadSeq) return
    console.error(e)
    const next = fileTabs.close(id)
    if (next) await router.replace(`/edit/${next}`)
    else await router.replace('/home')
  }
}

function clearSaveFeedback() {
  if (saveFeedbackTimer) {
    clearTimeout(saveFeedbackTimer)
    saveFeedbackTimer = null
  }
  saveFeedback.value = 'idle'
}

function showSavedFeedback() {
  saveFeedback.value = 'saved'
  if (saveFeedbackTimer) clearTimeout(saveFeedbackTimer)
  saveFeedbackTimer = setTimeout(() => {
    saveFeedback.value = 'idle'
    saveFeedbackTimer = null
  }, 2000)
}

function markDirty() {
  saveState.value = 'dirty'
  scheduleSave()
}

function bindDirtyListeners() {
  const el = container.value
  if (!el || dirtyListenerBound) return
  dirtyListenerBound = true
  el.addEventListener('keydown', markDirty, true)
  el.addEventListener('pointerup', markDirty, true)
}

function scheduleSave() {
  if (saveTimer) clearTimeout(saveTimer)
  saveTimer = setTimeout(() => {
    void save()
  }, 1200)
}

function onDocumentChanged() {
  markDirty()
}

async function confirmLeaveIfDirty(): Promise<boolean> {
  if (saveState.value !== 'dirty' && saveState.value !== 'saving') return true
  const ok = await uiConfirm({
    title: 'Сохранить изменения?',
    message:
      'Документ ещё не сохранён. Без сохранения правки пропадут.',
    confirmLabel: 'Сохранить',
    cancelLabel: 'Отмена',
    danger: false,
  })
  if (!ok) return false
  await flushSave()
  return saveState.value !== 'dirty'
}

async function save() {
  const id = loadedId || docId.value
  const snapshot = getSnapshot(docType.value)
  if (!snapshot) return
  saveState.value = 'saving'
  if (saveFeedbackTimer) {
    clearTimeout(saveFeedbackTimer)
    saveFeedbackTimer = null
  }
  saveFeedback.value = 'saving'
  try {
    await api.saveFile(id, snapshot, docName.value)
    saveState.value = 'saved'
    fileTabs.rename(id, docName.value)
    showSavedFeedback()
  } catch (e) {
    saveState.value = 'dirty'
    saveFeedback.value = 'idle'
    console.error(e)
  }
}

function onTitleChange() {
  fileTabs.rename(docId.value, docName.value)
  markDirty()
}

async function download(format: 'json' | 'csv' | 'xlsx' | 'docx') {
  await save()
  if (format === 'xlsx' && docType.value === 'SHEET') {
    const snapshot = getSnapshot('SHEET')
    if (!snapshot) {
      alert('Нет данных таблицы для скачивания')
      return
    }
    try {
      const { downloadWorkbookXlsx } = await import('@/univer/excelImport')
      await downloadWorkbookXlsx(snapshot, docName.value || 'workbook')
    } catch (err) {
      alert(err instanceof Error ? err.message : 'Ошибка экспорта Excel')
    }
    return
  }
  window.open(api.downloadUrl(docId.value, format), '_blank')
}

function toggleChat() {
  chatCollapsed.value = !chatCollapsed.value
  sessionStorage.setItem(
    'office_chat_collapsed',
    chatCollapsed.value ? '1' : '0',
  )
}

function takeCheckpoint() {
  checkpoint.value = getSnapshot(docType.value)
}

async function restoreCheckpoint() {
  if (!checkpoint.value) return
  const ok = await uiConfirm({
    title: 'Откатить документ?',
    message:
      'Документ вернётся к контрольной точке. Текущие изменения после неё будут потеряны.',
    confirmLabel: 'Откатить',
  })
  if (!ok) return
  if (docType.value === 'SHEET') await mountSheet(checkpoint.value as object)
  else await mountDoc(checkpoint.value as object)
  markDirty()
}

function selectionContext() {
  return formatSelectionContext(univerApi.value, docType.value)
}

async function switchTab(id: string) {
  if (id === docId.value) return
  await flushSave()
  fileTabs.focus(id)
  await router.push(`/edit/${id}`)
}

async function closeTab(id: string, e?: Event) {
  e?.stopPropagation()
  if (id === docId.value) await flushSave()
  const next = fileTabs.close(id)
  if (!next) {
    await router.push('/home')
    return
  }
  if (id === docId.value) {
    await router.push(`/edit/${next}`)
  }
}

function toggleNewMenu() {
  newMenuOpen.value = !newMenuOpen.value
}

function closeNewMenu() {
  newMenuOpen.value = false
}

async function createAndOpen(type: 'SHEET' | 'DOC') {
  if (creating.value) return
  creating.value = true
  newMenuOpen.value = false
  try {
    await flushSave()
    const doc = await api.createFile(type)
    fileTabs.open(doc)
    await router.push(`/edit/${doc.id}`)
  } catch (e) {
    console.error(e)
    alert(e instanceof Error ? e.message : 'Не удалось создать файл')
  } finally {
    creating.value = false
  }
}

async function openFromFiles() {
  newMenuOpen.value = false
  await router.push('/home')
}

onBeforeRouteLeave(async (to) => {
  if (String(to.path).startsWith('/edit/')) {
    await flushSave()
    return true
  }
  return confirmLeaveIfDirty()
})

watch(
  docId,
  (id) => {
    if (!id || id === 'undefined') return
    fileTabs.hydrateFromSession()
    void load(id)
  },
  { immediate: true },
)
</script>

<template>
  <div class="editor-page">
    <header class="top">
      <button type="button" class="brand-btn" @click="router.push('/home')">
        <AppLogo :size="28" />
      </button>
      <input
        v-model="docName"
        class="title"
        :style="{ width: titleWidth }"
        @change="onTitleChange"
      />
      <span v-if="saveFeedback !== 'idle'" class="status">
        {{ saveFeedback === 'saving' ? 'Сохранение…' : 'Сохранено' }}
      </span>
      <div class="top-spacer" aria-hidden="true" />
      <div class="actions">
        <button
          type="button"
          class="top-icon-btn"
          title="Сохранить"
          aria-label="Сохранить"
          @click="save"
        >
          <svg viewBox="0 0 24 24" width="18" height="18" aria-hidden="true">
            <path
              d="M5 3h11l3 3v14a1 1 0 0 1-1 1H5a1 1 0 0 1-1-1V4a1 1 0 0 1 1-1z"
              fill="none"
              stroke="currentColor"
              stroke-width="1.5"
            />
            <path
              d="M8 3v6h8V3M8 21v-7h8v7"
              fill="none"
              stroke="currentColor"
              stroke-width="1.5"
            />
          </svg>
        </button>
        <button
          v-if="docType === 'SHEET'"
          type="button"
          class="top-icon-btn"
          title="Скачать Excel"
          aria-label="Скачать Excel"
          @click="download('xlsx')"
        >
          <svg viewBox="0 0 24 24" width="18" height="18" aria-hidden="true">
            <path
              d="M12 3v12m0 0-4-4m4 4 4-4M5 19h14"
              fill="none"
              stroke="currentColor"
              stroke-width="1.5"
              stroke-linecap="round"
              stroke-linejoin="round"
            />
          </svg>
        </button>
        <button
          v-else
          type="button"
          class="top-icon-btn"
          title="Скачать Word"
          aria-label="Скачать Word"
          @click="download('docx')"
        >
          <svg viewBox="0 0 24 24" width="18" height="18" aria-hidden="true">
            <path
              d="M12 3v12m0 0-4-4m4 4 4-4M5 19h14"
              fill="none"
              stroke="currentColor"
              stroke-width="1.5"
              stroke-linecap="round"
              stroke-linejoin="round"
            />
          </svg>
        </button>
        <button type="button" class="ui-btn ui-btn--ghost settings-btn" @click="auth.openSettings()">
          Настройки
        </button>
      </div>
    </header>

    <div class="file-tabs-row" @keydown.escape="closeNewMenu">
      <div class="file-tabs" role="tablist" aria-label="Открытые файлы">
        <button
          v-for="tab in fileTabs.tabs"
          :key="tab.id"
          type="button"
          role="tab"
          class="file-tab"
          :class="{ active: tab.id === docId }"
          :aria-selected="tab.id === docId"
          :title="tab.name"
          @click="switchTab(tab.id)"
          @auxclick.middle.prevent="closeTab(tab.id, $event)"
        >
          <span
            class="file-tab-type"
            :class="tab.type === 'SHEET' ? 'sheet' : 'doc'"
            aria-hidden="true"
          >
            <svg v-if="tab.type === 'SHEET'" viewBox="0 0 24 24" width="14" height="14">
              <rect x="3" y="4" width="18" height="16" rx="2" fill="currentColor" opacity="0.2" />
              <path d="M3 9h18M9 4v16" stroke="currentColor" stroke-width="1.5" fill="none" />
            </svg>
            <svg v-else viewBox="0 0 24 24" width="14" height="14">
              <path
                d="M6 3h8l4 4v14a1 1 0 0 1-1 1H6a1 1 0 0 1-1-1V4a1 1 0 0 1 1-1z"
                fill="currentColor"
                opacity="0.2"
              />
              <path
                d="M14 3v4h4M8 12h8M8 16h5"
                stroke="currentColor"
                stroke-width="1.5"
                fill="none"
                stroke-linecap="round"
              />
            </svg>
          </span>
          <span class="file-tab-name">{{ tab.name }}</span>
          <span
            class="file-tab-close"
            title="Закрыть"
            @click="closeTab(tab.id, $event)"
          >
            ×
          </span>
        </button>
        <div class="file-tabs-new">
          <button
            type="button"
            class="file-tab-plus"
            title="Новая вкладка"
            :disabled="creating"
            aria-haspopup="menu"
            :aria-expanded="newMenuOpen"
            @click="toggleNewMenu"
          >
            +
          </button>
          <div v-if="newMenuOpen" class="ui-dropdown file-tabs-menu" role="menu">
            <button type="button" role="menuitem" @click="createAndOpen('SHEET')">
              Новая таблица
            </button>
            <button type="button" role="menuitem" @click="createAndOpen('DOC')">
              Новый документ
            </button>
            <button type="button" role="menuitem" @click="openFromFiles">
              Открыть файл…
            </button>
          </div>
        </div>
      </div>
      <button
        v-if="newMenuOpen"
        type="button"
        class="file-tabs-backdrop"
        aria-label="Закрыть меню"
        @click="closeNewMenu"
      />
    </div>

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
.top {
  display: flex;
  align-items: center;
  gap: 10px;
  height: var(--app-header-height);
  box-sizing: border-box;
  padding: 0 var(--app-header-padding-x);
  border-bottom: 1px solid var(--border);
  background: var(--bg-1);
}
.title {
  flex: 0 1 auto;
  min-width: 8ch;
  max-width: min(40ch, 40vw);
  height: var(--app-header-control-height);
  box-sizing: border-box;
  padding: 0 2px;
  background: transparent;
  border: 0;
  color: var(--fg-0);
  font-family: var(--font-display);
  font-size: 1rem;
  font-weight: 600;
  letter-spacing: -0.02em;
  overflow: hidden;
  text-overflow: ellipsis;
}
.title:focus {
  outline: none;
}
.title:focus-visible {
  box-shadow: inset 0 -2px 0 var(--accent);
}
.status {
  color: var(--fg-3);
  font-size: 0.85rem;
  flex-shrink: 0;
}
.top-spacer {
  flex: 1;
  min-width: 8px;
}
.actions {
  display: flex;
  align-items: center;
  gap: 4px;
  flex-shrink: 0;
}
.top-icon-btn {
  box-sizing: border-box;
  display: grid;
  place-items: center;
  width: 32px;
  height: 32px;
  padding: 0;
  border: 0;
  border-radius: var(--radius-sm);
  background: transparent;
  color: var(--fg-2);
  cursor: pointer;
  line-height: 0;
}
.top-icon-btn svg {
  display: block;
}
.top-icon-btn:hover {
  color: var(--fg-0);
  background: var(--bg-hover);
}
.top-icon-btn:focus-visible {
  outline: none;
  box-shadow: var(--focus-ring);
}
.brand-btn {
  display: inline-flex;
  align-items: center;
  flex-shrink: 0;
  height: 28px;
  margin: 0;
  padding: 0;
  border: 0;
  appearance: none;
  line-height: 0;
  color: inherit;
  font: inherit;
  text-decoration: none;
  background: none;
  cursor: pointer;
  border-radius: var(--radius-sm);
}
.brand-btn:hover {
  background: var(--bg-hover);
}
.brand-btn:focus-visible {
  outline: none;
  box-shadow: var(--focus-ring);
}

@media (max-width: 720px) {
  .top {
    padding-left: 16px;
    padding-right: 16px;
  }
  .status {
    display: none;
  }
  .settings-btn {
    display: none;
  }
}

.file-tabs-row {
  position: relative;
  display: flex;
  align-items: stretch;
  background: var(--bg-1);
  border-bottom: 1px solid var(--border);
  min-height: 34px;
}
.file-tabs {
  display: flex;
  align-items: stretch;
  gap: 0;
  overflow-x: auto;
  min-width: 0;
  flex: 1;
  scrollbar-width: thin;
}
.file-tabs-new {
  position: relative;
  flex: 0 0 auto;
  display: flex;
  align-items: center;
  align-self: center;
  margin-left: 2px;
  padding: 0 4px;
  z-index: 2;
}
.file-tab-plus {
  width: 28px;
  height: 28px;
  padding: 0;
  border: 0;
  border-radius: var(--radius-sm);
  background: transparent;
  color: var(--fg-2);
  font-size: 1.15rem;
  line-height: 1;
  cursor: pointer;
}
.file-tab-plus:hover:not(:disabled) {
  background: var(--bg-hover);
  color: var(--fg-0);
}
.file-tab-plus:disabled {
  opacity: 0.5;
  cursor: default;
}
.file-tabs-menu {
  top: calc(100% + 4px);
  left: 0;
  min-width: 180px;
  z-index: 3;
}
.file-tabs-backdrop {
  position: fixed;
  inset: 0;
  border: 0;
  padding: 0;
  margin: 0;
  background: transparent;
  cursor: default;
  z-index: 1;
}
.file-tab {
  display: inline-flex;
  align-items: center;
  gap: 6px;
  max-width: 200px;
  min-width: 96px;
  padding: 6px 8px 6px 10px;
  border: 0;
  border-right: 1px solid var(--border);
  border-radius: 0;
  background: var(--bg-2);
  color: var(--fg-2);
  font-size: 0.85rem;
  cursor: pointer;
  flex: 0 0 auto;
}
.file-tab:hover {
  background: var(--bg-hover);
  color: var(--fg-1);
}
.file-tab.active {
  background: var(--bg-0);
  color: var(--fg-0);
  box-shadow: inset 0 -2px 0 var(--accent);
}
.file-tab-type {
  flex: 0 0 auto;
  width: 18px;
  height: 18px;
  display: grid;
  place-items: center;
  border-radius: 4px;
}
.file-tab-type.sheet {
  color: var(--success);
  background: color-mix(in srgb, var(--success) 14%, transparent);
}
.file-tab-type.doc {
  color: var(--accent);
  background: color-mix(in srgb, var(--accent) 14%, transparent);
}
.file-tab-name {
  flex: 1;
  min-width: 0;
  overflow: hidden;
  text-overflow: ellipsis;
  white-space: nowrap;
  text-align: left;
}
.file-tab-close {
  flex: 0 0 auto;
  width: 18px;
  height: 18px;
  line-height: 16px;
  text-align: center;
  border-radius: var(--radius-sm);
  opacity: 0;
  color: var(--fg-3);
  font-size: 1rem;
}
.file-tab:hover .file-tab-close,
.file-tab.active .file-tab-close {
  opacity: 1;
}
.file-tab-close:hover {
  background: var(--bg-3);
  color: var(--fg-0);
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
