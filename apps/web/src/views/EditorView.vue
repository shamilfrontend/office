<script setup lang="ts">
import { computed, nextTick, ref, watch } from 'vue'
import { useRoute, useRouter } from 'vue-router'
import { api } from '@/api/client'
import AiChatPanel from '@/components/AiChatPanel.vue'
import AppLogo from '@/components/AppLogo.vue'
import { uiConfirm } from '@/composables/useConfirm'
import { useAuthStore } from '@/stores/auth'
import { useFileTabsStore } from '@/stores/fileTabs'
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
const chatCollapsed = ref(sessionStorage.getItem('office_chat_collapsed') === '1')
const checkpoint = ref<unknown>(null)
const newMenuOpen = ref(false)
const creating = ref(false)
let saveTimer: ReturnType<typeof setTimeout> | null = null
let loadedId = ''
let loadSeq = 0
let dirtyListenerBound = false

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

function bindDirtyListeners() {
  const el = container.value
  if (!el || dirtyListenerBound) return
  dirtyListenerBound = true
  el.addEventListener(
    'keydown',
    () => {
      saveState.value = 'dirty'
      scheduleSave()
    },
    true,
  )
  el.addEventListener(
    'pointerup',
    () => {
      saveState.value = 'dirty'
      scheduleSave()
    },
    true,
  )
}

function scheduleSave() {
  if (saveTimer) clearTimeout(saveTimer)
  saveTimer = setTimeout(() => {
    void save()
  }, 1200)
}

async function save() {
  const id = loadedId || docId.value
  const snapshot = getSnapshot(docType.value)
  if (!snapshot) return
  saveState.value = 'saving'
  try {
    await api.saveFile(id, snapshot, docName.value)
    saveState.value = 'saved'
    fileTabs.rename(id, docName.value)
  } catch (e) {
    saveState.value = 'dirty'
    console.error(e)
  }
}

function onTitleChange() {
  fileTabs.rename(docId.value, docName.value)
  scheduleSave()
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
  saveState.value = 'dirty'
  scheduleSave()
}

function selectionContext() {
  return `Тип: ${docType.value}. Имя: ${docName.value}. Используй tools для чтения содержимого.`
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
  await flushSave()
  await router.push('/home')
}

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
        <AppLogo :size="24" :with-label="false" />
        <span>Файлы</span>
      </button>
      <input v-model="docName" class="title" @change="onTitleChange" />
      <span class="status">
        {{
          saveState === 'saved'
            ? 'Сохранено'
            : saveState === 'saving'
              ? 'Сохранение…'
              : 'Есть изменения'
        }}
      </span>
      <button type="button" @click="save">Сохранить</button>
      <button
        v-if="docType === 'SHEET'"
        type="button"
        @click="download('xlsx')"
      >
        Скачать Excel
      </button>
      <button v-else type="button" @click="download('docx')">
        Скачать Word
      </button>
      <button type="button" class="link" @click="auth.openSettings()">
        Настройки
      </button>
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
          <span class="file-tab-type">{{ tab.type === 'SHEET' ? '⊞' : '¶' }}</span>
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
          <div v-if="newMenuOpen" class="file-tabs-menu" role="menu">
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
  padding: 8px 12px;
  border-bottom: 1px solid var(--border);
  background: var(--bg-1);
}
.title {
  flex: 1;
  min-width: 120px;
  background: transparent;
  border: 0;
  color: var(--fg-0);
  font-size: 1rem;
  font-weight: 600;
}
.status {
  color: var(--fg-3);
  font-size: 0.85rem;
}
.brand-btn {
  display: inline-flex;
  align-items: center;
  gap: 8px;
  color: var(--fg-1);
  text-decoration: none;
  background: none;
  border: 0;
  cursor: pointer;
  padding: 4px 6px;
  border-radius: 8px;
}
.brand-btn:hover {
  background: var(--bg-hover);
}
button {
  border: 1px solid var(--border);
  background: var(--bg-2);
  color: var(--fg-1);
  border-radius: 8px;
  padding: 6px 10px;
  cursor: pointer;
}
button.link {
  color: var(--fg-2);
  text-decoration: none;
  background: none;
  border: 0;
  padding: 0;
  border-radius: 0;
  font: inherit;
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
  border-radius: 6px;
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
  position: absolute;
  top: calc(100% + 4px);
  left: 0;
  min-width: 180px;
  padding: 4px;
  border: 1px solid var(--border);
  border-radius: 8px;
  background: var(--bg-1);
  box-shadow: var(--shadow);
  display: flex;
  flex-direction: column;
  gap: 2px;
  z-index: 3;
}
.file-tabs-menu button {
  border: 0;
  border-radius: 6px;
  background: transparent;
  color: var(--fg-1);
  text-align: left;
  padding: 8px 10px;
  cursor: pointer;
}
.file-tabs-menu button:hover {
  background: var(--bg-hover);
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
  padding: 6px 8px 6px 12px;
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
  opacity: 0.7;
  font-size: 0.8rem;
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
  border-radius: 4px;
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
  background: #fff;
}
</style>
