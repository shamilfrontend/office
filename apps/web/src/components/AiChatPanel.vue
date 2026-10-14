<script setup lang="ts">
import { computed, nextTick, onBeforeUnmount, onMounted, ref, watch } from 'vue'
import DOMPurify from 'dompurify'
import { marked } from 'marked'
import { useRouter } from 'vue-router'
import {
  estimateChatContextUsage,
  type ChatContextUsage,
} from '@/ai/estimateChatContextUsage'
import { chatTokenCalibrator } from '@/ai/estimateTokens'
import { api, type ChatRecord, type ChatSummary } from '@/api/client'
import ContextUsageIndicator from '@/components/ContextUsageIndicator.vue'
import { uiConfirm } from '@/composables/useConfirm'
import { isToolAutoAllowed } from '@/composables/useAiToolPermissions'
import { useAuthStore } from '@/stores/auth'
import type { UniverApi } from '@/univer/useUniver'
import {
  describeImpact,
  isDocumentWriteTool,
  needsApprove,
  runOfficeTool,
  toolsForMode,
  type ChatMode,
  type ImpactSummary,
} from '@/univer/officeTools'

type ToolCall = { id: string; name: string; arguments: string }
type Plan = {
  name: string
  overview: string
  todos: Array<{ id: string; content: string; status?: string }>
}
type UiMsg =
  | { key: string; role: 'user'; content: string; wire: string }
  | { key: string; role: 'assistant'; content: string; toolCalls?: ToolCall[] }
  | {
      key: string
      role: 'tool'
      name: string
      content: string
      full: string
      status: 'ok' | 'error' | 'pending'
      toolCallId: string
    }

type ModelOption = { providerId: string; provider: string; model: string }

type SessionLogEntry = {
  ts: string
  chatId: string
  chatTitle: string
  kind: 'user' | 'assistant' | 'tool' | 'system'
  name?: string
  content: string
  status?: 'ok' | 'error' | 'pending'
}

type PlanRunState = 'idle' | 'running' | 'done'

type Tab = {
  id: string
  title: string
  mode: ChatMode
  providerId: string
  model: string
  draft: string
  attachments: Array<{ name: string; text: string }>
  includeDoc: boolean
  messages: UiMsg[]
  plan: Plan | null
  planTitle: string | null
  planMarkdown: string | null
  planEdited: boolean
  planRunState: PlanRunState
  streaming: boolean
  tokenRate: number | null
  tokenMeta: string | null
  showCheckpoint: boolean
  wroteSinceCheckpoint: boolean
  canRetry: boolean
  retryReason: string | null
  executingPlan: boolean
  abort: AbortController | null
  allowedTools: Set<string>
  pendingApprove: {
    name: string
    args: Record<string, unknown>
    summary: ImpactSummary
    resolve: (ok: boolean) => void
  } | null
}

type LoopStopReason =
  | 'rejected'
  | 'aborted'
  | 'editor'
  | 'tools'
  | 'stream'
  | 'loop'
  | null

function retryLabel(reason: LoopStopReason): string | null {
  if (reason === 'rejected') return 'Отклонено'
  if (reason === 'aborted') return 'Остановлено'
  if (reason === 'editor') return 'Редактор не готов'
  if (reason === 'tools') return 'Стоп по ошибкам tools'
  if (reason === 'stream') return 'Ошибка сети'
  if (reason === 'loop') return 'Зацикливание tools'
  return null
}

const props = defineProps<{
  collapsed: boolean
  docType?: 'SHEET' | 'DOC' | null
  docName?: string
  documentId?: string
  univerApi?: UniverApi | null
  getSelectionContext?: () => string
}>()

const emit = defineEmits<{
  toggle: []
  checkpoint: []
  restoreCheckpoint: []
}>()

const WIDTH_KEY = 'office_chat_width'
const auth = useAuthStore()
const router = useRouter()
const tabs = ref<Tab[]>([])
const activeId = ref('')
const width = ref(readWidth())
const modeMenu = ref(false)
const modelMenu = ref(false)
const modelOptions = ref<ModelOption[]>([])
const modelMenuLoading = ref(false)
const modelMenuError = ref('')
const planOpen = ref(false)
const planDraft = ref('')
const fileInput = ref<HTMLInputElement | null>(null)
const composerEl = ref<HTMLTextAreaElement | null>(null)
const transcriptEl = ref<HTMLElement | null>(null)
const modeWrap = ref<HTMLElement | null>(null)
const modelWrap = ref<HTMLElement | null>(null)
const historyBtn = ref<HTMLElement | null>(null)
const historyPanel = ref<HTMLElement | null>(null)
const planModal = ref<HTMLElement | null>(null)
const attachError = ref('')
const sessionLog = ref<SessionLogEntry[]>([])
const logCopied = ref(false)
const historyOpen = ref(false)
const historyItems = ref<ChatSummary[]>([])
const historyLoading = ref(false)
const historyError = ref('')
const renamingId = ref('')
const renameDraft = ref('')
const expandedTools = ref<Set<string>>(new Set())
const approveExpanded = ref(false)
const copiedKey = ref('')
const streamHtmlKey = ref('')
const streamHtml = ref('')
const contextUsage = ref<ChatContextUsage | null>(null)
let logCopiedTimer: ReturnType<typeof setTimeout> | null = null
let copiedTimer: ReturnType<typeof setTimeout> | null = null
let streamHtmlTimer: ReturnType<typeof setTimeout> | null = null
let contextUsageTimer: ReturnType<typeof setTimeout> | null = null
let lastTokenRateAt = 0

const STREAM_HTML_MS = 80
const TOKEN_RATE_MS = 200
const CONTEXT_USAGE_MS = 400

const attachHelpText =
  'Можно прикрепить файл (скрепка): Excel, Word, PowerPoint, PDF, TXT, MD, CSV, JSON. Текст файла попадёт в контекст чата (до ~20 тыс. символов).'

const hasProvider = computed(() => Boolean(auth.me?.hasAiProvider))
const providers = computed(() => auth.me?.providers?.filter((p) => p.enabled) ?? [])
const active = computed(() => tabs.value.find((t) => t.id === activeId.value) ?? null)
const hasMessages = computed(() => Boolean(active.value?.messages.length))
const hasSessionLog = computed(() => sessionLog.value.length > 0)
const panelStyle = computed(() =>
  props.collapsed ? undefined : { width: `${width.value}px` },
)
const planOverview = computed(() => {
  const tab = active.value
  if (!tab?.planMarkdown) return ''
  if (tab.plan?.overview) return tab.plan.overview.slice(0, 160)
  const body = tab.planMarkdown.replace(/^#.*\n+/, '')
  return body.split('\n').filter(Boolean).slice(0, 2).join(' ').slice(0, 160)
})
const composerPlaceholder = computed(() => {
  const mode = active.value?.mode
  if (mode === 'plan') return 'Опишите задачу для плана…'
  if (mode === 'chat') return 'Спросите что угодно…'
  return 'Что сделать с документом?'
})
const emptyHints = computed(() => {
  const mode = active.value?.mode || 'agent'
  const isDoc = props.docType === 'DOC'
  if (mode === 'plan') {
    return isDoc
      ? [
          'Составь план правки выделенного фрагмента',
          'Разбей задачу на шаги без правок файла',
          'Какие риски у глобальной замены текста?',
        ]
      : [
          'Составь план очистки и нормализации таблицы',
          'Разбей задачу на шаги без правок файла',
          'Какие риски у массовой замены значений?',
        ]
  }
  if (mode === 'chat') {
    return isDoc
      ? [
          'Объясни структуру текущего документа',
          'Что сейчас выделено?',
          'Кратко перескажи начало текста',
        ]
      : [
          'Объясни структуру текущего файла',
          'Что в текущем выделении?',
          'Найди пустые ячейки на активном листе',
        ]
  }
  return isDoc
    ? [
        'Сделай выделенный заголовок жирным',
        'Вставь абзац в место курсора',
        'Замени выделенный текст на более короткий',
      ]
    : [
        'Сделай первую строку жирной и подгони ширину колонок',
        'Объедини A1:C1 и добавь внешнюю рамку',
        'Отформатируй заголовки: жирный, заливка, по центру',
      ]
})
const approveSummary = computed(() => {
  const pending = active.value?.pendingApprove
  if (!pending) return ''
  const { title, detail } = pending.summary
  return detail ? `${title}: ${detail}` : title
})

watch(
  () => active.value?.messages.length,
  async () => {
    await nextTick()
    const el = transcriptEl.value
    if (el) el.scrollTop = el.scrollHeight
  },
)

function resetTabs() {
  for (const tab of tabs.value) {
    tab.abort?.abort()
    tab.pendingApprove?.resolve(false)
  }
  tabs.value = []
  activeId.value = ''
  sessionLog.value = []
  logCopied.value = false
  historyOpen.value = false
  historyItems.value = []
  historyError.value = ''
  renamingId.value = ''
  renameDraft.value = ''
  expandedTools.value = new Set()
  approveExpanded.value = false
  copiedKey.value = ''
  clearStreamHtml()
  lastTokenRateAt = 0
  contextUsage.value = null
  if (logCopiedTimer) {
    clearTimeout(logCopiedTimer)
    logCopiedTimer = null
  }
  if (copiedTimer) {
    clearTimeout(copiedTimer)
    copiedTimer = null
  }
  if (contextUsageTimer) {
    clearTimeout(contextUsageTimer)
    contextUsageTimer = null
  }
}

watch(
  () => active.value?.pendingApprove,
  () => {
    approveExpanded.value = false
  },
)

function refreshContextUsage() {
  const tab = active.value
  if (!tab || !tab.model) {
    contextUsage.value = null
    return
  }
  const hasDocument = Boolean(props.documentId || props.docName)
  const wireMessages = trimWireForProvider(toWire(tab))
  contextUsage.value = estimateChatContextUsage({
    mode: tab.mode,
    model: tab.model,
    wireMessages,
    draft: tab.draft,
    attachments: tab.attachments,
    includeDoc: tab.includeDoc,
    docName: props.docName,
    docType: props.docType,
    selectionContext: tab.includeDoc
      ? props.getSelectionContext?.()
      : undefined,
    buildingPlan: tab.executingPlan,
    hasDocument,
    calibrator: chatTokenCalibrator,
  })
}

watch(
  () => {
    const tab = active.value
    if (!tab) return null
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
      props.docName,
      props.docType,
      props.documentId,
    ] as const
  },
  () => {
    if (contextUsageTimer) clearTimeout(contextUsageTimer)
    contextUsageTimer = setTimeout(() => {
      contextUsageTimer = null
      refreshContextUsage()
    }, CONTEXT_USAGE_MS)
  },
  { immediate: true },
)

function pushSessionLog(
  tab: Tab | null | undefined,
  entry: Omit<SessionLogEntry, 'ts' | 'chatId' | 'chatTitle'> & {
    chatId?: string
    chatTitle?: string
  },
) {
  sessionLog.value.push({
    ts: new Date().toISOString(),
    chatId: entry.chatId ?? tab?.id ?? '',
    chatTitle: entry.chatTitle ?? tab?.title ?? '',
    kind: entry.kind,
    name: entry.name,
    content: entry.content,
    status: entry.status,
  })
}

function formatSessionLog() {
  const header = [
    '=== Office session log ===',
    `exportedAt: ${new Date().toISOString()}`,
    `documentId: ${props.documentId || '-'}`,
    `docName: ${props.docName || '-'}`,
    `docType: ${props.docType || '-'}`,
    `entries: ${sessionLog.value.length}`,
    '',
  ].join('\n')
  const body = sessionLog.value
    .map((entry) => {
      const meta = [
        entry.ts,
        entry.kind,
        entry.name || '',
        entry.status || '',
        `chat=${entry.chatTitle || entry.chatId || '-'}`,
      ]
        .filter(Boolean)
        .join(' | ')
      return `[${meta}]\n${entry.content}`
    })
    .join('\n\n---\n\n')
  return `${header}${body}`
}

async function copySessionLog() {
  if (!sessionLog.value.length) return
  try {
    await navigator.clipboard.writeText(formatSessionLog())
    logCopied.value = true
    if (logCopiedTimer) clearTimeout(logCopiedTimer)
    logCopiedTimer = setTimeout(() => {
      logCopied.value = false
      logCopiedTimer = null
    }, 1500)
  } catch {
    attachError.value = 'Не удалось скопировать лог'
  }
}

function toolResultStatus(result: string): 'ok' | 'error' {
  try {
    const parsed = JSON.parse(result) as { error?: unknown }
    if (parsed && typeof parsed === 'object' && parsed.error != null) return 'error'
  } catch {
    /* plain text result */
  }
  return 'ok'
}

watch(hasProvider, (ok) => {
  if (ok && props.documentId && !tabs.value.length) void loadTabs()
})

watch(
  () => props.documentId,
  (id, prev) => {
    if (id === prev) return
    resetTabs()
    if (id) void loadTabs()
  },
)

function closeFloatingOnOutside(event: PointerEvent) {
  const target = event.target
  if (!(target instanceof Node)) return

  if (modeMenu.value && !modeWrap.value?.contains(target)) {
    modeMenu.value = false
  }
  if (modelMenu.value && !modelWrap.value?.contains(target)) {
    modelMenu.value = false
  }
  if (
    historyOpen.value &&
    !historyPanel.value?.contains(target) &&
    !historyBtn.value?.contains(target)
  ) {
    historyOpen.value = false
  }
  if (planOpen.value && !planModal.value?.contains(target)) {
    planOpen.value = false
  }
}

function onFloatingKeydown(event: KeyboardEvent) {
  if (event.key !== 'Escape') return
  if (planOpen.value) {
    planOpen.value = false
    return
  }
  if (modeMenu.value || modelMenu.value || historyOpen.value) {
    modeMenu.value = false
    modelMenu.value = false
    historyOpen.value = false
  }
}

const floatingUiOpen = computed(
  () => modeMenu.value || modelMenu.value || historyOpen.value || planOpen.value,
)

watch(
  floatingUiOpen,
  (open, _prev, onCleanup) => {
    if (!open) return
    document.addEventListener('pointerdown', closeFloatingOnOutside, true)
    document.addEventListener('keydown', onFloatingKeydown)
    onCleanup(() => {
      document.removeEventListener('pointerdown', closeFloatingOnOutside, true)
      document.removeEventListener('keydown', onFloatingKeydown)
    })
  },
  { flush: 'post' },
)

onMounted(() => {
  if (props.documentId) void loadTabs()
})
onBeforeUnmount(() => {
  if (contextUsageTimer) {
    clearTimeout(contextUsageTimer)
    contextUsageTimer = null
  }
  document.removeEventListener('pointerdown', closeFloatingOnOutside, true)
  document.removeEventListener('keydown', onFloatingKeydown)
  resetTabs()
})

function readWidth() {
  const n = Number(localStorage.getItem(WIDTH_KEY))
  return Number.isFinite(n) && n >= 320 ? n : 420
}

function renderMd(text: string) {
  const raw = marked.parse(text || '', { async: false }) as string
  const withCopy = raw.replace(
    /<pre><code([^>]*)>([\s\S]*?)<\/code><\/pre>/g,
    (_match, attrs: string, code: string) =>
      `<div class="code-block"><button type="button" class="copy-code" data-copy="1">Копировать</button><pre><code${attrs}>${code}</code></pre></div>`,
  )
  return DOMPurify.sanitize(withCopy, {
    ADD_ATTR: ['data-copy'],
    ADD_TAGS: ['button'],
  })
}

async function copyText(text: string, key: string) {
  try {
    await navigator.clipboard.writeText(text)
    copiedKey.value = key
    if (copiedTimer) clearTimeout(copiedTimer)
    copiedTimer = setTimeout(() => {
      copiedKey.value = ''
      copiedTimer = null
    }, 1500)
  } catch {
    attachError.value = 'Не удалось скопировать'
  }
}

function onTranscriptClick(event: MouseEvent) {
  const target = event.target as HTMLElement | null
  const btn = target?.closest?.('.copy-code') as HTMLElement | null
  if (!btn) return
  const pre = btn.parentElement?.querySelector('code')
  const text = pre?.textContent || ''
  void copyText(text, `code:${text.slice(0, 24)}`)
  btn.textContent = 'Скопировано'
  setTimeout(() => {
    btn.textContent = 'Копировать'
  }, 1200)
}

function isToolOpen(message: Extract<UiMsg, { role: 'tool' }>) {
  if (message.status === 'error' || message.status === 'pending') return true
  return expandedTools.value.has(message.key)
}

function toggleTool(key: string) {
  const next = new Set(expandedTools.value)
  if (next.has(key)) next.delete(key)
  else next.add(key)
  expandedTools.value = next
}

function applyHint(text: string) {
  const tab = active.value
  if (!tab || tab.streaming) return
  tab.draft = text
}

function formatHistoryDate(value: string) {
  try {
    return new Date(value).toLocaleString('ru-RU', {
      day: '2-digit',
      month: 'short',
      hour: '2-digit',
      minute: '2-digit',
    })
  } catch {
    return value
  }
}

function visibleUser(text: string) {
  const idx = text.indexOf('\n\n---\n')
  return idx === -1 ? text : text.slice(0, idx)
}

function asMode(value: string): ChatMode {
  if (value === 'chat' || value === 'plan' || value === 'agent') return value
  return 'agent'
}

function modeLabel(mode: ChatMode) {
  if (mode === 'plan') return 'План'
  if (mode === 'chat') return 'Чат'
  return 'Агент'
}

function compact(n: number) {
  if (n >= 1000) return `${Math.max(1, Math.round(n / 1000))}K`
  return String(Math.max(0, Math.round(n)))
}

function planToMarkdown(plan: Plan) {
  const todos = plan.todos
    .map((todo) => {
      const done = todo.status === 'done' || todo.status === 'completed'
      return `- [${done ? 'x' : ' '}] ${todo.content}`
    })
    .join('\n')
  return `# ${plan.name}\n\n${plan.overview}\n\n${todos}`.trim()
}

function parsePlanMarkdown(markdown: string): Plan {
  const lines = markdown.replace(/\r\n/g, '\n').split('\n')
  let name = 'План'
  let i = 0
  if (lines[0]?.startsWith('# ')) {
    name = lines[0].slice(2).trim() || 'План'
    i = 1
    while (i < lines.length && !lines[i].trim()) i += 1
  }
  const overviewLines: string[] = []
  while (i < lines.length && !/^-\s+\[[ xX]\]\s+/.test(lines[i])) {
    overviewLines.push(lines[i])
    i += 1
  }
  const todos: Plan['todos'] = []
  while (i < lines.length) {
    const match = lines[i].match(/^-\s+\[([ xX])\]\s+(.+)$/)
    if (match) {
      const done = match[1].toLowerCase() === 'x'
      todos.push({
        id: String(todos.length + 1),
        content: match[2].trim(),
        ...(done ? { status: 'completed' } : {}),
      })
    }
    i += 1
  }
  return {
    name,
    overview: overviewLines.join('\n').trim(),
    todos,
  }
}

function planAllTodosDone(plan: Plan | null): boolean {
  if (!plan?.todos.length) return false
  return plan.todos.every((t) => t.status === 'done' || t.status === 'completed')
}

function titleFromMarkdown(markdown: string, fallback: string) {
  const heading = markdown.match(/^#\s+(.+)$/m)
  return (heading?.[1] || fallback).trim() || 'План'
}

function shortTitle(text: string) {
  const clean = text.replace(/\s+/g, ' ').trim()
  if (!clean) return 'Новый чат'
  return clean.length > 22 ? `${clean.slice(0, 22)}…` : clean
}

function tabFromRecord(record: ChatRecord): Tab {
  const providerId = record.providerId || providers.value[0]?.id || ''
  const providerModel =
    providers.value.find((item) => item.id === providerId)?.model ||
    providers.value[0]?.model ||
    ''
  const plan = record.planMarkdown ? parsePlanMarkdown(record.planMarkdown) : null
  return {
    id: record.id,
    title: record.title || 'Новый чат',
    mode: asMode(record.mode),
    providerId,
    model: record.model || providerModel || '',
    draft: '',
    attachments: [],
    includeDoc: true,
    messages: record.messages.map(messageFromRecord),
    plan,
    planTitle: record.planTitle,
    planMarkdown: record.planMarkdown,
    planEdited: false,
    planRunState: planAllTodosDone(plan) ? 'done' : 'idle',
    streaming: false,
    tokenRate: null,
    tokenMeta: null,
    showCheckpoint: false,
    wroteSinceCheckpoint: false,
    canRetry: false,
    retryReason: null,
    executingPlan: false,
    abort: null,
    allowedTools: new Set(),
    pendingApprove: null,
  }
}

function messageFromRecord(row: ChatRecord['messages'][number]): UiMsg {
  if (row.role === 'tool') {
    return {
      key: row.id,
      role: 'tool',
      name: row.toolName || 'tool',
      content: row.content.slice(0, 500),
      full: row.content,
      status: row.toolStatus === 'error' ? 'error' : 'ok',
      toolCallId: row.toolCallId || '',
    }
  }
  if (row.role === 'assistant') {
    let toolCalls: ToolCall[] | undefined
    if (row.toolCallsJson) {
      try {
        toolCalls = JSON.parse(row.toolCallsJson) as ToolCall[]
      } catch {
        toolCalls = undefined
      }
    }
    return { key: row.id, role: 'assistant', content: row.content, toolCalls }
  }
  return {
    key: row.id,
    role: 'user',
    content: visibleUser(row.content),
    wire: row.content,
  }
}

async function loadTabs() {
  if (!hasProvider.value || !props.documentId) return
  const forDoc = props.documentId
  try {
    const rows = await api.listOpenChats(forDoc)
    if (props.documentId !== forDoc) return
    tabs.value = rows.map(tabFromRecord)
    activeId.value = tabs.value[0]?.id || ''
    await ensureDefaultModel()
  } catch (e) {
    if (props.documentId !== forDoc) return
    attachError.value = e instanceof Error ? e.message : 'Не удалось открыть чаты'
  }
}

function findTab(id: string) {
  return tabs.value.find((tab) => tab.id === id)
}

async function addTab() {
  if (!props.documentId) return
  const created = await api.createChat({
    mode: active.value?.mode || 'agent',
    providerId: active.value?.providerId || providers.value[0]?.id,
    model: active.value?.model || undefined,
    documentId: props.documentId,
  })
  tabs.value.push(tabFromRecord(created))
  activeId.value = created.id
  await ensureDefaultModel()
}

async function closeTab(id: string) {
  const tab = findTab(id)
  if (!tab) return
  tab.abort?.abort()
  tab.pendingApprove?.resolve(false)
  if (renamingId.value === id) {
    renamingId.value = ''
    renameDraft.value = ''
  }
  try {
    const result = await api.closeChat(id)
    tabs.value = tabs.value.filter((item) => item.id !== id)
    if (result.chat) tabs.value.push(tabFromRecord(result.chat))
    if (!tabs.value.some((item) => item.id === activeId.value)) {
      activeId.value = tabs.value[0]?.id || ''
    }
    if (historyOpen.value) void loadHistory()
    await ensureDefaultModel()
  } catch (e) {
    attachError.value = e instanceof Error ? e.message : 'Не удалось закрыть чат'
  }
}

async function loadHistory() {
  if (!props.documentId) return
  historyLoading.value = true
  historyError.value = ''
  try {
    historyItems.value = await api.listClosedChats(props.documentId)
  } catch (e) {
    historyError.value = e instanceof Error ? e.message : 'Не удалось загрузить историю'
  } finally {
    historyLoading.value = false
  }
}

async function toggleHistory() {
  historyOpen.value = !historyOpen.value
  modeMenu.value = false
  modelMenu.value = false
  if (historyOpen.value) await loadHistory()
}

async function reopenChat(id: string) {
  try {
    const record = await api.updateChat(id, { pinnedOpen: true })
    if (!tabs.value.some((tab) => tab.id === record.id)) {
      tabs.value.push(tabFromRecord(record))
    }
    activeId.value = record.id
    historyOpen.value = false
    historyItems.value = historyItems.value.filter((item) => item.id !== id)
    await ensureDefaultModel()
  } catch (e) {
    historyError.value = e instanceof Error ? e.message : 'Не удалось открыть чат'
  }
}

async function deleteFromHistory(id: string) {
  const item = historyItems.value.find((row) => row.id === id)
  const ok = await uiConfirm({
    title: 'Удалить чат?',
    message: item?.title
      ? `Чат «${item.title}» будет удалён безвозвратно.`
      : 'Чат будет удалён безвозвратно.',
    confirmLabel: 'Удалить',
  })
  if (!ok) return
  try {
    await api.deleteChat(id)
    historyItems.value = historyItems.value.filter((row) => row.id !== id)
  } catch (e) {
    historyError.value = e instanceof Error ? e.message : 'Не удалось удалить чат'
  }
}

function startRename(tab: Tab) {
  renamingId.value = tab.id
  renameDraft.value = tab.title
  void nextTick(() => {
    const input = document.querySelector(
      `.tab-rename[data-id="${tab.id}"]`,
    ) as HTMLInputElement | null
    input?.focus()
    input?.select()
  })
}

function cancelRename() {
  renamingId.value = ''
  renameDraft.value = ''
}

async function commitRename(tab: Tab) {
  if (renamingId.value !== tab.id) return
  const next = renameDraft.value.replace(/\s+/g, ' ').trim() || 'Новый чат'
  renamingId.value = ''
  renameDraft.value = ''
  if (next === tab.title) return
  tab.title = next
  await patchTab(tab, { title: next })
}

function onRenameKey(event: KeyboardEvent, tab: Tab) {
  if (event.key === 'Enter') {
    event.preventDefault()
    void commitRename(tab)
  }
  if (event.key === 'Escape') {
    event.preventDefault()
    cancelRename()
  }
}

async function patchTab(tab: Tab, body: Record<string, unknown>) {
  try {
    await api.updateChat(tab.id, body)
  } catch (e) {
    console.error(e)
  }
}

function setMode(mode: ChatMode) {
  const tab = active.value
  modeMenu.value = false
  if (!tab || tab.mode === mode) return
  tab.mode = mode
  void patchTab(tab, { mode })
}

async function toggleModelMenu() {
  modeMenu.value = false
  historyOpen.value = false
  modelMenu.value = !modelMenu.value
  if (!modelMenu.value) return
  await loadModelOptions()
}

async function loadModelOptions() {
  modelMenuLoading.value = true
  modelMenuError.value = ''
  try {
    const errors: string[] = []
    const rows = await Promise.all(
      providers.value.map(async (provider) => {
        try {
          const res = await api.listProviderModels(provider.id)
          return res.models.map((model) => ({
            providerId: provider.id,
            provider: provider.provider,
            model,
          }))
        } catch (e) {
          const msg = e instanceof Error ? e.message : 'Ошибка загрузки моделей'
          errors.push(
            providers.value.length > 1 ? `${provider.provider}: ${msg}` : msg,
          )
          return [] as ModelOption[]
        }
      }),
    )
    modelOptions.value = rows.flat()
    if (!modelOptions.value.length) {
      modelMenuError.value =
        errors[0] || 'Не удалось загрузить модели'
    }
  } finally {
    modelMenuLoading.value = false
  }
}

function setModel(option: ModelOption) {
  const tab = active.value
  modelMenu.value = false
  if (!tab) return
  tab.providerId = option.providerId
  tab.model = option.model
  void patchTab(tab, { providerId: option.providerId, model: option.model })
}

async function ensureDefaultModel() {
  const tab = active.value
  if (!tab || tab.model || !providers.value.length) return
  if (!modelOptions.value.length) await loadModelOptions()
  const first = modelOptions.value[0]
  if (first) setModel(first)
}

function modelLabel(tab: Tab) {
  if (tab.model) return tab.model
  return 'модель'
}

function pickAttach() {
  fileInput.value?.click()
}

async function onAttach(e: Event) {
  const inputEl = e.target as HTMLInputElement
  const file = inputEl.files?.[0]
  inputEl.value = ''
  const tab = active.value
  if (!file || !tab) return
  attachError.value = ''
  try {
    const extracted = await api.extractAttachment(file)
    tab.attachments.push(extracted)
  } catch (err) {
    attachError.value = err instanceof Error ? err.message : 'Не удалось прочитать файл'
  }
}

function removeAttach(index: number) {
  active.value?.attachments.splice(index, 1)
}

function clearStreamHtml() {
  if (streamHtmlTimer) {
    clearTimeout(streamHtmlTimer)
    streamHtmlTimer = null
  }
  streamHtmlKey.value = ''
  streamHtml.value = ''
}

function flushStreamHtml(key: string, content: string) {
  streamHtmlKey.value = key
  streamHtml.value = renderMd(content)
}

function scheduleStreamHtml(key: string, content: string) {
  if (!streamHtmlTimer) {
    streamHtmlTimer = setTimeout(() => {
      streamHtmlTimer = null
      flushStreamHtml(key, content)
    }, STREAM_HTML_MS)
    return
  }
  // Keep latest content for the pending flush
  const pendingKey = key
  const pendingContent = content
  clearTimeout(streamHtmlTimer)
  streamHtmlTimer = setTimeout(() => {
    streamHtmlTimer = null
    flushStreamHtml(pendingKey, pendingContent)
  }, STREAM_HTML_MS)
}

function assistantHtml(message: Extract<UiMsg, { role: 'assistant' }>) {
  if (message.key === streamHtmlKey.value && streamHtml.value) {
    return streamHtml.value
  }
  return renderMd(message.content)
}

function askApprove(tab: Tab, name: string, args: Record<string, unknown>) {
  return new Promise<boolean>((resolve) => {
    if (tab.abort?.signal.aborted) {
      resolve(false)
      return
    }
    let settled = false
    const finish = (ok: boolean) => {
      if (settled) return
      settled = true
      tab.abort?.signal.removeEventListener('abort', onAbort)
      if (tab.pendingApprove) tab.pendingApprove = null
      approveExpanded.value = false
      resolve(ok)
    }
    const onAbort = () => finish(false)
    tab.abort?.signal.addEventListener('abort', onAbort, { once: true })
    const summary = describeImpact(name, args, props.univerApi ?? null)
    tab.pendingApprove = { name, args, summary, resolve: finish }
  })
}

function resolveApprove(ok: boolean, always = false) {
  const tab = active.value
  if (tab?.pendingApprove) {
    if (always && ok) tab.allowedTools.add(tab.pendingApprove.name)
    pushSessionLog(tab, {
      kind: 'system',
      name: 'approve',
      content: JSON.stringify({
        tool: tab.pendingApprove.name,
        args: tab.pendingApprove.args,
        decision: always && ok ? 'always' : ok ? 'accepted' : 'rejected',
      }),
      status: ok ? 'ok' : 'error',
    })
  }
  tab?.pendingApprove?.resolve(ok)
  if (tab) tab.pendingApprove = null
  approveExpanded.value = false
}

function openPlan() {
  if (!active.value?.planMarkdown) return
  planDraft.value = active.value.planMarkdown
  planOpen.value = true
}

async function savePlan() {
  const tab = active.value
  if (!tab) return
  tab.planMarkdown = planDraft.value
  tab.planTitle = titleFromMarkdown(planDraft.value, tab.planTitle || 'План')
  tab.plan = parsePlanMarkdown(planDraft.value)
  tab.planEdited = true
  tab.planRunState = 'idle'
  planOpen.value = false
  await patchTab(tab, { planTitle: tab.planTitle, planMarkdown: tab.planMarkdown })
}

async function runPlan() {
  const tab = active.value
  if (!tab?.planMarkdown || tab.streaming) return
  if (!tab.plan) tab.plan = parsePlanMarkdown(tab.planMarkdown)
  const title = tab.planTitle || tab.plan.name || 'План'
  tab.mode = 'agent'
  tab.executingPlan = true
  tab.planRunState = 'running'
  await patchTab(tab, { mode: 'agent' })
  try {
    await send({
      display: `Сделать по плану «${title}»`,
      wire: `Сделать по плану «${title}»\n\n---\n${tab.planMarkdown}`,
    })
    const stopped =
      tab.retryReason === 'Остановлено' || tab.retryReason === 'Отклонено'
    if (stopped) {
      tab.planRunState = planAllTodosDone(tab.plan) ? 'done' : 'idle'
    } else {
      tab.planRunState = 'done'
    }
  } finally {
    tab.executingPlan = false
    if (tab.planRunState === 'running') {
      tab.planRunState = planAllTodosDone(tab.plan) ? 'done' : 'idle'
    }
  }
}

function goSettings() {
  void router.push('/home')
  auth.openSettings()
}

async function send(preset?: { display: string; wire: string }) {
  const tab = active.value
  if (!tab || tab.streaming || !hasProvider.value) return
  const display = (preset?.display ?? tab.draft).trim()
  if (!display) return
  if (!tab.model) {
    attachError.value = 'Выберите модель в чате ассистента'
    void toggleModelMenu()
    return
  }
  const contextParts: string[] = []
  if (tab.includeDoc && props.docName) {
    contextParts.push(`Текущий документ: ${props.docName} (${props.docType})`)
    const sel = props.getSelectionContext?.()
    if (sel) contextParts.push(sel)
  }
  if (!preset) {
    for (const file of tab.attachments) {
      contextParts.push(`Файл ${file.name}:\n${file.text}`)
    }
  }
  const contextBlock = contextParts.length
    ? `\n\n---\nКонтекст:\n${contextParts.join('\n\n')}`
    : ''
  const wire = preset?.wire
    ? `${preset.wire}${tab.includeDoc && contextParts.length ? contextBlock : ''}`
    : contextParts.length
      ? `${display}${contextBlock}`
      : display

  tab.messages.push({ key: crypto.randomUUID(), role: 'user', content: display, wire })
  if (tab.title === 'Новый чат') {
    tab.title = shortTitle(display)
    void patchTab(tab, { title: tab.title })
  }
  pushSessionLog(tab, { kind: 'user', content: wire })
  tab.draft = ''
  tab.attachments = []
  void nextTick(resizeComposer)
  tab.streaming = true
  tab.tokenRate = null
  tab.tokenMeta = null
  tab.showCheckpoint = false
  tab.wroteSinceCheckpoint = false
  tab.canRetry = false
  tab.retryReason = null
  clearStreamHtml()
  lastTokenRateAt = 0
  const controller = new AbortController()
  tab.abort = controller
  if (props.univerApi) emit('checkpoint')
  void api.addChatMessage(tab.id, { role: 'user', content: wire })

  try {
    const reason = await agentLoop(tab.id)
    const label = retryLabel(reason)
    if (label) {
      tab.canRetry = true
      tab.retryReason = label
    }
  } catch (e) {
    const aborted = e instanceof DOMException && e.name === 'AbortError'
    if (!aborted) {
      const content = e instanceof Error ? e.message : 'Ошибка'
      tab.messages.push({
        key: crypto.randomUUID(),
        role: 'assistant',
        content,
      })
      pushSessionLog(tab, { kind: 'system', name: 'error', content, status: 'error' })
      tab.canRetry = true
      tab.retryReason = 'Ошибка сети'
    } else {
      pushSessionLog(tab, {
        kind: 'system',
        name: 'abort',
        content: 'Остановлено пользователем',
        status: 'error',
      })
      tab.canRetry = true
      tab.retryReason = 'Остановлено'
    }
  } finally {
    tab.showCheckpoint = tab.wroteSinceCheckpoint && Boolean(props.univerApi)
    tab.streaming = false
    tab.abort = null
    clearStreamHtml()
  }
}

async function retryTurn() {
  const tab = active.value
  if (!tab || tab.streaming || !tab.canRetry) return
  tab.canRetry = false
  tab.retryReason = null
  tab.streaming = true
  tab.tokenRate = null
  tab.tokenMeta = null
  tab.showCheckpoint = false
  clearStreamHtml()
  lastTokenRateAt = 0
  const controller = new AbortController()
  tab.abort = controller
  if (props.univerApi) emit('checkpoint')
  try {
    const reason = await agentLoop(tab.id)
    const label = retryLabel(reason)
    if (label) {
      tab.canRetry = true
      tab.retryReason = label
    }
  } catch (e) {
    const aborted = e instanceof DOMException && e.name === 'AbortError'
    if (!aborted) {
      const content = e instanceof Error ? e.message : 'Ошибка'
      tab.messages.push({
        key: crypto.randomUUID(),
        role: 'assistant',
        content,
      })
      pushSessionLog(tab, { kind: 'system', name: 'error', content, status: 'error' })
      tab.canRetry = true
      tab.retryReason = 'Ошибка сети'
    } else {
      tab.canRetry = true
      tab.retryReason = 'Остановлено'
    }
  } finally {
    tab.showCheckpoint = tab.wroteSinceCheckpoint && Boolean(props.univerApi)
    tab.streaming = false
    tab.abort = null
    clearStreamHtml()
  }
}

function stop() {
  const tab = active.value
  if (!tab) return
  tab.abort?.abort()
  if (tab.pendingApprove) {
    tab.pendingApprove.resolve(false)
    tab.pendingApprove = null
    approveExpanded.value = false
  }
}

async function agentLoop(tabId: string): Promise<LoopStopReason> {
  const tab = findTab(tabId)
  if (!tab) return null
  const signal = tab.abort?.signal
  let steps = 0
  let lastToolSig = ''
  let repeats = 0
  let consecutiveToolErrors = 0

  while (steps < 25) {
    if (signal?.aborted) return 'aborted'
    steps++
    const hasDocument = Boolean(props.documentId || props.docName)
    const docType = props.docType === 'DOC' ? 'DOC' : 'SHEET'
    const tools = toolsForMode(tab.mode, tab.executingPlan, hasDocument, docType)
    if (tools.length && !props.univerApi) {
      const content = 'Редактор ещё загружается, повторите.'
      tab.messages.push({
        key: crypto.randomUUID(),
        role: 'assistant',
        content,
      })
      pushSessionLog(tab, { kind: 'system', name: 'editor', content, status: 'error' })
      return 'editor'
    }
    const started = performance.now()
    let chars = 0
    const assistant: { content: string; toolCalls: ToolCall[] } = { content: '', toolCalls: [] }
    let assistantKey = ''
    let usagePrompt: number | null = null
    let usageCompletion: number | null = null
    const messages = trimWireForProvider(toWire(tab))
    const heuristicPrompt = JSON.stringify(messages).length / 4
    const promptEstimate = estimateChatContextUsage({
      mode: tab.mode,
      model: tab.model,
      wireMessages: messages,
      draft: '',
      attachments: [],
      includeDoc: false,
      buildingPlan: tab.executingPlan,
      hasDocument: Boolean(props.documentId || props.docName),
      docType: props.docType,
      calibrator: chatTokenCalibrator,
    }).usedTokens

    try {
      await streamChat(
        {
          mode: tab.mode,
          messages,
          docType,
          hasDocument,
          buildingPlan: tab.executingPlan,
          providerId: tab.providerId || providers.value[0]?.id,
          model: tab.model || undefined,
        },
        (ev) => {
          if (ev.type === 'delta' && typeof ev.content === 'string') {
            assistant.content += ev.content
            chars += ev.content.length
            const now = performance.now()
            if (now - lastTokenRateAt >= TOKEN_RATE_MS) {
              lastTokenRateAt = now
              const seconds = Math.max(0.2, (now - started) / 1000)
              tab.tokenRate = Math.round(chars / 4 / seconds)
            }
            if (!assistantKey) {
              assistantKey = crypto.randomUUID()
              tab.messages.push({
                key: assistantKey,
                role: 'assistant',
                content: assistant.content,
              })
              flushStreamHtml(assistantKey, assistant.content)
            } else {
              const row = tab.messages.find(
                (item) => item.key === assistantKey && item.role === 'assistant',
              )
              if (row && row.role === 'assistant') row.content = assistant.content
              scheduleStreamHtml(assistantKey, assistant.content)
            }
          }
          if (ev.type === 'usage') {
            usagePrompt =
              typeof ev.promptTokens === 'number' ? ev.promptTokens : usagePrompt
            usageCompletion =
              typeof ev.completionTokens === 'number'
                ? ev.completionTokens
                : usageCompletion
            if (
              typeof ev.completionTokens === 'number' &&
              Number.isFinite(ev.completionTokens)
            ) {
              const seconds = Math.max(0.2, (performance.now() - started) / 1000)
              tab.tokenRate = Math.round(ev.completionTokens / seconds)
            }
          }
          if (ev.type === 'done') {
            assistant.toolCalls = (ev.toolCalls as ToolCall[]) || []
            if (assistant.toolCalls.length && !assistantKey) {
              assistantKey = crypto.randomUUID()
              tab.messages.push({
                key: assistantKey,
                role: 'assistant',
                content: assistant.content,
                toolCalls: assistant.toolCalls,
              })
            }
            if (assistantKey) flushStreamHtml(assistantKey, assistant.content)
          }
          if (ev.type === 'error') throw new Error(String(ev.error || 'AI error'))
        },
        signal ?? new AbortController().signal,
      )
    } catch (error) {
      const aborted =
        (error instanceof DOMException && error.name === 'AbortError') ||
        Boolean(signal?.aborted)
      if (aborted) {
        if (assistantKey) flushStreamHtml(assistantKey, assistant.content)
        await persistAssistant(tab, assistant)
        pushSessionLog(tab, {
          kind: 'system',
          name: 'abort',
          content: 'Стрим остановлен',
          status: 'error',
        })
        return 'aborted'
      }
      throw error
    }

    if (signal?.aborted) return 'aborted'

    const promptTokens = usagePrompt ?? heuristicPrompt
    const completionTokens = usageCompletion ?? assistant.content.length / 4
    if (usagePrompt != null) {
      const rawEstimate = Math.round(
        promptEstimate / Math.max(0.01, chatTokenCalibrator.getFactor()),
      )
      chatTokenCalibrator.observe(rawEstimate, usagePrompt)
      refreshContextUsage()
    }
    const rate = tab.tokenRate ?? 0
    const usageTag =
      usagePrompt != null || usageCompletion != null ? '' : ' ≈'
    tab.tokenMeta = `${compact(promptTokens)} → ${compact(completionTokens)} · всего ${compact(promptTokens + completionTokens)}${usageTag} · ${rate} ток/с`

    if (assistant.toolCalls.length) {
      if (!assistantKey) {
        assistantKey = crypto.randomUUID()
        tab.messages.push({
          key: assistantKey,
          role: 'assistant',
          content: assistant.content,
          toolCalls: assistant.toolCalls,
        })
      } else {
        const row = tab.messages.find((item) => item.key === assistantKey && item.role === 'assistant')
        if (row && row.role === 'assistant') row.toolCalls = assistant.toolCalls
      }
    }
    if (assistant.content || assistant.toolCalls.length) {
      pushSessionLog(tab, {
        kind: 'assistant',
        content: [
          assistant.content,
          assistant.toolCalls.length
            ? `toolCalls: ${JSON.stringify(assistant.toolCalls)}`
            : '',
        ]
          .filter(Boolean)
          .join('\n'),
      })
    }
    await persistAssistant(tab, assistant)
    if (!assistant.toolCalls.length) break

    const sig = JSON.stringify(assistant.toolCalls.map((call) => [call.name, call.arguments]))
    if (sig === lastToolSig) {
      repeats++
      if (repeats >= 2) {
        const content = 'Остановка: повторяющиеся tool-calls.'
        tab.messages.push({
          key: crypto.randomUUID(),
          role: 'assistant',
          content,
        })
        pushSessionLog(tab, { kind: 'system', name: 'loop', content, status: 'error' })
        return 'loop'
      }
    } else {
      repeats = 0
      lastToolSig = sig
    }

    for (const call of assistant.toolCalls) {
      if (signal?.aborted) return 'aborted'
      let args: Record<string, unknown> = {}
      try {
        args = JSON.parse(call.arguments || '{}') as Record<string, unknown>
      } catch {
        args = {}
      }
      const toolKey = crypto.randomUUID()
      tab.messages.push({
        key: toolKey,
        role: 'tool',
        name: call.name,
        content: 'Выполняется…',
        full: '',
        status: 'pending',
        toolCallId: call.id,
      })
      const apiRef = props.univerApi ?? null
      if (
        !tab.allowedTools.has(call.name) &&
        !isToolAutoAllowed(call.name) &&
        needsApprove(call.name, args, apiRef)
      ) {
        const ok = await askApprove(tab, call.name, args)
        if (signal?.aborted) return 'aborted'
        if (!ok) {
          const rejected = 'Отклонено пользователем'
          setToolRow(tab, toolKey, rejected, rejected, 'error')
          pushSessionLog(tab, {
            kind: 'tool',
            name: call.name,
            content: JSON.stringify({ error: 'rejected by user', args }),
            status: 'error',
          })
          await api.addChatMessage(tab.id, {
            role: 'tool',
            content: JSON.stringify({ error: 'rejected by user' }),
            toolName: call.name,
            toolCallId: call.id,
            toolStatus: 'error',
          })
          return 'rejected'
        }
      }
      if (signal?.aborted) return 'aborted'
      let result: string
      try {
        result = await runOfficeTool(call.name, args, {
          api: apiRef,
          docType: props.docType === 'DOC' ? 'DOC' : 'SHEET',
          plan: tab.plan,
          setPlan: (plan) => {
            tab.plan = plan
          },
        })
      } catch (error) {
        result = JSON.stringify({
          error: error instanceof Error ? error.message : 'tool failed',
        })
      }
      if ((call.name === 'create_plan' || call.name === 'update_plan') && tab.plan) {
        if (call.name === 'create_plan') {
          tab.planEdited = false
          tab.planRunState = 'idle'
          tab.planTitle = tab.plan.name
          tab.planMarkdown = planToMarkdown(tab.plan)
          void patchTab(tab, { planTitle: tab.planTitle, planMarkdown: tab.planMarkdown })
        } else if (tab.executingPlan || !tab.planEdited) {
          tab.planTitle = tab.plan.name
          tab.planMarkdown = planToMarkdown(tab.plan)
          if (planAllTodosDone(tab.plan)) tab.planRunState = 'done'
          void patchTab(tab, { planTitle: tab.planTitle, planMarkdown: tab.planMarkdown })
        }
      }
      const status = toolResultStatus(result)
      setToolRow(tab, toolKey, result.slice(0, 500), result, status)
      if (status !== 'error' && isDocumentWriteTool(call.name)) {
        tab.wroteSinceCheckpoint = true
      }
      pushSessionLog(tab, {
        kind: 'tool',
        name: call.name,
        content: `args: ${JSON.stringify(args)}\nresult: ${result}`,
        status,
      })
      await api.addChatMessage(tab.id, {
        role: 'tool',
        content: result,
        toolName: call.name,
        toolCallId: call.id,
        toolStatus: status,
      })
      if (status === 'error') {
        consecutiveToolErrors += 1
        if (consecutiveToolErrors >= 3) {
          const content =
            'Остановка: несколько tool-вызовов подряд завершились ошибкой. Уточните задачу или проверьте выделение.'
          tab.messages.push({
            key: crypto.randomUUID(),
            role: 'assistant',
            content,
          })
          pushSessionLog(tab, {
            kind: 'system',
            name: 'tool-errors',
            content,
            status: 'error',
          })
          return 'tools'
        }
      } else {
        consecutiveToolErrors = 0
      }
    }
  }
  return null
}

function setToolRow(tab: Tab, key: string, content: string, full: string, status: 'ok' | 'error') {
  const row = tab.messages.find((item) => item.key === key && item.role === 'tool')
  if (row && row.role === 'tool') {
    row.content = content
    row.full = full
    row.status = status
  }
}

async function persistAssistant(tab: Tab, assistant: { content: string; toolCalls: ToolCall[] }) {
  if (!assistant.content && !assistant.toolCalls.length) return
  await api.addChatMessage(tab.id, {
    role: 'assistant',
    content: assistant.content,
    toolCallsJson: assistant.toolCalls.length ? JSON.stringify(assistant.toolCalls) : undefined,
  })
}

const TOOL_WIRE_MAX = 4000
const ARGS_WIRE_MAX = 2000
const MSG_WIRE_MAX = 32_000
const WIRE_MAX_MESSAGES = 60
const WIRE_MAX_TOTAL_CHARS = 140_000

function clipWireText(text: string, max: number): string {
  if (text.length <= max) return text
  return JSON.stringify({
    truncated: true,
    preview: text.slice(0, max),
    originalChars: text.length,
  })
}

function wireContentChars(msg: Record<string, unknown>): number {
  let n = 0
  if (typeof msg.content === 'string') n += msg.content.length
  const toolCalls = msg.tool_calls
  if (Array.isArray(toolCalls)) {
    for (const tc of toolCalls) {
      const args = (tc as { function?: { arguments?: string } })?.function?.arguments
      if (typeof args === 'string') n += args.length
    }
  }
  return n
}

function trimWireForProvider(wire: Array<Record<string, unknown>>) {
  const totalChars = () => wire.reduce((sum, m) => sum + wireContentChars(m), 0)
  if (wire.length <= WIRE_MAX_MESSAGES && totalChars() <= WIRE_MAX_TOTAL_CHARS) {
    return wire
  }

  const turns: Array<Array<Record<string, unknown>>> = []
  let current: Array<Record<string, unknown>> = []
  for (const msg of wire) {
    if (msg.role === 'user' && current.length) {
      turns.push(current)
      current = []
    }
    current.push(msg)
  }
  if (current.length) turns.push(current)

  let dropped = 0
  while (
    turns.length > 1 &&
    (turns.reduce((n, t) => n + t.length, 0) > WIRE_MAX_MESSAGES ||
      turns.reduce((n, t) => n + t.reduce((s, m) => s + wireContentChars(m), 0), 0) >
        WIRE_MAX_TOTAL_CHARS)
  ) {
    dropped += turns.shift()!.length
  }

  const kept = turns.flat()
  if (!dropped) return kept
  return [
    {
      role: 'user',
      content: `[Пропущены ранние сообщения: ${dropped}]`,
    },
    ...kept,
  ]
}

function closeOpenToolCalls(
  wire: Array<Record<string, unknown>>,
  openToolCallIds: Set<string> | null,
) {
  if (!openToolCallIds?.size) return
  for (const id of openToolCallIds) {
    wire.push({
      role: 'tool',
      tool_call_id: id,
      content: JSON.stringify({ error: 'tool result missing' }),
    })
  }
  openToolCallIds.clear()
}

function toWire(tab: Tab) {
  const wire: Array<Record<string, unknown>> = []
  let openToolCallIds: Set<string> | null = null
  for (const message of tab.messages) {
    if (message.role === 'user') {
      closeOpenToolCalls(wire, openToolCallIds)
      openToolCallIds = null
      wire.push({
        role: 'user',
        content: clipWireText(message.wire, MSG_WIRE_MAX),
      })
    }
    if (message.role === 'assistant') {
      closeOpenToolCalls(wire, openToolCallIds)
      const toolCalls = message.toolCalls?.length
        ? message.toolCalls.map((call) => ({
            id: call.id,
            type: 'function' as const,
            function: {
              name: call.name,
              arguments: clipWireText(call.arguments, ARGS_WIRE_MAX),
            },
          }))
        : undefined
      openToolCallIds = toolCalls
        ? new Set(toolCalls.map((call) => call.id))
        : null
      const content = message.content
        ? clipWireText(message.content, MSG_WIRE_MAX)
        : null
      wire.push({
        role: 'assistant',
        content,
        ...(toolCalls ? { tool_calls: toolCalls } : {}),
      })
    }
    if (message.role === 'tool' && message.status !== 'pending') {
      if (!message.toolCallId || !openToolCallIds?.has(message.toolCallId)) continue
      openToolCallIds.delete(message.toolCallId)
      wire.push({
        role: 'tool',
        tool_call_id: message.toolCallId,
        name: message.name,
        content: clipWireText(message.full || message.content, TOOL_WIRE_MAX),
      })
    }
  }
  closeOpenToolCalls(wire, openToolCallIds)
  return wire
}

function parseSsePart(part: string): Record<string, unknown> | null {
  const line = part.trim()
  if (!line.startsWith('data:')) return null
  try {
    return JSON.parse(line.slice(5).trim()) as Record<string, unknown>
  } catch {
    return null
  }
}

async function streamChat(
  body: Record<string, unknown>,
  onEvent: (ev: Record<string, unknown>) => void,
  signal: AbortSignal,
) {
  const res = await fetch('/ai/chat', {
    method: 'POST',
    credentials: 'include',
    signal,
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(body),
  })
  if (!res.ok || !res.body) throw new Error('Не удалось связаться с AI gateway')
  const reader = res.body.getReader()
  const decoder = new TextDecoder()
  let buffer = ''
  try {
    while (true) {
      const { done, value } = await reader.read()
      if (done) break
      buffer += decoder.decode(value, { stream: true })
      const parts = buffer.split('\n\n')
      buffer = parts.pop() || ''
      for (const part of parts) {
        const event = parseSsePart(part)
        if (event) onEvent(event)
      }
    }
    buffer += decoder.decode()
    if (buffer.trim()) {
      const event = parseSsePart(buffer)
      if (event) onEvent(event)
    }
  } finally {
    try {
      await reader.cancel()
    } catch {
      /* ignore */
    }
  }
}

function startResize(event: PointerEvent) {
  if (props.collapsed) return
  event.preventDefault()
  const startX = event.clientX
  const startW = width.value
  const move = (ev: PointerEvent) => {
    const max = Math.max(320, Math.floor(window.innerWidth * 0.5))
    width.value = Math.min(max, Math.max(320, startW + (startX - ev.clientX)))
  }
  const up = () => {
    localStorage.setItem(WIDTH_KEY, String(Math.round(width.value)))
    window.removeEventListener('pointermove', move)
    window.removeEventListener('pointerup', up)
  }
  window.addEventListener('pointermove', move)
  window.addEventListener('pointerup', up)
}

function resizeComposer() {
  const el = composerEl.value
  if (!el) return
  el.style.height = 'auto'
  el.style.height = `${el.scrollHeight}px`
}

function onComposerKey(event: KeyboardEvent) {
  if (event.key !== 'Enter' || event.isComposing) return
  if (event.shiftKey) return
  event.preventDefault()
  void send()
}

watch(
  () => [activeId.value, active.value?.draft] as const,
  async () => {
    await nextTick()
    resizeComposer()
  },
)
</script>

<template>
  <aside class="panel" :class="{ collapsed }" :style="panelStyle">
    <div v-if="!collapsed" class="resize" @pointerdown="startResize" />
    <button v-if="collapsed" type="button" class="expand" @click="emit('toggle')">
      <span class="expand-label">AI</span>
    </button>
    <template v-else>
      <header class="tabs">
        <div class="tab-scroll">
          <div
            v-for="tab in tabs"
            :key="tab.id"
            class="tab"
            :class="{ on: tab.id === activeId }"
            @click="activeId = tab.id"
          >
            <span v-if="tab.streaming" class="spin" />
            <input
              v-if="renamingId === tab.id"
              class="tab-rename"
              :data-id="tab.id"
              v-model="renameDraft"
              @click.stop
              @keydown="onRenameKey($event, tab)"
              @blur="commitRename(tab)"
            />
            <span
              v-else
              class="tab-title"
              :title="tab.title"
              @dblclick.stop="startRename(tab)"
            >{{ tab.title }}</span>
            <button type="button" class="x" title="Закрыть" @click.stop="closeTab(tab.id)">×</button>
          </div>
        </div>
        <div class="tab-actions">
          <button type="button" class="plus" title="Новый чат" @click="addTab">+</button>
          <button
            ref="historyBtn"
            type="button"
            class="icon"
            title="История"
            :class="{ on: historyOpen }"
            @click="toggleHistory"
          >
            <svg viewBox="0 0 24 24" width="16" height="16" aria-hidden="true">
              <path
                d="M4 7h16M4 12h16M4 17h16"
                fill="none"
                stroke="currentColor"
                stroke-width="1.8"
                stroke-linecap="round"
              />
            </svg>
          </button>
          <button
            type="button"
            class="icon"
            :disabled="!hasSessionLog"
            :title="
              logCopied
                ? 'Скопировано'
                : hasSessionLog
                  ? 'Скопировать лог сессии'
                  : 'Лог пуст'
            "
            @click="copySessionLog"
          >
            <svg v-if="!logCopied" viewBox="0 0 24 24" width="16" height="16" aria-hidden="true">
              <rect x="9" y="9" width="11" height="11" rx="2" fill="none" stroke="currentColor" stroke-width="1.7" />
              <path d="M6 15H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h8a2 2 0 0 1 2 2v1" fill="none" stroke="currentColor" stroke-width="1.7" />
            </svg>
            <svg v-else viewBox="0 0 24 24" width="16" height="16" aria-hidden="true">
              <path d="M5 12.5l4.5 4.5L19 7" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" />
            </svg>
          </button>
          <button type="button" class="icon" title="Свернуть" @click="emit('toggle')">
            <svg viewBox="0 0 24 24" width="16" height="16" aria-hidden="true">
              <path d="M9 6l6 6-6 6" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round" />
            </svg>
          </button>
        </div>
      </header>

      <div v-if="historyOpen" ref="historyPanel" class="history-panel">
        <div class="history-head">
          <span>Закрытые чаты</span>
          <button type="button" class="text" @click="historyOpen = false">Закрыть</button>
        </div>
        <p v-if="historyLoading" class="history-hint">Загрузка…</p>
        <p v-else-if="historyError" class="history-hint error">{{ historyError }}</p>
        <p v-else-if="!historyItems.length" class="history-hint">Пока пусто</p>
        <ul v-else class="history-list">
          <li v-for="item in historyItems" :key="item.id">
            <button type="button" class="history-main" @click="reopenChat(item.id)">
              <span class="history-title">{{ item.title || 'Без названия' }}</span>
              <span class="history-meta">{{ formatHistoryDate(item.updatedAt) }} · {{ modeLabel(asMode(item.mode)) }}</span>
            </button>
            <button type="button" class="history-del" title="Удалить" @click="deleteFromHistory(item.id)">×</button>
          </li>
        </ul>
      </div>

      <div v-if="docName" class="doc-context" :title="docName">
        Файл: {{ docName }}
      </div>

      <div v-if="!hasProvider" class="cta">
        <p class="cta-title">Подключите AI-провайдер</p>
        <p class="cta-text">Добавьте ключ в настройках, чтобы ассистент мог работать с документом.</p>
        <button type="button" class="ui-btn ui-btn--primary" @click="goSettings">Открыть настройки</button>
      </div>
      <div v-else-if="!active" class="cta"><p>Открываю чат…</p></div>
      <div v-else class="stack" :class="{ empty: !hasMessages }">
        <div v-if="!hasMessages" class="empty-state">
          <h3>{{ modeLabel(active.mode) }}</h3>
          <p>Напишите задачу или выберите подсказку.</p>
          <p class="attach-help">{{ attachHelpText }}</p>
          <div class="hints">
            <button
              v-for="hint in emptyHints"
              :key="hint"
              type="button"
              class="hint"
              :disabled="active.streaming"
              @click="applyHint(hint)"
            >
              {{ hint }}
            </button>
          </div>
        </div>

        <div
          v-else
          ref="transcriptEl"
          class="transcript"
          @click="onTranscriptClick"
        >
          <div v-for="message in active.messages" :key="message.key" class="msg" :class="message.role">
            <div v-if="message.role === 'assistant' && message.content" class="assistant-wrap">
              <div class="bubble md" v-html="assistantHtml(message)" />
              <button
                type="button"
                class="msg-copy"
                @click="copyText(message.content, message.key)"
              >
                {{ copiedKey === message.key ? 'Скопировано' : 'Копировать' }}
              </button>
            </div>
            <div v-else-if="message.role === 'user'" class="bubble user-bubble">{{ message.content }}</div>
            <button
              v-else-if="message.role === 'tool'"
              type="button"
              class="tool"
              :class="[message.status, { open: isToolOpen(message) }]"
              @click="toggleTool(message.key)"
            >
              <span class="tool-head">
                <strong>{{ message.name }}</strong>
                <span class="tool-status">{{ message.status }}</span>
                <span class="tool-caret">{{ isToolOpen(message) ? '▾' : '▸' }}</span>
              </span>
              <pre v-if="isToolOpen(message)">{{ message.content }}</pre>
            </button>
          </div>

          <div v-if="active.streaming" class="status-line">
            <span class="thinking">Думаю<span class="dots" /></span>
            <span v-if="active.tokenRate">{{ active.tokenRate }} ток/с</span>
          </div>
          <div v-else-if="active.tokenMeta" class="status-line meta">
            <span>{{ active.tokenMeta }}</span>
          </div>
          <button
            v-if="active.showCheckpoint && !active.streaming"
            type="button"
            class="checkpoint"
            @click="emit('restoreCheckpoint')"
          >
            Контрольная точка
          </button>
          <div v-if="active.canRetry && !active.streaming" class="retry-row">
            <button type="button" class="retry" @click="retryTurn">Повторить</button>
            <span v-if="active.retryReason" class="retry-reason">{{ active.retryReason }}</span>
          </div>

          <div v-if="active.pendingApprove" class="approve">
            <p>
              Подтвердить правку
              <strong>{{ active.pendingApprove.name }}</strong>?
            </p>
            <p class="approve-summary">{{ approveSummary }}</p>
            <p class="approve-hint">Откат turn — кнопка «Контрольная точка»</p>
            <button type="button" class="text" @click="approveExpanded = !approveExpanded">
              {{ approveExpanded ? 'Скрыть детали' : 'Показать детали' }}
            </button>
            <pre v-if="approveExpanded">{{ JSON.stringify(active.pendingApprove.args, null, 2) }}</pre>
            <div class="row">
              <button type="button" @click="resolveApprove(false)">Отклонить</button>
              <button type="button" class="text" @click="resolveApprove(true, true)">
                Всегда в этой сессии
              </button>
              <button type="button" class="primary" @click="resolveApprove(true)">Принять</button>
            </div>
          </div>
        </div>

        <section v-if="active.planMarkdown" class="plan-card" :class="active.planRunState">
          <div class="kicker">
            <span>План</span>
            <span v-if="active.planRunState === 'done'" class="plan-badge">Выполнено</span>
            <span v-else-if="active.planRunState === 'running'" class="plan-badge running">
              Выполняется…
            </span>
          </div>
          <h4>{{ active.planTitle || 'План' }}</h4>
          <p>{{ planOverview }}</p>
          <div class="plan-actions">
            <button type="button" class="text" @click="openPlan">Открыть</button>
            <button
              v-if="active.planRunState === 'done'"
              type="button"
              class="plan-repeat"
              :disabled="active.streaming"
              @click="runPlan"
            >
              Повторить
            </button>
            <button
              v-else
              type="button"
              class="go"
              :disabled="active.streaming || active.planRunState === 'running'"
              @click="runPlan"
            >
              {{ active.planRunState === 'running' ? 'Выполняется…' : 'Сделать по плану' }}
            </button>
          </div>
        </section>

        <div class="composer drop-up">
          <p v-if="attachError" class="attach-error">{{ attachError }}</p>
          <div v-if="active.attachments.length" class="context-chips">
            <button
              v-for="(file, index) in active.attachments"
              :key="file.name + index"
              type="button"
              class="chip on attach-chip"
              :title="file.name"
              @click="removeAttach(index)"
            >
              {{ file.name }} ×
            </button>
          </div>
          <textarea
            ref="composerEl"
            v-model="active.draft"
            rows="1"
            :placeholder="composerPlaceholder"
            :disabled="active.streaming"
            @input="resizeComposer"
            @keydown="onComposerKey"
          />
          <div class="composer-bar">
            <div class="menus">
              <div ref="modeWrap" class="menu-wrap">
                <button
                  type="button"
                  class="pill"
                  :class="active.mode"
                  :aria-expanded="modeMenu"
                  @click="modeMenu = !modeMenu; modelMenu = false; historyOpen = false"
                >
                  <span v-if="active.mode === 'agent'" class="glyph" aria-hidden="true">
                    <svg viewBox="0 0 24 24" width="12" height="12">
                      <circle cx="12" cy="12" r="3.5" fill="currentColor" />
                      <path d="M12 3v3M12 18v3M3 12h3M18 12h3" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" fill="none" />
                    </svg>
                  </span>
                  <span v-else-if="active.mode === 'plan'" class="glyph" aria-hidden="true">
                    <svg viewBox="0 0 24 24" width="12" height="12">
                      <path d="M5 7h14M5 12h14M5 17h10" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" />
                    </svg>
                  </span>
                  <span v-else class="glyph" aria-hidden="true">
                    <svg viewBox="0 0 24 24" width="12" height="12">
                      <circle cx="12" cy="12" r="7" fill="none" stroke="currentColor" stroke-width="1.8" />
                    </svg>
                  </span>
                  <span class="pill-label">{{ modeLabel(active.mode) }}</span>
                  <span class="caret">▾</span>
                </button>
                <div v-if="modeMenu" class="menu">
                  <button type="button" @click="setMode('agent')">Агент</button>
                  <button type="button" @click="setMode('plan')">План</button>
                  <button type="button" @click="setMode('chat')">Чат</button>
                </div>
              </div>
              <div ref="modelWrap" class="menu-wrap model-wrap">
                <button
                  type="button"
                  class="model"
                  :title="modelLabel(active)"
                  :aria-expanded="modelMenu"
                  @click="toggleModelMenu"
                >
                  <span class="model-label">{{ modelLabel(active) }}</span>
                  <span class="caret">▾</span>
                </button>
                <div v-if="modelMenu" class="menu model-menu">
                  <p v-if="modelMenuLoading" class="menu-hint">Загрузка…</p>
                  <p v-else-if="modelMenuError" class="menu-hint">{{ modelMenuError }}</p>
                  <button
                    v-for="option in modelOptions"
                    :key="`${option.providerId}:${option.model}`"
                    type="button"
                    :class="{ active: active.providerId === option.providerId && active.model === option.model }"
                    @click="setModel(option)"
                  >
                    <span v-if="providers.length > 1" class="menu-provider">{{ option.provider }} · </span>
                    {{ option.model }}
                  </button>
                </div>
              </div>
            </div>
            <div class="composer-actions">
              <ContextUsageIndicator :usage="contextUsage" />
              <div class="attach-wrap">
                <button
                  type="button"
                  class="icon-btn"
                  aria-label="Вложение"
                  aria-describedby="attach-help-popover"
                  @click="pickAttach"
                >
                  <svg viewBox="0 0 24 24" width="16" height="16" aria-hidden="true">
                    <path
                      fill="none"
                      stroke="currentColor"
                      stroke-width="1.8"
                      d="M8 12.5 14.2 6.3a3 3 0 0 1 4.2 4.2l-7.4 7.4a4.2 4.2 0 0 1-6-6L12 5"
                    />
                  </svg>
                </button>
                <div id="attach-help-popover" class="attach-popover" role="tooltip">
                  {{ attachHelpText }}
                </div>
                <input
                  ref="fileInput"
                  type="file"
                  accept=".xlsx,.xls,.docx,.pptx,.pdf,.txt,.md,.csv,.json"
                  hidden
                  @change="onAttach"
                />
              </div>
              <button
                type="button"
                class="send"
                :class="active.mode"
                :disabled="!active.streaming && !active.draft.trim()"
                @click="active.streaming ? stop() : send()"
              >
                <span v-if="active.streaming" class="stop-square" />
                <span v-else>↑</span>
              </button>
            </div>
          </div>
        </div>
      </div>
    </template>

    <div v-if="planOpen" class="modal-back" @click.self="planOpen = false">
      <div ref="planModal" class="modal" role="dialog" aria-label="План">
        <header>План</header>
        <textarea v-model="planDraft" class="plan-edit" />
        <footer>
          <button type="button" @click="planOpen = false">Закрыть</button>
          <button type="button" class="primary" @click="savePlan">Сохранить</button>
        </footer>
      </div>
    </div>
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
.tabs {
  display: flex;
  align-items: center;
  gap: 4px;
  padding: 6px 6px 6px 8px;
  border-bottom: 1px solid var(--border);
  min-height: 40px;
}
.tab-scroll {
  display: flex;
  gap: 4px;
  overflow: auto;
  flex: 1;
  min-width: 0;
}
.tab-actions {
  display: flex;
  align-items: center;
  gap: 2px;
  flex-shrink: 0;
}
.tab {
  display: inline-flex;
  align-items: center;
  gap: 6px;
  max-width: 170px;
  border: 0;
  background: transparent;
  color: var(--fg-2);
  border-radius: 8px;
  padding: 5px 7px;
  cursor: pointer;
}
.tab.on {
  background: var(--bg-3);
  color: var(--fg-0);
}
.tab-title {
  overflow: hidden;
  text-overflow: ellipsis;
  white-space: nowrap;
  font-size: 0.84rem;
}
.tab-rename {
  width: 110px;
  border: 1px solid var(--accent);
  border-radius: 6px;
  background: var(--bg-1);
  color: var(--fg-0);
  padding: 2px 6px;
  font: inherit;
  font-size: 0.84rem;
}
.x {
  border: 0;
  background: transparent;
  color: var(--fg-3);
  cursor: pointer;
  padding: 0 2px;
  line-height: 1;
  border-radius: 4px;
}
.x:hover {
  color: var(--fg-0);
  background: var(--bg-hover);
}
.plus,
.icon {
  border: 0;
  background: transparent;
  color: var(--fg-1);
  width: 28px;
  height: 28px;
  border-radius: var(--radius-sm);
  cursor: pointer;
  display: grid;
  place-items: center;
  padding: 0;
}
.icon:focus-visible,
.plus:focus-visible {
  outline: none;
  box-shadow: var(--focus-ring);
}
.plus:hover,
.icon:hover:not(:disabled),
.tab:hover {
  background: var(--bg-hover);
}
.icon:disabled {
  opacity: 0.4;
  cursor: default;
}
.icon.on {
  background: var(--bg-3);
  color: var(--accent);
}
.spin {
  width: 10px;
  height: 10px;
  border: 2px solid var(--fg-3);
  border-top-color: var(--accent);
  border-radius: 50%;
  animation: spin 0.8s linear infinite;
  flex-shrink: 0;
}
@keyframes spin {
  to { transform: rotate(360deg); }
}
.history-panel {
  border-bottom: 1px solid var(--border);
  background: var(--bg-2);
  max-height: 220px;
  display: flex;
  flex-direction: column;
  min-height: 0;
}
.history-head {
  display: flex;
  justify-content: space-between;
  align-items: center;
  padding: 8px 12px 4px;
  font-size: 0.78rem;
  color: var(--fg-3);
}
.history-hint {
  margin: 0;
  padding: 8px 12px 12px;
  color: var(--fg-3);
  font-size: 0.82rem;
}
.history-hint.error {
  color: var(--danger);
}
.history-list {
  list-style: none;
  margin: 0;
  padding: 0 6px 8px;
  overflow: auto;
}
.history-list li {
  display: flex;
  align-items: stretch;
  gap: 2px;
}
.history-main {
  flex: 1;
  min-width: 0;
  text-align: left;
  border: 0;
  background: transparent;
  color: var(--fg-1);
  border-radius: 8px;
  padding: 8px;
  cursor: pointer;
  display: grid;
  gap: 2px;
}
.history-main:hover {
  background: var(--bg-hover);
}
.history-title {
  overflow: hidden;
  text-overflow: ellipsis;
  white-space: nowrap;
  font-size: 0.86rem;
  color: var(--fg-0);
}
.history-meta {
  font-size: 0.72rem;
  color: var(--fg-3);
}
.history-del {
  border: 0;
  background: transparent;
  color: var(--fg-3);
  width: 28px;
  border-radius: 8px;
  cursor: pointer;
}
.history-del:hover {
  color: var(--danger);
  background: var(--bg-hover);
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
.cta-title {
  margin: 0;
  color: var(--fg-0);
  font-size: 1rem;
  font-weight: 650;
}
.cta-text {
  margin: 0 0 4px;
  max-width: 260px;
  font-size: 0.88rem;
  color: var(--fg-3);
  line-height: 1.45;
}
.empty-state {
  flex: 1;
  align-self: stretch;
  width: 100%;
  display: flex;
  flex-direction: column;
  justify-content: center;
  gap: 8px;
  padding: 20px 16px;
  order: 1;
}
.empty-state h3 {
  margin: 0;
  width: 100%;
  font-family: var(--font-display);
  font-size: 1.05rem;
  letter-spacing: -0.02em;
  color: var(--fg-0);
  font-weight: 650;
}
.empty-state p {
  margin: 0 0 8px;
  width: 100%;
  color: var(--fg-3);
  font-size: 0.88rem;
}
.empty-state .attach-help {
  margin: 0 0 10px;
  width: 100%;
  max-width: none;
  font-size: 0.8rem;
  line-height: 1.4;
  color: var(--fg-3);
}
.hints {
  display: grid;
  gap: 6px;
}
.hint {
  text-align: left;
  border: 1px solid var(--border);
  background: var(--bg-2);
  color: var(--fg-1);
  border-radius: var(--radius-md);
  padding: 10px 12px;
  cursor: pointer;
  font-size: 0.86rem;
  line-height: 1.35;
}
.hint:focus-visible {
  outline: none;
  box-shadow: var(--focus-ring);
}
.hint:hover:not(:disabled) {
  border-color: var(--accent);
  color: var(--fg-0);
}
.hint:disabled {
  opacity: 0.5;
  cursor: default;
}
.transcript {
  flex: 1;
  overflow: auto;
  padding: 12px 12px 8px;
  display: flex;
  flex-direction: column;
  gap: 10px;
  order: 1;
}
.stack.empty .transcript {
  display: none;
}
.stack .plan-card { order: 2; }
.stack .composer { order: 3; }
.stack.empty .empty-state { order: 1; }
.stack.empty .plan-card { order: 2; }
.stack.empty .composer { order: 3; }
.msg.user {
  display: flex;
  justify-content: flex-end;
}
.user-bubble {
  max-width: 92%;
  background: var(--bg-3);
  color: var(--fg-1);
  border-radius: 12px 12px 4px 12px;
  padding: 8px 11px;
  font-size: 0.92rem;
  line-height: 1.45;
  overflow-wrap: anywhere;
}
.assistant-wrap {
  display: grid;
  gap: 4px;
}
.bubble {
  padding: 2px 0;
  line-height: 1.5;
  overflow-wrap: anywhere;
  color: var(--fg-0);
  font-size: 0.93rem;
}
.msg-copy {
  align-self: start;
  border: 0;
  background: transparent;
  color: var(--fg-3);
  font-size: 0.72rem;
  cursor: pointer;
  padding: 2px 0;
}
.msg-copy:hover {
  color: var(--fg-1);
}
.tool {
  width: 100%;
  text-align: left;
  border: 1px solid var(--border);
  border-radius: 10px;
  padding: 7px 9px;
  font-size: 0.78rem;
  background: var(--bg-2);
  color: var(--fg-2);
  cursor: pointer;
}
.tool:hover {
  border-color: var(--fg-3);
}
.tool-head {
  display: flex;
  align-items: center;
  gap: 8px;
}
.tool-head strong {
  color: var(--fg-0);
  font-weight: 600;
}
.tool-status {
  text-transform: uppercase;
  font-size: 0.68rem;
  letter-spacing: 0.03em;
}
.tool.ok .tool-status { color: var(--success); }
.tool.error .tool-status { color: var(--danger); }
.tool.pending .tool-status { color: var(--accent); }
.tool-caret {
  margin-left: auto;
  color: var(--fg-3);
}
.tool pre {
  white-space: pre-wrap;
  margin: 8px 0 0;
  color: var(--fg-3);
  font-size: 0.75rem;
  max-height: 180px;
  overflow: auto;
}
.status-line {
  display: flex;
  justify-content: space-between;
  gap: 12px;
  color: var(--fg-3);
  font-size: 0.78rem;
  padding: 2px 2px 0;
}
.status-line.meta {
  justify-content: flex-start;
}
.thinking .dots::after {
  content: '';
  animation: dots 1.2s steps(4, end) infinite;
}
@keyframes dots {
  0% { content: ''; }
  25% { content: '.'; }
  50% { content: '..'; }
  75% { content: '...'; }
}
.checkpoint {
  align-self: stretch;
  border: 0;
  background: transparent;
  color: var(--fg-3);
  font-size: 0.75rem;
  cursor: pointer;
  display: flex;
  align-items: center;
  gap: 10px;
}
.checkpoint::before,
.checkpoint::after {
  content: '';
  flex: 1;
  height: 1px;
  background: var(--border);
}
.retry-row {
  display: flex;
  align-items: center;
  gap: 8px;
  margin-top: 4px;
}
.retry {
  border: 1px solid var(--border);
  background: var(--bg-2);
  border-radius: 8px;
  padding: 4px 10px;
  font-size: 0.8rem;
  cursor: pointer;
  color: var(--fg-2);
}
.retry-reason {
  font-size: 0.75rem;
  color: var(--fg-3);
}
.approve,
.plan-card {
  border: 1px solid var(--border);
  border-radius: 12px;
  padding: 12px;
  background: var(--bg-2);
}
.approve {
  margin-top: 4px;
}
.approve p {
  margin: 0 0 6px;
  font-size: 0.88rem;
}
.approve-summary {
  color: var(--fg-3);
  font-size: 0.8rem !important;
}
.approve-hint {
  color: var(--fg-3);
  font-size: 0.75rem !important;
  opacity: 0.85;
}
.approve .row {
  flex-wrap: wrap;
  gap: 6px;
}
.approve pre {
  margin: 8px 0;
  max-height: 160px;
  overflow: auto;
  font-size: 0.75rem;
  color: var(--fg-3);
  white-space: pre-wrap;
}
.plan-card {
  margin: 0 10px 4px;
  padding: 10px 12px;
}
.plan-card h4 {
  margin: 2px 0 4px;
  font-size: 0.9rem;
  white-space: nowrap;
  overflow: hidden;
  text-overflow: ellipsis;
}
.plan-card p {
  margin: 0 0 8px;
  color: var(--fg-2);
  font-size: 0.82rem;
  line-height: 1.35;
  display: -webkit-box;
  -webkit-box-orient: vertical;
  -webkit-line-clamp: 2;
  overflow: hidden;
}
.kicker {
  display: flex;
  align-items: center;
  justify-content: space-between;
  gap: 8px;
  color: var(--fg-3);
  font-size: 0.7rem;
  letter-spacing: 0.04em;
  text-transform: uppercase;
}
.plan-badge {
  letter-spacing: 0.02em;
  text-transform: none;
  font-weight: 650;
  color: var(--success, #16a34a);
}
.plan-badge.running {
  color: var(--fg-2);
  font-weight: 500;
}
.plan-actions {
  display: flex;
  justify-content: space-between;
  align-items: center;
  gap: 8px;
}
button.plan-repeat {
  background: transparent;
  border: 1px solid var(--border);
  color: var(--fg-1);
  font-weight: 550;
  border-radius: 8px;
  padding: 7px 11px;
  cursor: pointer;
  font-size: 0.84rem;
}
button.plan-repeat:disabled {
  opacity: 0.5;
  cursor: default;
}
button.text {
  background: transparent;
  border: 0;
  color: var(--fg-1);
  cursor: pointer;
  padding: 4px 0;
  font-size: 0.84rem;
}
button.go {
  background: var(--success);
  border: 0;
  color: #052e16;
  font-weight: 650;
  border-radius: 8px;
  padding: 7px 11px;
  cursor: pointer;
  font-size: 0.84rem;
}
button.go:disabled {
  opacity: 0.5;
  cursor: default;
}
.composer {
  position: relative;
  margin: 8px;
  border: 1px solid var(--border);
  border-radius: var(--radius-lg);
  background: var(--bg-2);
  padding: 8px;
}
.glyph {
  display: inline-flex;
  align-items: center;
  justify-content: center;
  margin-right: 2px;
}
.context-chips {
  display: flex;
  flex-wrap: wrap;
  gap: 6px;
  padding: 2px 4px 6px;
  min-height: 0;
}
.chip {
  max-width: 100%;
  border: 1px solid var(--border);
  background: transparent;
  color: var(--fg-3);
  border-radius: var(--radius-sm);
  padding: 3px 9px;
  font-size: 0.74rem;
  cursor: pointer;
  overflow: hidden;
  text-overflow: ellipsis;
  white-space: nowrap;
}
.chip.on {
  color: var(--fg-0);
  border-color: color-mix(in srgb, var(--accent) 55%, var(--border));
  background: color-mix(in srgb, var(--accent) 12%, transparent);
}
.chip:disabled {
  opacity: 0.45;
  cursor: default;
}
.attach-chip {
  max-width: 140px;
}
.composer textarea {
  width: 100%;
  resize: none;
  border: 0;
  outline: none;
  box-shadow: none;
  background: transparent;
  color: var(--fg-0);
  padding: 4px 6px;
  min-height: 56px;
  max-height: min(40vh, 280px);
  overflow-y: auto;
  border-radius: var(--radius-sm);
}
.composer textarea:focus-visible {
  outline: none;
  box-shadow: var(--focus-ring);
}
.composer :is(button, .pill, .model, .send, .icon-btn, .chip):focus {
  outline: none;
}
.composer :is(button, .pill, .model, .send, .icon-btn, .chip):focus-visible {
  outline: none;
  box-shadow: var(--focus-ring);
}
.composer-bar {
  display: flex;
  justify-content: space-between;
  align-items: center;
  gap: 6px;
  min-width: 0;
}
.menus {
  display: flex;
  flex: 1 1 auto;
  flex-wrap: nowrap;
  align-items: center;
  gap: 6px;
  min-width: 0;
}
.menu-wrap {
  position: relative;
  display: flex;
  flex: 0 1 auto;
  max-width: 42%;
  min-width: 0;
}
.model-wrap {
  flex: 1 1 auto;
  max-width: none;
}
.pill,
.model {
  display: inline-flex;
  align-items: center;
  gap: 4px;
  width: 100%;
  max-width: 100%;
  border: 1px solid var(--border);
  background: transparent;
  color: var(--fg-2);
  border-radius: var(--radius-sm);
  padding: 3px 8px;
  cursor: pointer;
  font-size: 0.78rem;
  white-space: nowrap;
  min-width: 0;
}
.pill-label,
.model-label {
  overflow: hidden;
  text-overflow: ellipsis;
  white-space: nowrap;
  min-width: 0;
}
.pill.agent {
  color: #93c5fd;
  border-color: color-mix(in srgb, #3b82f6 55%, var(--border));
}
.pill.plan {
  color: #facc15;
  border-color: color-mix(in srgb, #eab308 55%, var(--border));
}
.caret {
  flex-shrink: 0;
  color: var(--fg-3);
}
.menu {
  position: absolute;
  left: 0;
  bottom: calc(100% + 6px);
  z-index: 20;
  min-width: 160px;
  max-width: min(280px, 70vw);
  background: var(--bg-1);
  border: 1px solid var(--border);
  border-radius: var(--radius-md);
  box-shadow: var(--shadow);
  padding: 4px;
  display: grid;
}
.menu button {
  text-align: left;
  background: transparent;
  border: 0;
  color: var(--fg-0);
  padding: 8px 10px;
  border-radius: 8px;
  cursor: pointer;
}
.menu button:hover { background: var(--bg-hover); }
.menu button.active { background: var(--bg-hover); color: var(--accent); }
.menu-hint {
  margin: 0;
  padding: 8px 10px;
  color: var(--fg-3);
  font-size: 0.82rem;
}
.menu-provider {
  color: var(--fg-3);
}
.model-menu {
  left: auto;
  right: 0;
  min-width: min(220px, 70vw);
  max-height: 280px;
  overflow-y: auto;
}
.composer-actions {
  display: flex;
  flex-shrink: 0;
  align-items: center;
  gap: 4px;
}
.attach-wrap {
  position: relative;
  display: flex;
  align-items: center;
}
.attach-popover {
  position: absolute;
  right: 0;
  bottom: calc(100% + 8px);
  z-index: 6;
  width: 240px;
  padding: 10px 12px;
  background: var(--bg-1);
  border: 1px solid var(--border);
  border-radius: var(--radius-md);
  box-shadow: var(--shadow);
  color: var(--fg-2);
  font-size: 0.78rem;
  line-height: 1.4;
  pointer-events: none;
  opacity: 0;
  visibility: hidden;
  transition: opacity 0.12s ease, visibility 0.12s ease;
}
.attach-wrap:hover .attach-popover,
.attach-wrap:focus-within .attach-popover {
  opacity: 1;
  visibility: visible;
}
.icon-btn {
  box-sizing: border-box;
  display: grid;
  place-items: center;
  width: 30px;
  height: 30px;
  padding: 0;
  border: 0;
  background: transparent;
  color: var(--fg-2);
  cursor: pointer;
  border-radius: 8px;
  line-height: 0;
}
.icon-btn svg {
  display: block;
}
.icon-btn:hover {
  color: var(--fg-0);
  background: var(--bg-hover);
}
.send {
  width: 30px;
  height: 30px;
  border-radius: 50%;
  border: 0;
  cursor: pointer;
  display: grid;
  place-items: center;
  font-size: 0.95rem;
}
.send:disabled {
  background: var(--bg-3);
  color: var(--fg-3);
  cursor: default;
}
.send.agent:not(:disabled),
.send.chat:not(:disabled) {
  background: var(--accent);
  color: white;
}
.send.plan:not(:disabled) {
  background: #eab308;
  color: #111;
}
.stop-square {
  width: 10px;
  height: 10px;
  background: currentColor;
  border-radius: 2px;
}
.attach-error {
  color: var(--danger);
  font-size: 0.78rem;
  margin: 0 6px 4px;
}
.row {
  display: flex;
  justify-content: flex-end;
  gap: 8px;
  margin-top: 8px;
}
.modal-back {
  position: fixed;
  inset: 0;
  background: var(--overlay);
  z-index: var(--z-plan-modal);
  display: grid;
  place-items: center;
  padding: 16px;
}
.modal {
  width: min(720px, 100%);
  height: min(80vh, 640px);
  background: var(--bg-1);
  border: 1px solid var(--border);
  border-radius: var(--radius-lg);
  box-shadow: var(--shadow);
  display: flex;
  flex-direction: column;
  padding: 12px;
  gap: 8px;
}
.modal header { font-weight: 650; }
.plan-edit {
  flex: 1;
  resize: none;
  border-radius: 10px;
  border: 1px solid var(--border);
  background: var(--bg-2);
  color: var(--fg-0);
  padding: 10px;
  font: inherit;
}
.modal footer {
  display: flex;
  justify-content: flex-end;
  gap: 8px;
}
:deep(.md p) { margin: 0.35em 0; }
:deep(.md ul),
:deep(.md ol) {
  margin: 0.4em 0;
  padding-left: 1.2em;
}
:deep(.md li) { margin: 0.15em 0; }
:deep(.md table) {
  border-collapse: collapse;
  width: 100%;
  font-size: 0.84rem;
  margin: 0.5em 0;
}
:deep(.md th),
:deep(.md td) {
  border: 1px solid var(--border);
  padding: 4px 6px;
}
:deep(.md code) {
  background: var(--bg-3);
  padding: 0 4px;
  border-radius: 4px;
  font-size: 0.86em;
}
:deep(.code-block) {
  position: relative;
  margin: 0.5em 0;
  border: 1px solid var(--border);
  border-radius: 10px;
  background: var(--bg-2);
  overflow: hidden;
}
:deep(.code-block .copy-code) {
  position: absolute;
  top: 6px;
  right: 6px;
  z-index: 1;
  border: 1px solid var(--border);
  background: var(--bg-1);
  color: var(--fg-2);
  border-radius: 6px;
  padding: 2px 7px;
  font-size: 0.7rem;
  cursor: pointer;
}
:deep(.code-block pre) {
  margin: 0;
  padding: 28px 10px 10px;
  overflow: auto;
}
:deep(.code-block code) {
  background: transparent;
  padding: 0;
}
</style>
