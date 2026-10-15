<script setup lang="ts">
import { ref } from 'vue';
import type { UiMsg } from '../types';

defineProps<{
  messages: UiMsg[];
  streaming: boolean;
  tokenRate: number | null;
  tokenMeta: string | null;
  showCheckpoint: boolean;
  canRetry: boolean;
  retryReason: string | null;
  pendingApprove: { name: string; args: Record<string, unknown> } | null;
  approveSummary: string;
  approveExpanded: boolean;
  copiedKey: string;
  assistantHtml: (message: Extract<UiMsg, { role: 'assistant' }>) => string;
  isToolOpen: (message: Extract<UiMsg, { role: 'tool' }>) => boolean;
}>();

const emit = defineEmits<{
  'copy-message': [text: string, key: string];
  'toggle-tool': [key: string];
  'transcript-click': [event: MouseEvent];
  'restore-checkpoint': [];
  retry: [];
  'update:approveExpanded': [value: boolean];
  'resolve-approve': [ok: boolean, always?: boolean];
}>();

const el = ref<HTMLElement | null>(null);

defineExpose({ el });
</script>

<template>
  <div ref="el" class="transcript" @click="emit('transcript-click', $event)">
    <div
      v-for="message in messages"
      :key="message.key"
      class="msg"
      :class="message.role"
    >
      <div
        v-if="message.role === 'assistant' && message.content"
        class="assistant-wrap"
      >
        <div class="bubble md" v-html="assistantHtml(message)" />
        <button
          type="button"
          class="msg-copy"
          @click="emit('copy-message', message.content, message.key)"
        >
          {{ copiedKey === message.key ? 'Скопировано' : 'Копировать' }}
        </button>
      </div>
      <div v-else-if="message.role === 'user'" class="bubble user-bubble">
        {{ message.content }}
      </div>
      <button
        v-else-if="message.role === 'tool'"
        type="button"
        class="tool"
        :class="[message.status, { open: isToolOpen(message) }]"
        @click="emit('toggle-tool', message.key)"
      >
        <span class="tool-head">
          <strong>{{ message.name }}</strong>
          <span class="tool-status">{{ message.status }}</span>
          <span class="tool-caret">{{ isToolOpen(message) ? '▾' : '▸' }}</span>
        </span>
        <pre v-if="isToolOpen(message)">{{ message.content }}</pre>
      </button>
    </div>

    <div v-if="streaming" class="status-line">
      <span class="thinking">Думаю<span class="dots" /></span>
      <span v-if="tokenRate">{{ tokenRate }} ток/с</span>
    </div>
    <div v-else-if="tokenMeta" class="status-line meta">
      <span>{{ tokenMeta }}</span>
    </div>
    <button
      v-if="showCheckpoint && !streaming"
      type="button"
      class="checkpoint"
      @click="emit('restore-checkpoint')"
    >
      Контрольная точка
    </button>
    <div v-if="canRetry && !streaming" class="retry-row">
      <button type="button" class="retry" @click="emit('retry')">
        Повторить
      </button>
      <span v-if="retryReason" class="retry-reason">{{ retryReason }}</span>
    </div>

    <div v-if="pendingApprove" class="approve">
      <p>
        Подтвердить правку
        <strong>{{ pendingApprove.name }}</strong
        >?
      </p>
      <p class="approve-summary">
        {{ approveSummary }}
      </p>
      <p class="approve-hint">Откат turn — кнопка «Контрольная точка»</p>
      <button
        type="button"
        class="text"
        @click="emit('update:approveExpanded', !approveExpanded)"
      >
        {{ approveExpanded ? 'Скрыть детали' : 'Показать детали' }}
      </button>
      <pre v-if="approveExpanded">{{
        JSON.stringify(pendingApprove.args, null, 2)
      }}</pre>
      <div class="row">
        <button type="button" @click="emit('resolve-approve', false)">
          Отклонить
        </button>
        <button
          type="button"
          class="text"
          @click="emit('resolve-approve', true, true)"
        >
          Всегда в этой сессии
        </button>
        <button
          type="button"
          class="primary"
          @click="emit('resolve-approve', true)"
        >
          Принять
        </button>
      </div>
    </div>
  </div>
</template>

<style scoped lang="scss">
.transcript {
  flex: 1;
  overflow: auto;
  padding: 12px 12px 8px;
  display: flex;
  flex-direction: column;
  gap: 10px;
  order: 1;
}
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
.tool.ok .tool-status {
  color: var(--success);
}
.tool.error .tool-status {
  color: var(--danger);
}
.tool.pending .tool-status {
  color: var(--accent);
}
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
  0% {
    content: '';
  }
  25% {
    content: '.';
  }
  50% {
    content: '..';
  }
  75% {
    content: '...';
  }
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
.approve {
  margin-top: 4px;
  border: 1px solid var(--border);
  border-radius: 12px;
  padding: 12px;
  background: var(--bg-2);
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
button.text {
  background: transparent;
  border: 0;
  color: var(--fg-1);
  cursor: pointer;
  padding: 4px 0;
  font-size: 0.84rem;
}
.row {
  display: flex;
  justify-content: flex-end;
  gap: 8px;
  margin-top: 8px;
}
:deep(.md p) {
  margin: 0.35em 0;
}
:deep(.md ul),
:deep(.md ol) {
  margin: 0.4em 0;
  padding-left: 1.2em;
}
:deep(.md li) {
  margin: 0.15em 0;
}
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
