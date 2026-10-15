<script setup lang="ts">
import { ref } from 'vue';
import type { ChatSummary } from '@/api/client';
import { asMode, formatHistoryDate, modeLabel } from '../formatters';

defineProps<{
  items: ChatSummary[];
  loading: boolean;
  error: string;
}>();

const emit = defineEmits<{
  close: [];
  reopen: [id: string];
  delete: [id: string];
}>();

const panelEl = ref<HTMLElement | null>(null);

defineExpose({ panelEl });
</script>

<template>
  <div ref="panelEl" class="history-panel">
    <div class="history-head">
      <span>Закрытые чаты</span>
      <button type="button" class="text" @click="emit('close')">Закрыть</button>
    </div>
    <p v-if="loading" class="history-hint">Загрузка…</p>
    <p v-else-if="error" class="history-hint error">{{ error }}</p>
    <p v-else-if="!items.length" class="history-hint">Пока пусто</p>
    <ul v-else class="history-list">
      <li v-for="item in items" :key="item.id">
        <button
          type="button"
          class="history-main"
          @click="emit('reopen', item.id)"
        >
          <span class="history-title">{{ item.title || 'Без названия' }}</span>
          <span class="history-meta"
            >{{ formatHistoryDate(item.updatedAt) }} ·
            {{ modeLabel(asMode(item.mode)) }}</span
          >
        </button>
        <button
          type="button"
          class="history-del"
          title="Удалить"
          @click="emit('delete', item.id)"
        >
          ×
        </button>
      </li>
    </ul>
  </div>
</template>

<style scoped lang="scss">
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
button.text {
  background: transparent;
  border: 0;
  color: var(--fg-1);
  cursor: pointer;
  padding: 4px 0;
  font-size: 0.84rem;
}
</style>
