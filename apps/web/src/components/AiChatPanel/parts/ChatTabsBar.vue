<script setup lang="ts">
import { ref } from 'vue';

export type TabBarTab = { id: string; title: string; streaming: boolean };

defineProps<{
  tabs: TabBarTab[];
  activeId: string;
  renamingId: string;
  renameDraft: string;
  historyOpen: boolean;
  hasSessionLog: boolean;
  logCopied: boolean;
}>();

const emit = defineEmits<{
  'update:activeId': [id: string];
  'update:renameDraft': [value: string];
  add: [];
  close: [id: string];
  'start-rename': [tab: TabBarTab];
  'commit-rename': [tab: TabBarTab];
  'rename-key': [event: KeyboardEvent, tab: TabBarTab];
  'toggle-history': [];
  'copy-log': [];
  collapse: [];
}>();

const historyBtn = ref<HTMLElement | null>(null);

defineExpose({ historyBtn });
</script>

<template>
  <header class="tabs">
    <div class="tab-scroll">
      <div
        v-for="tab in tabs"
        :key="tab.id"
        class="tab"
        :class="{ on: tab.id === activeId }"
        @click="emit('update:activeId', tab.id)"
      >
        <span v-if="tab.streaming" class="spin" />
        <input
          v-if="renamingId === tab.id"
          :value="renameDraft"
          class="tab-rename"
          :data-id="tab.id"
          @input="
            emit(
              'update:renameDraft',
              ($event.target as HTMLInputElement).value,
            )
          "
          @click.stop
          @keydown="emit('rename-key', $event, tab)"
          @blur="emit('commit-rename', tab)"
        />
        <span
          v-else
          class="tab-title"
          :title="tab.title"
          @dblclick.stop="emit('start-rename', tab)"
          >{{ tab.title }}</span
        >
        <button
          type="button"
          class="x"
          title="Закрыть"
          @click.stop="emit('close', tab.id)"
        >
          ×
        </button>
      </div>
    </div>
    <div class="tab-actions">
      <button type="button" class="plus" title="Новый чат" @click="emit('add')">
        +
      </button>
      <button
        ref="historyBtn"
        type="button"
        class="icon"
        title="История"
        :class="{ on: historyOpen }"
        @click="emit('toggle-history')"
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
        @click="emit('copy-log')"
      >
        <svg
          v-if="!logCopied"
          viewBox="0 0 24 24"
          width="16"
          height="16"
          aria-hidden="true"
        >
          <rect
            x="9"
            y="9"
            width="11"
            height="11"
            rx="2"
            fill="none"
            stroke="currentColor"
            stroke-width="1.7"
          />
          <path
            d="M6 15H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h8a2 2 0 0 1 2 2v1"
            fill="none"
            stroke="currentColor"
            stroke-width="1.7"
          />
        </svg>
        <svg
          v-else
          viewBox="0 0 24 24"
          width="16"
          height="16"
          aria-hidden="true"
        >
          <path
            d="M5 12.5l4.5 4.5L19 7"
            fill="none"
            stroke="currentColor"
            stroke-width="2"
            stroke-linecap="round"
            stroke-linejoin="round"
          />
        </svg>
      </button>
      <button
        type="button"
        class="icon"
        title="Свернуть"
        @click="emit('collapse')"
      >
        <svg viewBox="0 0 24 24" width="16" height="16" aria-hidden="true">
          <path
            d="M9 6l6 6-6 6"
            fill="none"
            stroke="currentColor"
            stroke-width="1.8"
            stroke-linecap="round"
            stroke-linejoin="round"
          />
        </svg>
      </button>
    </div>
  </header>
</template>

<style scoped lang="scss">
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
  to {
    transform: rotate(360deg);
  }
}
</style>
