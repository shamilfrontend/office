<script setup lang="ts">
defineProps<{
  tabs: Array<{ id: string; name: string; type: 'SHEET' | 'DOC' }>;
  activeId: string;
  creating: boolean;
  newMenuOpen: boolean;
}>();

const emit = defineEmits<{
  switch: [id: string];
  close: [id: string, event?: Event];
  'toggle-new': [];
  'close-new': [];
  create: [type: 'SHEET' | 'DOC'];
  'open-files': [];
}>();
</script>

<template>
  <div class="file-tabs-row" @keydown.escape="emit('close-new')">
    <div class="file-tabs" role="tablist" aria-label="Открытые файлы">
      <button
        v-for="tab in tabs"
        :key="tab.id"
        type="button"
        role="tab"
        class="file-tab"
        :class="{ active: tab.id === activeId }"
        :aria-selected="tab.id === activeId"
        :title="tab.name"
        @click="emit('switch', tab.id)"
        @auxclick.middle.prevent="emit('close', tab.id, $event)"
      >
        <span
          class="file-tab-type"
          :class="tab.type === 'SHEET' ? 'sheet' : 'doc'"
          aria-hidden="true"
        >
          <svg
            v-if="tab.type === 'SHEET'"
            viewBox="0 0 24 24"
            width="14"
            height="14"
          >
            <rect
              x="3"
              y="4"
              width="18"
              height="16"
              rx="2"
              fill="currentColor"
              opacity="0.2"
            />
            <path
              d="M3 9h18M9 4v16"
              stroke="currentColor"
              stroke-width="1.5"
              fill="none"
            />
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
          @click="emit('close', tab.id, $event)"
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
          @click="emit('toggle-new')"
        >
          +
        </button>
        <div v-if="newMenuOpen" class="ui-dropdown file-tabs-menu" role="menu">
          <button
            type="button"
            role="menuitem"
            @click="emit('create', 'SHEET')"
          >
            Новая таблица
          </button>
          <button type="button" role="menuitem" @click="emit('create', 'DOC')">
            Новый документ
          </button>
          <button type="button" role="menuitem" @click="emit('open-files')">
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
      @click="emit('close-new')"
    />
  </div>
</template>

<style scoped lang="scss">
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
</style>
