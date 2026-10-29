<script setup lang="ts">
import AppLogo from '@/components/AppLogo.vue';

defineProps<{
  docName: string;
  titleWidth: string;
  saveFeedback: 'idle' | 'saving' | 'saved';
  docType: 'SHEET' | 'DOC' | 'SLIDE';
}>();

const emit = defineEmits<{
  'update:docName': [value: string];
  'title-change': [];
  home: [];
  save: [];
  download: [format: 'xlsx' | 'docx' | 'pptx'];
  settings: [];
}>();
</script>

<template>
  <header class="top">
    <button type="button" class="brand-btn" @click="emit('home')">
      <AppLogo :size="28" />
    </button>
    <input
      :value="docName"
      class="title"
      :style="{ width: titleWidth }"
      @input="emit('update:docName', ($event.target as HTMLInputElement).value)"
      @change="emit('title-change')"
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
        @click="emit('save')"
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
        @click="emit('download', 'xlsx')"
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
        v-else-if="docType === 'SLIDE'"
        type="button"
        class="top-icon-btn"
        title="Скачать PowerPoint"
        aria-label="Скачать PowerPoint"
        @click="emit('download', 'pptx')"
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
        @click="emit('download', 'docx')"
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
        type="button"
        class="ui-btn ui-btn--ghost settings-btn"
        @click="emit('settings')"
      >
        Настройки
      </button>
    </div>
  </header>
</template>

<style scoped lang="scss">
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
</style>
