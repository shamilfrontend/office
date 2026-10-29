<script setup lang="ts">
import { ref } from 'vue';

defineProps<{
  busy: boolean;
  uploading: boolean;
}>();

const emit = defineEmits<{
  create: [type: 'SHEET' | 'DOC' | 'SLIDE'];
  'local-file': [event: Event];
}>();

const fileInput = ref<HTMLInputElement | null>(null);

function pickLocal() {
  fileInput.value?.click();
}

defineExpose({ fileInput });
</script>

<template>
  <div class="create-block">
    <h2 class="section-title">Создать</h2>
    <div class="tiles">
      <button
        type="button"
        class="tile"
        :disabled="busy"
        @click="emit('create', 'SHEET')"
      >
        <span class="tile-icon sheet" aria-hidden="true">
          <svg viewBox="0 0 24 24" width="22" height="22">
            <rect
              x="3"
              y="4"
              width="18"
              height="16"
              rx="2"
              fill="currentColor"
              opacity="0.18"
            />
            <path
              d="M3 9h18M9 4v16"
              stroke="currentColor"
              stroke-width="1.5"
              fill="none"
            />
          </svg>
        </span>
        <span class="tile-text">
          <span class="tile-label">Новая таблица</span>
        </span>
      </button>
      <button
        type="button"
        class="tile"
        :disabled="busy"
        @click="emit('create', 'DOC')"
      >
        <span class="tile-icon doc" aria-hidden="true">
          <svg viewBox="0 0 24 24" width="22" height="22">
            <path
              d="M6 3h8l4 4v14a1 1 0 0 1-1 1H6a1 1 0 0 1-1-1V4a1 1 0 0 1 1-1z"
              fill="currentColor"
              opacity="0.18"
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
        <span class="tile-text">
          <span class="tile-label">Новый документ</span>
        </span>
      </button>
      <button
        type="button"
        class="tile"
        :disabled="busy"
        @click="emit('create', 'SLIDE')"
      >
        <span class="tile-icon slide" aria-hidden="true">
          <svg viewBox="0 0 24 24" width="22" height="22">
            <rect
              x="3"
              y="5"
              width="18"
              height="12"
              rx="2"
              fill="currentColor"
              opacity="0.18"
            />
            <path
              d="M8 21h8M12 17v4"
              stroke="currentColor"
              stroke-width="1.5"
              stroke-linecap="round"
              fill="none"
            />
          </svg>
        </span>
        <span class="tile-text">
          <span class="tile-label">Новая презентация</span>
        </span>
      </button>
      <button
        type="button"
        class="tile"
        :disabled="busy || uploading"
        :aria-busy="uploading"
        @click="pickLocal"
      >
        <span class="tile-icon upload" aria-hidden="true">
          <span v-if="uploading" class="tile-spinner" />
          <svg v-else viewBox="0 0 24 24" width="22" height="22">
            <path
              d="M12 16V6M12 6l-4 4M12 6l4 4M5 18v1a1 1 0 0 0 1 1h12a1 1 0 0 0 1-1v-1"
              stroke="currentColor"
              stroke-width="1.6"
              stroke-linecap="round"
              stroke-linejoin="round"
              fill="none"
            />
          </svg>
        </span>
        <span class="tile-text">
          <span class="tile-label">{{
            uploading ? 'Открываем…' : 'Открыть с компьютера'
          }}</span>
        </span>
      </button>
      <input
        ref="fileInput"
        type="file"
        accept=".xlsx,.xls,.docx,.pptx,.csv,.json"
        hidden
        @change="emit('local-file', $event)"
      />
    </div>
  </div>
</template>

<style scoped lang="scss">
.create-block {
  margin-bottom: 32px;
}
.recent-block {
  width: 100%;
}
.section-title {
  margin: 0 0 14px;
  font-family: var(--font-display);
  font-size: 1.05rem;
  letter-spacing: -0.02em;
  color: var(--fg-0);
  font-weight: 600;
}
.tiles {
  display: grid;
  grid-template-columns: repeat(auto-fill, minmax(220px, 1fr));
  gap: 10px;
  max-width: 920px;
}
.tile {
  display: flex;
  align-items: center;
  gap: 12px;
  min-height: 64px;
  padding: 12px 14px;
  border: 1px solid var(--border);
  border-radius: var(--radius-md);
  background: var(--bg-2);
  color: var(--fg-0);
  cursor: pointer;
  text-align: left;
  transition:
    background 0.12s ease,
    border-color 0.12s ease;
}
.tile:hover:not(:disabled) {
  background: var(--bg-hover);
  border-color: var(--fg-3);
}
.tile:focus-visible {
  outline: none;
  box-shadow: var(--focus-ring);
}
.tile:disabled {
  cursor: not-allowed;
  opacity: 0.65;
}
.tile-icon {
  flex-shrink: 0;
  width: 40px;
  height: 40px;
  border-radius: var(--radius-sm);
  display: grid;
  place-items: center;
}
.tile-icon.sheet {
  background: color-mix(in srgb, var(--success) 14%, transparent);
  color: var(--success);
}
.tile-icon.doc {
  background: color-mix(in srgb, var(--accent) 14%, transparent);
  color: var(--accent);
}
.tile-icon.slide {
  background: color-mix(in srgb, #e8a317 16%, transparent);
  color: #e8a317;
}
.tile-icon.upload {
  background: var(--bg-3);
  color: var(--fg-2);
}
.tile-spinner {
  width: 20px;
  height: 20px;
  border: 2px solid color-mix(in srgb, currentColor 25%, transparent);
  border-top-color: currentColor;
  border-radius: 50%;
  animation: tile-spin 0.7s linear infinite;
}
@keyframes tile-spin {
  to {
    transform: rotate(360deg);
  }
}
.tile-text {
  display: grid;
  gap: 2px;
  min-width: 0;
  flex: 1;
}
.tile-label {
  font-size: 0.9rem;
  font-weight: 600;
  color: var(--fg-0);
  line-height: 1.25;
}
</style>
