<script setup lang="ts">
import type { ChatMode } from '@/univer/officeTools';
import { modeLabel } from '../formatters';

defineProps<{
  mode: ChatMode;
  hints: string[];
  streaming: boolean;
  attachHelpText: string;
}>();

const emit = defineEmits<{
  'apply-hint': [text: string];
}>();
</script>

<template>
  <div class="empty-state">
    <h3>{{ modeLabel(mode) }}</h3>
    <p>Напишите задачу или выберите подсказку.</p>
    <p class="attach-help">
      {{ attachHelpText }}
    </p>
    <div class="hints">
      <button
        v-for="hint in hints"
        :key="hint"
        type="button"
        class="hint"
        :disabled="streaming"
        @click="emit('apply-hint', hint)"
      >
        {{ hint }}
      </button>
    </div>
  </div>
</template>

<style scoped lang="scss">
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
</style>
