<script setup lang="ts">
import type { PlanRunState } from '../types';

defineProps<{
  title: string | null;
  overview: string;
  runState: PlanRunState;
  streaming: boolean;
}>();

const emit = defineEmits<{
  open: [];
  run: [];
}>();
</script>

<template>
  <section class="plan-card" :class="runState">
    <div class="kicker">
      <span>План</span>
      <span v-if="runState === 'done'" class="plan-badge">Выполнено</span>
      <span v-else-if="runState === 'running'" class="plan-badge running">
        Выполняется…
      </span>
    </div>
    <h4>{{ title || 'План' }}</h4>
    <p>{{ overview }}</p>
    <div class="plan-actions">
      <button type="button" class="text" @click="emit('open')">Открыть</button>
      <button
        v-if="runState === 'done'"
        type="button"
        class="plan-repeat"
        :disabled="streaming"
        @click="emit('run')"
      >
        Повторить
      </button>
      <button
        v-else
        type="button"
        class="go"
        :disabled="streaming || runState === 'running'"
        @click="emit('run')"
      >
        {{ runState === 'running' ? 'Выполняется…' : 'Сделать по плану' }}
      </button>
    </div>
  </section>
</template>

<style scoped lang="scss">
.plan-card {
  margin: 0 10px 4px;
  border: 1px solid var(--border);
  border-radius: 12px;
  padding: 10px 12px;
  background: var(--bg-2);
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
</style>
