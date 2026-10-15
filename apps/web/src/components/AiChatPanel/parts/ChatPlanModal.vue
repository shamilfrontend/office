<script setup lang="ts">
import { ref } from 'vue';

defineProps<{
  open: boolean;
  draft: string;
}>();

const emit = defineEmits<{
  'update:open': [value: boolean];
  'update:draft': [value: string];
  save: [];
}>();

const modalEl = ref<HTMLElement | null>(null);

defineExpose({ modalEl });
</script>

<template>
  <div v-if="open" class="modal-back" @click.self="emit('update:open', false)">
    <div ref="modalEl" class="modal" role="dialog" aria-label="План">
      <header>План</header>
      <textarea
        :value="draft"
        class="plan-edit"
        @input="
          emit('update:draft', ($event.target as HTMLTextAreaElement).value)
        "
      />
      <footer>
        <button type="button" @click="emit('update:open', false)">
          Закрыть
        </button>
        <button type="button" class="primary" @click="emit('save')">
          Сохранить
        </button>
      </footer>
    </div>
  </div>
</template>

<style scoped lang="scss">
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
.modal header {
  font-weight: 650;
}
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
</style>
