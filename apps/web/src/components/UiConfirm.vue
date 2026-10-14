<script setup lang="ts">
import { nextTick, onUnmounted, ref, watch } from 'vue'
import {
  confirmAccept,
  confirmCancel,
  useConfirmState,
} from '@/composables/useConfirm'

const state = useConfirmState()
const confirmBtn = ref<HTMLButtonElement | null>(null)

function onKeydown(e: KeyboardEvent) {
  if (!state.open) return
  if (e.key === 'Escape') {
    e.preventDefault()
    e.stopImmediatePropagation()
    confirmCancel()
  }
}

watch(
  () => state.open,
  async (open) => {
    if (open) {
      window.addEventListener('keydown', onKeydown, true)
      await nextTick()
      confirmBtn.value?.focus()
    } else {
      window.removeEventListener('keydown', onKeydown, true)
    }
  },
)

onUnmounted(() => {
  window.removeEventListener('keydown', onKeydown, true)
})
</script>

<template>
  <div
    v-if="state.open"
    class="overlay"
    @click.self="confirmCancel()"
  >
    <div
      class="dialog"
      role="dialog"
      aria-modal="true"
      aria-labelledby="ui-confirm-title"
      :aria-describedby="state.message ? 'ui-confirm-message' : undefined"
    >
      <h2 id="ui-confirm-title">{{ state.title }}</h2>
      <p v-if="state.message" id="ui-confirm-message" class="message">
        {{ state.message }}
      </p>
      <div class="actions">
        <button type="button" class="cancel" @click="confirmCancel()">
          {{ state.cancelLabel }}
        </button>
        <button
          ref="confirmBtn"
          type="button"
          class="confirm"
          :class="{ danger: state.danger }"
          @click="confirmAccept()"
        >
          {{ state.confirmLabel }}
        </button>
      </div>
    </div>
  </div>
</template>

<style scoped lang="scss">
.overlay {
  position: fixed;
  inset: 0;
  background: rgba(0, 0, 0, 0.55);
  display: grid;
  place-items: center;
  z-index: 60;
  padding: 16px;
}
.dialog {
  width: min(400px, 100%);
  padding: 20px 20px 16px;
  background: var(--bg-2);
  border: 1px solid var(--border);
  border-radius: 16px;
  box-shadow: var(--shadow);
}
h2 {
  margin: 0;
  font-size: 1.1rem;
  color: var(--fg-0);
}
.message {
  margin: 10px 0 0;
  color: var(--fg-2);
  font-size: 0.95rem;
  line-height: 1.45;
}
.actions {
  display: flex;
  justify-content: flex-end;
  gap: 8px;
  margin-top: 20px;
}
button {
  padding: 10px 14px;
  border-radius: 8px;
  border: 1px solid var(--border);
  background: var(--bg-3);
  color: var(--fg-0);
  cursor: pointer;
}
button:hover {
  background: var(--bg-hover);
}
.confirm.danger {
  background: var(--danger);
  border-color: var(--danger);
  color: #fff;
}
.confirm.danger:hover {
  filter: brightness(1.08);
  background: var(--danger);
}
.confirm:not(.danger) {
  background: var(--accent);
  border-color: var(--accent);
  color: #fff;
}
.confirm:not(.danger):hover {
  background: var(--accent-hover);
  border-color: var(--accent-hover);
}
</style>
