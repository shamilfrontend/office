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
    class="ui-overlay overlay"
    @click.self="confirmCancel()"
  >
    <div
      class="ui-modal dialog"
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
        <button type="button" class="ui-btn" @click="confirmCancel()">
          {{ state.cancelLabel }}
        </button>
        <button
          ref="confirmBtn"
          type="button"
          class="ui-btn"
          :class="state.danger ? 'ui-btn--danger' : 'ui-btn--primary'"
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
  z-index: var(--z-confirm);
}
.dialog {
  width: min(400px, 100%);
  padding: 20px 20px 16px;
}
h2 {
  margin: 0;
  font-family: var(--font-display);
  font-size: 1.1rem;
  font-weight: 600;
  letter-spacing: -0.02em;
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
</style>
