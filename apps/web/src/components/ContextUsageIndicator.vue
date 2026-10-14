<script setup lang="ts">
import { computed, onBeforeUnmount, ref, watch } from 'vue'
import {
  RESERVE_COLOR,
  type ChatContextUsage,
} from '@/ai/estimateChatContextUsage'
import { formatTokenCount } from '@/ai/estimateTokens'

const props = defineProps<{ usage: ChatContextUsage | null }>()

const open = ref(false)
const root = ref<HTMLElement | null>(null)

const ringSize = 16
const stroke = 2
const r = (ringSize - stroke) / 2
const circumference = 2 * Math.PI * r

const pct = computed(() => props.usage?.percentFull ?? 0)
const dashOffset = computed(
  () =>
    circumference -
    (Math.min(100, Math.max(0, pct.value)) / 100) * circumference,
)
const ringColor = computed(() => {
  const p = pct.value
  if (p >= 95) return '#f87171'
  if (p >= 80) return '#fbbf24'
  return 'currentColor'
})

function toggle() {
  if (!props.usage) return
  open.value = !open.value
}

function onDocMouseDown(e: MouseEvent) {
  const target = e.target as Node
  if (root.value && !root.value.contains(target)) open.value = false
}

function onKeyDown(e: KeyboardEvent) {
  if (e.key === 'Escape') open.value = false
}

watch(open, (isOpen) => {
  if (isOpen) {
    document.addEventListener('mousedown', onDocMouseDown)
    document.addEventListener('keydown', onKeyDown)
  } else {
    document.removeEventListener('mousedown', onDocMouseDown)
    document.removeEventListener('keydown', onKeyDown)
  }
})

onBeforeUnmount(() => {
  document.removeEventListener('mousedown', onDocMouseDown)
  document.removeEventListener('keydown', onKeyDown)
})
</script>

<template>
  <div ref="root" class="ctx-usage">
    <button
      type="button"
      class="icon-btn ctx-btn"
      title="Использование в контексте"
      aria-label="Использование в контексте"
      :aria-expanded="open"
      :disabled="!usage"
      @click.stop="toggle"
    >
      <svg
        :width="ringSize"
        :height="ringSize"
        :viewBox="`0 0 ${ringSize} ${ringSize}`"
        class="ring"
        aria-hidden="true"
      >
        <circle
          :cx="ringSize / 2"
          :cy="ringSize / 2"
          :r="r"
          fill="none"
          stroke="currentColor"
          :stroke-width="stroke"
          opacity="0.25"
        />
        <circle
          :cx="ringSize / 2"
          :cy="ringSize / 2"
          :r="r"
          fill="none"
          :stroke="ringColor"
          :stroke-width="stroke"
          stroke-linecap="round"
          :stroke-dasharray="circumference"
          :stroke-dashoffset="dashOffset"
          :transform="`rotate(-90 ${ringSize / 2} ${ringSize / 2})`"
        />
      </svg>
    </button>

    <div
      v-if="open && usage"
      class="ctx-popover"
      role="dialog"
      aria-label="Использование в контексте"
      @click.stop
    >
      <div class="ctx-title">Использование в контексте</div>
      <div class="ctx-summary">
        <span>{{ pct }}% заполнено</span>
        <span class="muted">
          ~{{ formatTokenCount(usage.usedTokens) }} /
          {{ formatTokenCount(usage.contextWindow) }} токенов
        </span>
      </div>

      <div class="ctx-bar">
        <div
          class="ctx-bar-used"
          :style="{
            width:
              usage.contextWindow > 0
                ? `${Math.min(100, (usage.usedTokens / usage.contextWindow) * 100)}%`
                : '0%',
          }"
        >
          <div
            v-for="cat in usage.categories"
            :key="cat.id"
            class="ctx-seg"
            :style="{
              width:
                usage.usedTokens > 0
                  ? `${(cat.tokens / usage.usedTokens) * 100}%`
                  : '0%',
              backgroundColor: cat.color,
            }"
            :title="`${cat.label}: ${formatTokenCount(cat.tokens)}`"
          />
        </div>
        <div
          v-if="usage.reservedTokens > 0"
          class="ctx-bar-reserve"
          :style="{
            width:
              usage.contextWindow > 0
                ? `${Math.min(100, (usage.reservedTokens / usage.contextWindow) * 100)}%`
                : '0%',
            backgroundColor: RESERVE_COLOR,
          }"
          :title="`Резерв для ответа модели: ${formatTokenCount(usage.reservedTokens)}`"
        />
      </div>

      <div class="ctx-list">
        <div v-for="cat in usage.categories" :key="cat.id" class="ctx-row">
          <div class="ctx-row-label">
            <span class="swatch" :style="{ backgroundColor: cat.color }" />
            <span>{{ cat.label }}</span>
          </div>
          <span class="nums">{{ formatTokenCount(cat.tokens) }}</span>
        </div>
        <div class="ctx-row reserve">
          <div class="ctx-row-label">
            <span class="swatch" :style="{ backgroundColor: RESERVE_COLOR }" />
            <span>Резерв для ответа модели</span>
          </div>
          <span class="nums">{{ formatTokenCount(usage.reservedTokens) }}</span>
        </div>
      </div>
    </div>
  </div>
</template>

<style scoped lang="scss">
.ctx-usage {
  position: relative;
  flex-shrink: 0;
}

.icon-btn {
  box-sizing: border-box;
  display: grid;
  place-items: center;
  width: 30px;
  height: 30px;
  padding: 0;
  border: 0;
  background: transparent;
  color: var(--fg-2);
  border-radius: 999px;
  cursor: pointer;
}

.icon-btn:hover:not(:disabled) {
  color: var(--fg-0);
  background: var(--bg-hover);
}

.icon-btn:disabled {
  opacity: 0.45;
  cursor: default;
}

.ring {
  display: block;
}

.ctx-popover {
  position: absolute;
  right: 0;
  bottom: calc(100% + 8px);
  z-index: 8;
  width: min(280px, 70vw);
  max-height: min(360px, 50vh);
  overflow-y: auto;
  padding: 12px;
  background: var(--bg-1);
  border: 1px solid var(--border);
  border-radius: var(--radius-md);
  box-shadow: var(--shadow);
  text-align: left;
}

.ctx-title {
  margin-bottom: 8px;
  font-size: 0.9rem;
  font-weight: 600;
  color: var(--fg-0);
}

.ctx-summary {
  display: flex;
  align-items: baseline;
  justify-content: space-between;
  gap: 8px;
  margin-bottom: 8px;
  font-size: 0.78rem;
  color: var(--fg-2);
}

.ctx-summary .muted {
  color: var(--fg-3);
  white-space: nowrap;
}

.ctx-bar {
  position: relative;
  height: 6px;
  width: 100%;
  margin-bottom: 12px;
  overflow: hidden;
  border-radius: 999px;
  background: var(--bg-2);
}

.ctx-bar-used {
  position: absolute;
  inset: 0 auto 0 0;
  display: flex;
  min-width: 0;
}

.ctx-seg {
  height: 100%;
  min-width: 0;
}

.ctx-bar-reserve {
  position: absolute;
  inset: 0 0 0 auto;
}

.ctx-list {
  display: flex;
  flex-direction: column;
  gap: 6px;
}

.ctx-row {
  display: flex;
  align-items: center;
  justify-content: space-between;
  gap: 8px;
  font-size: 0.78rem;
}

.ctx-row.reserve {
  margin-top: 4px;
  padding-top: 8px;
  border-top: 1px solid var(--border);
}

.ctx-row-label {
  display: flex;
  align-items: center;
  gap: 8px;
  min-width: 0;
  color: var(--fg-2);
}

.swatch {
  flex-shrink: 0;
  width: 8px;
  height: 8px;
  border-radius: 2px;
}

.nums {
  flex-shrink: 0;
  font-variant-numeric: tabular-nums;
  color: var(--fg-3);
}
</style>
