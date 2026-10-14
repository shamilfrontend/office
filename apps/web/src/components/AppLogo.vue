<script setup lang="ts">
withDefaults(
  defineProps<{
    size?: number
    withLabel?: boolean
  }>(),
  {
    size: 28,
    withLabel: true,
  },
)
</script>

<template>
  <span class="app-logo" :style="{ '--logo-size': `${size}px` }">
    <img
      class="app-logo__mark"
      src="/logo-192.png"
      alt=""
      :width="size"
      :height="size"
      decoding="async"
    />
    <span v-if="withLabel" class="app-logo__label">Office</span>
  </span>
</template>

<style scoped>
.app-logo {
  display: inline-flex;
  align-items: center;
  gap: calc(var(--logo-size) * 0.3);
  line-height: 1;
}

.app-logo__mark {
  width: var(--logo-size);
  height: var(--logo-size);
  border-radius: 18%;
  flex-shrink: 0;
  display: block;
}

.app-logo__label {
  font-weight: 700;
  letter-spacing: -0.02em;
}
</style>
