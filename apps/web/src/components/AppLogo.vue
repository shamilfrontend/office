<script setup lang="ts">
withDefaults(
  defineProps<{
    size?: number;
    withLabel?: boolean;
  }>(),
  {
    size: 28,
    withLabel: true,
  },
);
</script>

<template>
  <span class="app-logo" :style="{ '--logo-size': `${size}px` }">
    <svg
      class="app-logo__mark"
      xmlns="http://www.w3.org/2000/svg"
      viewBox="0 0 512 512"
      role="img"
      aria-hidden="true"
      :width="size"
      :height="size"
    >
      <rect width="512" height="512" rx="96" fill="#5686fe" />
      <rect x="96" y="96" width="140" height="140" rx="28" fill="#ffffff" />
      <rect x="276" y="96" width="140" height="140" rx="28" fill="#dce6ff" />
      <rect x="96" y="276" width="140" height="140" rx="28" fill="#dce6ff" />
      <rect x="276" y="276" width="140" height="140" rx="28" fill="#ffffff" />
    </svg>
    <span v-if="withLabel" class="app-logo__label">Office</span>
  </span>
</template>

<style scoped>
.app-logo {
  display: inline-flex;
  align-items: center;
  gap: calc(var(--logo-size) * 0.22);
  height: var(--logo-size);
  line-height: 1;
  font-family: var(--font-display);
  color: var(--fg-0);
}

.app-logo__mark {
  width: var(--logo-size);
  height: var(--logo-size);
  border-radius: 18%;
  flex-shrink: 0;
  display: block;
  overflow: hidden;
}

.app-logo__label {
  font-size: calc(var(--logo-size) * 0.64);
  font-weight: 700;
  letter-spacing: -0.03em;
  line-height: 1;
}
</style>
