<script setup lang="ts">
import { onMounted } from 'vue';
import { useRoute, useRouter } from 'vue-router';
import AppLogo from '@/components/AppLogo.vue';
import { useAuthStore } from '@/stores/auth';

const route = useRoute();
const router = useRouter();
const auth = useAuthStore();

onMounted(async () => {
  await auth.refreshMe();
  if (route.query.onboarding === '1') {
    auth.showOnboarding = true;
  } else {
    auth.openOnboardingIfNeeded();
  }
  await router.replace('/home');
});
</script>

<template>
  <div class="wrap">
    <AppLogo :size="40" :with-label="false" />
    <div class="row">
      <span class="spinner" aria-hidden="true" />
      <span>Вход…</span>
    </div>
  </div>
</template>

<style scoped>
.wrap {
  display: grid;
  place-items: center;
  place-content: center;
  gap: 20px;
  min-height: 100%;
  background: linear-gradient(180deg, var(--bg-1) 0%, var(--bg-0) 100%);
  color: var(--fg-2);
}
.row {
  display: inline-flex;
  align-items: center;
  gap: 10px;
  font-size: 0.95rem;
}
.spinner {
  width: 16px;
  height: 16px;
  border: 2px solid color-mix(in srgb, var(--fg-3) 35%, transparent);
  border-top-color: var(--accent);
  border-radius: 50%;
  animation: spin 0.7s linear infinite;
}
@keyframes spin {
  to {
    transform: rotate(360deg);
  }
}
@media (prefers-reduced-motion: reduce) {
  .spinner {
    animation: none;
    border-top-color: var(--fg-3);
  }
}
</style>
