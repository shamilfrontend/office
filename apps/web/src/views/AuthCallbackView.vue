<script setup lang="ts">
import { onMounted } from 'vue'
import { useRoute, useRouter } from 'vue-router'
import { useAuthStore } from '@/stores/auth'

const route = useRoute()
const router = useRouter()
const auth = useAuthStore()

onMounted(async () => {
  await auth.refreshMe()
  if (route.query.onboarding === '1') {
    auth.showOnboarding = true
  } else {
    auth.openOnboardingIfNeeded()
  }
  await router.replace('/home')
})
</script>

<template>
  <div class="wrap">Вход…</div>
</template>

<style scoped>
.wrap {
  display: grid;
  place-items: center;
  min-height: 100%;
  color: var(--fg-2);
}
</style>
