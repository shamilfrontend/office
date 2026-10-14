<script setup lang="ts">
import { computed, onMounted, ref } from 'vue'
import AppLogo from '@/components/AppLogo.vue'
import { useAuthStore } from '@/stores/auth'
import { useRoute, useRouter } from 'vue-router'

const auth = useAuthStore()
const route = useRoute()
const router = useRouter()
const authError = ref('')

const yandexEnabled = computed(() => auth.authConfig.yandexEnabled)

onMounted(() => {
  const err = route.query.authError
  if (typeof err === 'string' && err) {
    authError.value = err
    void router.replace({ query: { ...route.query, authError: undefined } })
  }
})

function loginYandex() {
  window.location.href = '/auth/yandex'
}
</script>

<template>
  <div class="landing">
    <main class="gate">
      <div class="brand-row">
        <AppLogo :size="72" :with-label="false" />
        <p class="brand">Office</p>
      </div>
      <h1>Офис в браузере с AI-помощником</h1>
      <p class="lead">
        Таблицы и документы в браузере. Свои файлы на сервере, AI — по желанию.
      </p>
      <p v-if="authError" class="error">{{ authError }}</p>
      <div class="actions">
        <button
          v-if="yandexEnabled"
          class="ui-btn ui-btn--primary ui-btn--lg"
          type="button"
          @click="loginYandex"
        >
          Войти через Яндекс
        </button>
      </div>
    </main>
  </div>
</template>

<style scoped lang="scss">
.landing {
  min-height: 100%;
  display: grid;
  place-items: center;
  padding: 32px 20px;
  background: linear-gradient(180deg, var(--bg-1) 0%, var(--bg-0) 100%);
}

.gate {
  width: min(100%, 28rem);
  text-align: center;
  animation: rise-in 0.35s ease-out both;
}

.brand-row {
  display: inline-flex;
  align-items: center;
  justify-content: center;
  gap: 16px;
  margin: 0 0 20px;
}

.brand {
  margin: 0;
  font-family: var(--font-display);
  font-weight: 700;
  font-size: clamp(2.75rem, 6vw, 3.5rem);
  line-height: 1;
  letter-spacing: 0.03em;
  color: var(--fg-0);
}

h1 {
  margin: 0 0 12px;
  font-family: var(--font-display);
  font-weight: 600;
  font-size: 1.25rem;
  line-height: 1.3;
  letter-spacing: -0.02em;
  color: var(--fg-1);
}

.lead {
  margin: 0 auto 28px;
  max-width: 28ch;
  color: var(--fg-2);
  font-size: 1.05rem;
  line-height: 1.5;
}

.error {
  margin: -12px auto 24px;
  max-width: 36ch;
  color: var(--danger);
  font-size: 0.95rem;
  line-height: 1.4;
}

.actions {
  display: flex;
  justify-content: center;
  gap: 12px;
  flex-wrap: wrap;
}

@keyframes rise-in {
  from {
    opacity: 0;
    transform: translateY(8px);
  }
  to {
    opacity: 1;
    transform: translateY(0);
  }
}

@media (prefers-reduced-motion: reduce) {
  .gate {
    animation: none;
  }
}
</style>
