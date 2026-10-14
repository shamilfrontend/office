<script setup lang="ts">
import { reactive, ref } from 'vue'
import { api } from '@/api/client'
import AppLogo from '@/components/AppLogo.vue'
import { useAuthStore } from '@/stores/auth'

const auth = useAuthStore()
const form = reactive({
  provider: 'DEEPSEEK',
  baseUrl: 'https://api.deepseek.com/v1',
  apiKey: '',
})
const status = ref('')
const busy = ref(false)

const providers = [
  { id: 'DEEPSEEK', label: 'DeepSeek' },
  { id: 'OLLAMA', label: 'Ollama' },
  { id: 'LMSTUDIO', label: 'LM Studio' },
  { id: 'OPENAI_COMPATIBLE', label: 'OpenAI-compatible' },
]

async function onProviderChange() {
  const defaults = await api.aiDefaults()
  const d = defaults[form.provider] as { baseUrl: string }
  form.baseUrl = d.baseUrl
  status.value = ''
}

async function test() {
  busy.value = true
  status.value = 'Проверка…'
  try {
    const res = await api.testProvider({
      provider: form.provider,
      baseUrl: form.baseUrl,
      apiKey: form.apiKey || undefined,
    })
    if (!res.ok) {
      status.value = res.error || 'Ошибка'
      return
    }
    status.value = `OK · ${res.models.length} моделей доступно`
  } catch (e) {
    status.value = e instanceof Error ? e.message : 'Ошибка'
  } finally {
    busy.value = false
  }
}

async function save() {
  busy.value = true
  try {
    const defaults = await api.aiDefaults()
    const d = defaults[form.provider] as { needsKey?: boolean } | undefined
    if (d?.needsKey && !form.apiKey.trim()) {
      status.value = 'API-ключ обязателен для этого провайдера.'
      return
    }
    await api.upsertProvider({
      provider: form.provider,
      baseUrl: form.baseUrl,
      apiKey: form.apiKey || undefined,
      enabled: true,
    })
    await auth.refreshMe()
    auth.showOnboarding = false
  } catch (e) {
    status.value = e instanceof Error ? e.message : 'Ошибка сохранения'
  } finally {
    busy.value = false
  }
}

async function skip() {
  await auth.skipOnboarding()
}
</script>

<template>
  <div class="ui-overlay overlay">
    <div class="ui-modal modal" role="dialog" aria-modal="true">
      <div class="modal-brand">
        <AppLogo :size="36" />
      </div>
      <h2>Подключить AI-помощника?</h2>
      <p class="hint">Можно пропустить и настроить позже в Настройках. Модель выберете в чате.</p>

      <label>
        Провайдер
        <select v-model="form.provider" @change="onProviderChange">
          <option v-for="p in providers" :key="p.id" :value="p.id">
            {{ p.label }}
          </option>
        </select>
      </label>
      <label>
        Base URL
        <input v-model="form.baseUrl" type="text" />
      </label>
      <label>
        API key
        <input
          v-model="form.apiKey"
          type="password"
          :placeholder="
            form.provider === 'DEEPSEEK'
              ? 'обязателен для DeepSeek'
              : 'необязательно для локальных'
          "
        />
      </label>

      <p v-if="status" class="status">{{ status }}</p>

      <div class="row">
        <button type="button" class="ui-btn ui-btn--ghost" :disabled="busy" @click="test">
          Проверить
        </button>
        <button type="button" class="ui-btn ui-btn--ghost" :disabled="busy" @click="skip">
          Пропустить
        </button>
        <button type="button" class="ui-btn ui-btn--primary" :disabled="busy" @click="save">
          Подключить
        </button>
      </div>
    </div>
  </div>
</template>

<style scoped lang="scss">
.overlay {
  z-index: var(--z-settings);
}
.modal {
  width: min(480px, 100%);
}
.modal-brand {
  margin-bottom: 16px;
}
h2 {
  margin: 0 0 8px;
  font-family: var(--font-display);
  font-weight: 600;
  letter-spacing: -0.02em;
  color: var(--fg-0);
}
.hint {
  color: var(--fg-2);
  margin: 0 0 18px;
}
label {
  display: grid;
  gap: 6px;
  margin-bottom: 12px;
  color: var(--fg-2);
  font-size: 0.9rem;
}
input,
select {
  width: 100%;
  padding: 10px 12px;
  border-radius: var(--radius-sm);
  border: 1px solid var(--border);
  background: var(--bg-2);
  color: var(--fg-0);
}
input:focus-visible,
select:focus-visible {
  outline: none;
  box-shadow: var(--focus-ring);
}
.row {
  display: flex;
  gap: 8px;
  justify-content: flex-end;
  margin-top: 16px;
  flex-wrap: wrap;
}
.status {
  color: var(--fg-2);
  font-size: 0.9rem;
}
</style>
