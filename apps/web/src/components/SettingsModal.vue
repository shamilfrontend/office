<script setup lang="ts">
import { onMounted, onUnmounted, reactive, ref } from 'vue';
import { useRouter } from 'vue-router';
import { api } from '@/api/client';
import { uiConfirm } from '@/composables/useConfirm';
import { useAuthStore } from '@/stores/auth';
import { useTheme, type ThemePreference } from '@/composables/useTheme';
import { useAiToolPermissions } from '@/composables/useAiToolPermissions';

const auth = useAuthStore();
const router = useRouter();
const { preference: theme, setTheme } = useTheme();
const {
  permissions: aiPermissions,
  categories: aiPermissionCategories,
  setCategoryAllowed,
} = useAiToolPermissions();
const providers = ref<
  Array<{
    id: string;
    provider: string;
    baseUrl: string;
    model: string | null;
    enabled: boolean;
    hasApiKey: boolean;
  }>
>([]);
const form = reactive({
  id: '' as string | undefined,
  provider: 'DEEPSEEK',
  baseUrl: 'https://api.deepseek.com/v1',
  apiKey: '',
  hasApiKey: false,
});
const showProviderForm = ref(false);
const status = ref('');
const themeOptions: Array<{ value: ThemePreference; label: string }> = [
  { value: 'system', label: 'Системная' },
  { value: 'dark', label: 'Тёмная' },
  { value: 'light', label: 'Светлая' },
];

async function load() {
  providers.value = await api.listProviders();
}

function close() {
  auth.closeSettings();
}

function onKeydown(e: KeyboardEvent) {
  if (e.key === 'Escape') close();
}

function resetForm() {
  form.id = undefined;
  form.provider = 'DEEPSEEK';
  form.baseUrl = 'https://api.deepseek.com/v1';
  form.apiKey = '';
  form.hasApiKey = false;
}

function openProviderForm() {
  resetForm();
  showProviderForm.value = true;
  status.value = '';
}

function editProvider(id: string) {
  const provider = providers.value.find((p) => p.id === id);
  if (!provider) return;
  form.id = provider.id;
  form.provider = provider.provider;
  form.baseUrl = provider.baseUrl;
  form.apiKey = '';
  form.hasApiKey = provider.hasApiKey;
  showProviderForm.value = true;
  status.value = '';
}

function cancelProviderForm() {
  showProviderForm.value = false;
  resetForm();
  status.value = '';
}

async function onProviderTypeChange() {
  if (form.id) return;
  const defaults = await api.aiDefaults();
  const d = defaults[form.provider] as { baseUrl: string };
  form.baseUrl = d.baseUrl;
  status.value = '';
}

onMounted(() => {
  void load();
  window.addEventListener('keydown', onKeydown);
});

onUnmounted(() => {
  window.removeEventListener('keydown', onKeydown);
});

async function test() {
  status.value = 'Проверка…';
  try {
    if (form.id && !form.apiKey) {
      const res = await api.listProviderModels(form.id);
      status.value = `OK · ${res.models.length} моделей доступно`;
      return;
    }
    const res = await api.testProvider({
      provider: form.provider,
      baseUrl: form.baseUrl,
      apiKey: form.apiKey || undefined,
    });
    status.value = res.ok
      ? `OK · ${res.models.length} моделей доступно`
      : res.error || 'Ошибка';
  } catch (e) {
    status.value = e instanceof Error ? e.message : 'Ошибка';
  }
}

async function save() {
  try {
    await api.upsertProvider({
      ...(form.id ? { id: form.id } : {}),
      provider: form.provider,
      baseUrl: form.baseUrl,
      apiKey: form.apiKey || undefined,
      enabled: true,
    });
    resetForm();
    status.value = '';
    showProviderForm.value = false;
    await auth.refreshMe();
    await load();
  } catch (e) {
    status.value = e instanceof Error ? e.message : 'Ошибка сохранения';
  }
}

async function remove(id: string) {
  const provider = providers.value.find((p) => p.id === id);
  const ok = await uiConfirm({
    title: 'Удалить провайдер?',
    message: provider
      ? `Провайдер ${provider.provider} будет удалён.`
      : 'Провайдер будет удалён.',
    confirmLabel: 'Удалить',
  });
  if (!ok) return;
  await api.deleteProvider(id);
  if (form.id === id) cancelProviderForm();
  await auth.refreshMe();
  await load();
}

async function logout() {
  const ok = await uiConfirm({
    title: 'Выйти из аккаунта?',
    message: 'Сессия будет завершена на этом устройстве.',
    confirmLabel: 'Выйти',
  });
  if (!ok) return;
  auth.closeSettings();
  await auth.logout();
  await router.push('/');
}
</script>

<template>
  <div class="ui-overlay overlay" @click.self="close">
    <div
      class="modal"
      role="dialog"
      aria-modal="true"
      aria-labelledby="settings-title"
    >
      <header class="head">
        <h2 id="settings-title">Настройки</h2>
        <button
          type="button"
          class="ui-btn close"
          aria-label="Закрыть"
          @click="close"
        >
          ×
        </button>
      </header>

      <div class="body">
        <section>
          <h3>Профиль</h3>
          <div class="profile">
            <div class="profile-main">
              <img v-if="auth.me?.avatarUrl" :src="auth.me.avatarUrl" alt="" />
              <div>
                <div>{{ auth.me?.displayName }}</div>
                <div class="muted">
                  {{ auth.me?.email }}
                </div>
              </div>
            </div>
            <button type="button" class="ui-btn" @click="logout">Выйти</button>
          </div>
        </section>

        <section>
          <h3>Провайдеры</h3>
          <ul v-if="providers.length" class="list">
            <li v-for="p in providers" :key="p.id">
              <div>
                <strong>{{ p.provider }}</strong>
                <div class="muted">
                  {{ p.baseUrl }}
                </div>
                <div v-if="p.hasApiKey" class="muted">Ключ сохранён</div>
              </div>
              <div class="row">
                <button
                  type="button"
                  class="ui-btn"
                  @click="editProvider(p.id)"
                >
                  Изменить
                </button>
                <button type="button" class="ui-btn" @click="remove(p.id)">
                  Удалить
                </button>
              </div>
            </li>
          </ul>
          <div v-if="!showProviderForm" class="row">
            <button
              type="button"
              class="ui-btn ui-btn--primary"
              @click="openProviderForm"
            >
              Добавить провайдер
            </button>
          </div>
          <div v-else class="form">
            <label>
              Тип
              <select
                v-model="form.provider"
                :disabled="Boolean(form.id)"
                @change="onProviderTypeChange"
              >
                <option value="DEEPSEEK">DeepSeek</option>
                <option value="OLLAMA">Ollama</option>
                <option value="LMSTUDIO">LM Studio</option>
                <option value="OPENAI_COMPATIBLE">OpenAI-compatible</option>
              </select>
            </label>
            <label>
              Base URL
              <input v-model="form.baseUrl" />
            </label>
            <label>
              API key
              <input
                v-model="form.apiKey"
                type="password"
                :placeholder="
                  form.id && form.hasApiKey
                    ? 'оставьте пустым, чтобы не менять'
                    : 'обязателен для DeepSeek'
                "
              />
            </label>
            <div class="row">
              <button type="button" class="ui-btn" @click="test">
                Проверить
              </button>
              <button
                type="button"
                class="ui-btn ui-btn--primary"
                @click="save"
              >
                Сохранить
              </button>
              <button
                type="button"
                class="ui-btn ui-btn--ghost"
                @click="cancelProviderForm"
              >
                Отмена
              </button>
            </div>
            <p v-if="status" class="muted">
              {{ status }}
            </p>
          </div>
        </section>

        <section>
          <h3>Приложение</h3>
          <p class="muted">Язык интерфейса: Русский</p>
          <label class="theme-label">
            Тема
            <div class="theme-switch" role="radiogroup" aria-label="Тема">
              <button
                v-for="opt in themeOptions"
                :key="opt.value"
                type="button"
                role="radio"
                :aria-checked="theme === opt.value"
                :class="{ active: theme === opt.value }"
                @click="setTheme(opt.value)"
              >
                {{ opt.label }}
              </button>
            </div>
          </label>
        </section>

        <section>
          <h3>Разрешения AI</h3>
          <p class="muted">Не спрашивать подтверждение для выбранных правок</p>
          <ul class="perm-list">
            <li v-for="cat in aiPermissionCategories" :key="cat.id">
              <label class="perm-row">
                <input
                  type="checkbox"
                  :checked="aiPermissions[cat.id]"
                  @change="
                    setCategoryAllowed(
                      cat.id,
                      ($event.target as HTMLInputElement).checked,
                    )
                  "
                />
                <span>{{ cat.label }}</span>
              </label>
            </li>
          </ul>
        </section>
      </div>
    </div>
  </div>
</template>

<style scoped lang="scss">
.overlay {
  z-index: var(--z-settings);
}
.modal {
  width: min(640px, 100%);
  max-height: min(90vh, 900px);
  display: flex;
  flex-direction: column;
  background: var(--bg-1);
  border: 1px solid var(--border);
  border-radius: var(--radius-lg);
  box-shadow: var(--shadow);
  overflow: hidden;
  padding: 0;
}
.head {
  display: flex;
  align-items: center;
  justify-content: space-between;
  gap: 12px;
  padding: 16px 20px;
  border-bottom: 1px solid var(--border);
  flex-shrink: 0;
}
.head h2 {
  margin: 0;
  font-family: var(--font-display);
  font-size: 1.15rem;
  font-weight: 600;
  letter-spacing: -0.02em;
  color: var(--fg-0);
}
.close {
  width: 36px;
  height: 36px;
  padding: 0;
  font-size: 1.4rem;
  line-height: 1;
}
.body {
  padding: 16px 20px 24px;
  overflow-y: auto;
}
section {
  margin-bottom: 20px;
  padding: 14px;
  border: 1px solid var(--border);
  border-radius: var(--radius-md);
  background: var(--bg-2);
}
section:last-child {
  margin-bottom: 0;
}
h3 {
  margin: 0 0 12px;
  font-family: var(--font-display);
  font-size: 1rem;
  font-weight: 600;
  letter-spacing: -0.02em;
  color: var(--fg-0);
}
.profile {
  display: flex;
  align-items: center;
  justify-content: space-between;
  gap: 12px;
}
.profile-main {
  display: flex;
  align-items: center;
  gap: 12px;
  min-width: 0;
}
.profile img {
  width: 48px;
  height: 48px;
  border-radius: 50%;
}
.muted {
  color: var(--fg-3);
  font-size: 0.9rem;
}
.list {
  list-style: none;
  padding: 0;
  margin: 0 0 16px;
}
.list li {
  display: flex;
  align-items: center;
  justify-content: space-between;
  gap: 12px;
  padding: 10px 0;
  border-bottom: 1px solid var(--border);
}
.list .row .ui-btn,
.profile > .ui-btn {
  height: 32px;
  padding: 0 10px;
  font-size: 0.85rem;
}
.form {
  display: grid;
  gap: 10px;
}
label {
  display: grid;
  gap: 4px;
  color: var(--fg-2);
  font-size: 0.9rem;
}
input,
select {
  padding: 10px 12px;
  border-radius: var(--radius-sm);
  border: 1px solid var(--border);
  background: var(--bg-1);
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
  flex-wrap: wrap;
}
.theme-label {
  margin-top: 14px;
}
.theme-switch {
  display: flex;
  gap: 6px;
  flex-wrap: wrap;
}
.theme-switch button {
  flex: 1;
  min-width: 0;
  padding: 8px 10px;
  border-radius: var(--radius-sm);
  border: 1px solid var(--border);
  background: var(--bg-1);
  color: var(--fg-2);
  cursor: pointer;
}
.theme-switch button.active {
  background: var(--accent);
  border-color: var(--accent);
  color: #fff;
}
.theme-switch button:hover:not(.active) {
  background: var(--bg-hover);
  color: var(--fg-0);
}
.theme-switch button:focus-visible {
  outline: none;
  box-shadow: var(--focus-ring);
}
.perm-list {
  list-style: none;
  padding: 0;
  margin: 12px 0 0;
  display: grid;
  gap: 8px;
}
.perm-row {
  display: flex;
  align-items: center;
  gap: 10px;
  color: var(--fg-1);
  font-size: 0.9rem;
  cursor: pointer;
}
.perm-row input {
  width: 16px;
  height: 16px;
  margin: 0;
  accent-color: var(--accent);
  cursor: pointer;
}
</style>
