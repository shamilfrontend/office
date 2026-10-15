import { defineStore } from 'pinia';
import { computed, ref } from 'vue';
import { api, type Me } from '@/api/client';

export const useAuthStore = defineStore('auth', () => {
  const me = ref<Me | null>(null);
  const loaded = ref(false);
  const authConfig = ref({ yandexEnabled: false });
  const showOnboarding = ref(false);
  const showSettings = ref(false);

  const isAuthenticated = computed(() => Boolean(me.value));

  let initPromise: Promise<void> | null = null;

  async function init(): Promise<void> {
    if (loaded.value) return undefined;
    if (initPromise) return initPromise;

    initPromise = (async () => {
      try {
        authConfig.value = await api.authConfig();
      } catch {
        /* ignore */
      }
      try {
        me.value = await api.me();
      } catch {
        me.value = null;
      } finally {
        loaded.value = true;
        initPromise = null;
      }
    })();

    return initPromise;
  }

  async function logout() {
    await api.logout();
    me.value = null;
  }

  async function refreshMe() {
    me.value = await api.me();
  }

  async function skipOnboarding() {
    me.value = await api.skipOnboarding();
    showOnboarding.value = false;
  }

  function openOnboardingIfNeeded() {
    if (me.value && !me.value.hasAiProvider && !me.value.onboardingSkipped) {
      showOnboarding.value = true;
    }
  }

  function openSettings() {
    showSettings.value = true;
  }

  function closeSettings() {
    showSettings.value = false;
  }

  return {
    me,
    loaded,
    authConfig,
    showOnboarding,
    showSettings,
    isAuthenticated,
    init,
    logout,
    refreshMe,
    skipOnboarding,
    openOnboardingIfNeeded,
    openSettings,
    closeSettings,
  };
});
