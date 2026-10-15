<script setup lang="ts">
import { useAuthStore } from '@/stores/auth';
import OnboardingModal from '@/components/OnboardingModal.vue';
import SettingsModal from '@/components/SettingsModal.vue';
import UiConfirm from '@/components/UiConfirm.vue';

const auth = useAuthStore();
</script>

<template>
  <router-view />
  <OnboardingModal v-if="auth.showOnboarding" />
  <SettingsModal v-if="auth.showSettings" />
  <UiConfirm />
</template>
