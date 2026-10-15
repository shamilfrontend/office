<script setup lang="ts">
const emit = defineEmits<{
  settings: [];
}>();
</script>

<template>
  <div class="cta">
    <p class="cta-title">Подключите AI-провайдер</p>
    <p class="cta-text">
      Добавьте ключ в настройках, чтобы ассистент мог работать с документом.
    </p>
    <button
      type="button"
      class="ui-btn ui-btn--primary"
      @click="emit('settings')"
    >
      Открыть настройки
    </button>
  </div>
</template>

<style scoped lang="scss">
.cta {
  flex: 1;
  min-height: 0;
  display: flex;
  flex-direction: column;
  align-items: center;
  justify-content: center;
  gap: 10px;
  padding: 28px 24px;
  color: var(--fg-2);
  text-align: center;
}
.cta-title {
  margin: 0;
  color: var(--fg-0);
  font-size: 1rem;
  font-weight: 650;
}
.cta-text {
  margin: 0 0 4px;
  max-width: 260px;
  font-size: 0.88rem;
  color: var(--fg-3);
  line-height: 1.45;
}
</style>
