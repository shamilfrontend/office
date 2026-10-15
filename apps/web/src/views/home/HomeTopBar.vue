<script setup lang="ts">
import AppLogo from '@/components/AppLogo.vue';

defineProps<{
  query: string;
  profileMenuOpen: boolean;
  avatarUrl?: string | null;
  displayName?: string | null;
  email?: string | null;
}>();

const emit = defineEmits<{
  'update:query': [value: string];
  'toggle-profile': [event: Event];
  settings: [];
  logout: [];
}>();
</script>

<template>
  <header class="bar">
    <div class="brand">
      <AppLogo :size="28" />
    </div>
    <input
      :value="query"
      class="search"
      placeholder="Поиск файлов"
      @input="emit('update:query', ($event.target as HTMLInputElement).value)"
    />
    <div class="right">
      <div class="profile-wrap">
        <button
          type="button"
          class="avatar"
          aria-label="Меню аккаунта"
          :aria-expanded="profileMenuOpen"
          @click="emit('toggle-profile', $event)"
        >
          <img v-if="avatarUrl" :src="avatarUrl" alt="" />
          <span v-else>{{ displayName?.[0] || 'U' }}</span>
        </button>
        <div
          v-if="profileMenuOpen"
          class="ui-dropdown profile-menu"
          @click.stop
        >
          <div class="profile-meta">
            <div class="profile-name">
              {{ displayName || 'Аккаунт' }}
            </div>
            <div v-if="email" class="profile-email">
              {{ email }}
            </div>
          </div>
          <button type="button" @click="emit('settings')">Настройки</button>
          <button type="button" class="danger" @click="emit('logout')">
            Выйти
          </button>
        </div>
      </div>
    </div>
  </header>
</template>

<style scoped lang="scss">
.bar {
  display: grid;
  grid-template-columns: 1fr minmax(220px, 560px) 1fr;
  align-items: center;
  gap: 16px;
  height: var(--app-header-height);
  box-sizing: border-box;
  padding: 0 var(--app-header-padding-x);
  border-bottom: 1px solid var(--border);
  background: var(--bg-1);
}
.brand {
  display: inline-flex;
  align-items: center;
  justify-self: start;
  height: 28px;
  line-height: 0;
}
.search {
  width: 100%;
  height: var(--app-header-control-height);
  padding: 0 16px;
  border-radius: var(--radius-md);
  border: 1px solid var(--border);
  background: var(--bg-2);
  color: var(--fg-0);
}
.search:focus {
  outline: none;
  border-color: var(--accent);
  box-shadow: var(--focus-ring);
}
.right {
  display: flex;
  align-items: center;
  justify-content: flex-end;
  justify-self: end;
}
.profile-wrap {
  position: relative;
}
.avatar {
  width: 34px;
  height: 34px;
  border-radius: 50%;
  border: 1px solid var(--border);
  overflow: hidden;
  background: var(--bg-3);
  color: var(--fg-0);
  cursor: pointer;
  padding: 0;
  display: grid;
  place-items: center;
  font-size: 0.85rem;
  font-weight: 600;
}
.avatar:hover,
.avatar[aria-expanded='true'] {
  border-color: var(--accent);
}
.avatar:focus-visible {
  outline: none;
  box-shadow: var(--focus-ring);
}
.avatar img {
  width: 100%;
  height: 100%;
  object-fit: cover;
}
.profile-menu {
  right: 0;
  top: calc(100% + 8px);
  min-width: 200px;
}
.profile-meta {
  padding: 8px 10px 10px;
  border-bottom: 1px solid var(--border);
  margin-bottom: 4px;
}
.profile-name {
  font-size: 0.9rem;
  font-weight: 600;
  color: var(--fg-0);
  line-height: 1.3;
  overflow: hidden;
  text-overflow: ellipsis;
  white-space: nowrap;
}
.profile-email {
  margin-top: 2px;
  font-size: 0.8rem;
  color: var(--fg-2);
  overflow: hidden;
  text-overflow: ellipsis;
  white-space: nowrap;
}

@media (max-width: 720px) {
  .bar {
    padding-left: 16px;
    padding-right: 16px;
    grid-template-columns: auto 1fr auto;
  }
}
</style>
