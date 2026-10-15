<script setup lang="ts">
import { computed, onMounted, ref } from 'vue';
import { useRouter } from 'vue-router';
import { api, type DocItem } from '@/api/client';
import AppLogo from '@/components/AppLogo.vue';
import { uiConfirm } from '@/composables/useConfirm';
import { useAuthStore } from '@/stores/auth';
import { useFileTabsStore } from '@/stores/fileTabs';

type FileGroup = { key: string; label: string; items: DocItem[] };

const auth = useAuthStore();
const fileTabs = useFileTabsStore();
const router = useRouter();
const files = ref<DocItem[]>([]);
const query = ref('');
const busy = ref(false);
const uploading = ref(false);
const fileInput = ref<HTMLInputElement | null>(null);
const menuOpenId = ref<string | null>(null);
const profileMenuOpen = ref(false);

const filtered = computed(() => {
  const q = query.value.trim().toLowerCase();
  if (!q) return files.value;
  return files.value.filter((f) => f.name.toLowerCase().includes(q));
});

const groups = computed((): FileGroup[] => {
  const now = new Date();
  const startOfToday = new Date(
    now.getFullYear(),
    now.getMonth(),
    now.getDate(),
  ).getTime();
  const weekAgo = startOfToday - 6 * 86400000;
  const monthAgo = startOfToday - 29 * 86400000;

  const buckets: Record<string, DocItem[]> = {
    today: [],
    week: [],
    month: [],
    earlier: [],
  };

  filtered.value.forEach((doc) => {
    const t = new Date(doc.updatedAt).getTime();
    if (t >= startOfToday) buckets.today.push(doc);
    else if (t >= weekAgo) buckets.week.push(doc);
    else if (t >= monthAgo) buckets.month.push(doc);
    else buckets.earlier.push(doc);
  });

  const order: Array<{ key: string; label: string }> = [
    { key: 'today', label: 'Сегодня' },
    { key: 'week', label: 'За неделю' },
    { key: 'month', label: 'За месяц' },
    { key: 'earlier', label: 'Ранее' },
  ];

  return order
    .map(({ key, label }) => ({ key, label, items: buckets[key] }))
    .filter((g) => g.items.length > 0);
});

async function load() {
  files.value = await api.listFiles();
}

onMounted(async () => {
  auth.openOnboardingIfNeeded();
  await load();
});

async function openDoc(doc: Pick<DocItem, 'id' | 'name' | 'type'>) {
  fileTabs.open(doc);
  await router.push(`/edit/${doc.id}`);
}

async function create(type: 'SHEET' | 'DOC') {
  busy.value = true;
  try {
    const doc = await api.createFile(type);
    await openDoc(doc);
  } finally {
    busy.value = false;
  }
}

function pickLocal() {
  fileInput.value?.click();
}

async function onLocalFile(e: Event) {
  const input = e.target as HTMLInputElement;
  const file = input.files?.[0];
  if (!file) return;
  busy.value = true;
  uploading.value = true;
  try {
    let upload = file;
    let typeHint: string | undefined;
    const lower = file.name.toLowerCase();
    if (
      lower.endsWith('.xlsx') ||
      lower.endsWith('.xls') ||
      lower.endsWith('.csv')
    ) {
      const { spreadsheetFileToUpload } = await import('@/univer/excelImport');
      upload = await spreadsheetFileToUpload(file);
      typeHint = 'SHEET';
    }
    const doc = await api.uploadFile(upload, typeHint);
    await openDoc(doc);
  } catch (err) {
    alert(err instanceof Error ? err.message : 'Ошибка загрузки');
  } finally {
    busy.value = false;
    uploading.value = false;
    input.value = '';
  }
}

async function rename(doc: DocItem) {
  menuOpenId.value = null;
  const name = prompt('Новое имя', doc.name);
  if (!name || name === doc.name) return;
  await api.renameFile(doc.id, name);
  fileTabs.rename(doc.id, name);
  await load();
}

async function remove(doc: DocItem) {
  menuOpenId.value = null;
  if (!window.confirm(`Удалить «${doc.name}»?`)) return;
  await api.deleteFile(doc.id);
  fileTabs.remove(doc.id);
  await load();
}

async function download(doc: DocItem) {
  menuOpenId.value = null;
  const format = doc.type === 'SHEET' ? 'xlsx' : 'docx';
  window.open(api.downloadUrl(doc.id, format), '_blank');
}

function toggleMenu(id: string, e: Event) {
  e.stopPropagation();
  profileMenuOpen.value = false;
  menuOpenId.value = menuOpenId.value === id ? null : id;
}

function toggleProfileMenu(e: Event) {
  e.stopPropagation();
  menuOpenId.value = null;
  profileMenuOpen.value = !profileMenuOpen.value;
}

function closeMenu() {
  menuOpenId.value = null;
  profileMenuOpen.value = false;
}

function openSettings() {
  profileMenuOpen.value = false;
  auth.openSettings();
}

function formatDate(iso: string) {
  const d = new Date(iso);
  const now = new Date();
  const startOfToday = new Date(
    now.getFullYear(),
    now.getMonth(),
    now.getDate(),
  ).getTime();
  const t = d.getTime();
  if (t >= startOfToday) {
    return d.toLocaleTimeString('ru-RU', {
      hour: '2-digit',
      minute: '2-digit',
    });
  }
  if (d.getFullYear() === now.getFullYear()) {
    return d.toLocaleDateString('ru-RU', { day: 'numeric', month: 'short' });
  }
  return d.toLocaleDateString('ru-RU', {
    day: 'numeric',
    month: 'short',
    year: 'numeric',
  });
}

function formatSize(bytes: number) {
  if (!bytes || bytes < 0) return '—';
  if (bytes < 1024) return `${bytes} Б`;
  if (bytes < 1024 * 1024)
    return `${(bytes / 1024).toFixed(bytes < 10 * 1024 ? 1 : 0)} КБ`;
  if (bytes < 1024 * 1024 * 1024) {
    return `${(bytes / (1024 * 1024)).toFixed(bytes < 10 * 1024 * 1024 ? 1 : 0)} МБ`;
  }
  return `${(bytes / (1024 * 1024 * 1024)).toFixed(1)} ГБ`;
}

async function logout() {
  profileMenuOpen.value = false;
  const ok = await uiConfirm({
    title: 'Выйти из аккаунта?',
    message: 'Сессия будет завершена на этом устройстве.',
    confirmLabel: 'Выйти',
  });
  if (!ok) return;
  await auth.logout();
  await router.push('/');
}
</script>

<template>
  <div class="home" @click="closeMenu">
    <header class="bar">
      <div class="brand">
        <AppLogo :size="28" />
      </div>
      <input v-model="query" class="search" placeholder="Поиск файлов" />
      <div class="right">
        <div class="profile-wrap">
          <button
            type="button"
            class="avatar"
            aria-label="Меню аккаунта"
            :aria-expanded="profileMenuOpen"
            @click="toggleProfileMenu"
          >
            <img v-if="auth.me?.avatarUrl" :src="auth.me.avatarUrl" alt="" />
            <span v-else>{{ auth.me?.displayName?.[0] || 'U' }}</span>
          </button>
          <div
            v-if="profileMenuOpen"
            class="ui-dropdown profile-menu"
            @click.stop
          >
            <div class="profile-meta">
              <div class="profile-name">
                {{ auth.me?.displayName || 'Аккаунт' }}
              </div>
              <div v-if="auth.me?.email" class="profile-email">
                {{ auth.me.email }}
              </div>
            </div>
            <button type="button" @click="openSettings">Настройки</button>
            <button type="button" class="danger" @click="logout">Выйти</button>
          </div>
        </div>
      </div>
    </header>

    <section class="content">
      <div class="create-block">
        <h2 class="section-title">Создать</h2>
        <div class="tiles">
          <button
            type="button"
            class="tile"
            :disabled="busy"
            @click="create('SHEET')"
          >
            <span class="tile-icon sheet" aria-hidden="true">
              <svg viewBox="0 0 24 24" width="22" height="22">
                <rect
                  x="3"
                  y="4"
                  width="18"
                  height="16"
                  rx="2"
                  fill="currentColor"
                  opacity="0.18"
                />
                <path
                  d="M3 9h18M9 4v16"
                  stroke="currentColor"
                  stroke-width="1.5"
                  fill="none"
                />
              </svg>
            </span>
            <span class="tile-text">
              <span class="tile-label">Новая таблица</span>
            </span>
          </button>
          <button
            type="button"
            class="tile"
            :disabled="busy"
            @click="create('DOC')"
          >
            <span class="tile-icon doc" aria-hidden="true">
              <svg viewBox="0 0 24 24" width="22" height="22">
                <path
                  d="M6 3h8l4 4v14a1 1 0 0 1-1 1H6a1 1 0 0 1-1-1V4a1 1 0 0 1 1-1z"
                  fill="currentColor"
                  opacity="0.18"
                />
                <path
                  d="M14 3v4h4M8 12h8M8 16h5"
                  stroke="currentColor"
                  stroke-width="1.5"
                  fill="none"
                  stroke-linecap="round"
                />
              </svg>
            </span>
            <span class="tile-text">
              <span class="tile-label">Новый документ</span>
            </span>
          </button>
          <button
            type="button"
            class="tile"
            :disabled="busy || uploading"
            :aria-busy="uploading"
            @click="pickLocal"
          >
            <span class="tile-icon upload" aria-hidden="true">
              <span v-if="uploading" class="tile-spinner" />
              <svg v-else viewBox="0 0 24 24" width="22" height="22">
                <path
                  d="M12 16V6M12 6l-4 4M12 6l4 4M5 18v1a1 1 0 0 0 1 1h12a1 1 0 0 0 1-1v-1"
                  stroke="currentColor"
                  stroke-width="1.6"
                  stroke-linecap="round"
                  stroke-linejoin="round"
                  fill="none"
                />
              </svg>
            </span>
            <span class="tile-text">
              <span class="tile-label">{{
                uploading ? 'Открываем…' : 'Открыть с компьютера'
              }}</span>
            </span>
          </button>
          <input
            ref="fileInput"
            type="file"
            accept=".xlsx,.xls,.docx,.csv,.json"
            hidden
            @change="onLocalFile"
          />
        </div>
      </div>

      <div class="recent-block">
        <h2 class="section-title">Недавние</h2>

        <div v-if="!filtered.length" class="empty">
          Пока нет файлов. Создайте таблицу или документ выше — или откройте
          файл с компьютера.
        </div>

        <template v-else>
          <div class="list-head" aria-hidden="true">
            <span class="col-icon" />
            <span class="col-name">Название</span>
            <span class="col-size">Размер</span>
            <span class="col-date">Изменён</span>
            <span class="col-menu" />
          </div>

          <div v-for="group in groups" :key="group.key" class="group">
            <h3 class="group-label">
              {{ group.label }}
            </h3>
            <ul class="file-list">
              <li v-for="doc in group.items" :key="doc.id" class="file-item">
                <button type="button" class="file-main" @click="openDoc(doc)">
                  <span
                    class="type-icon"
                    :class="doc.type === 'SHEET' ? 'sheet' : 'doc'"
                    aria-hidden="true"
                  >
                    <svg
                      v-if="doc.type === 'SHEET'"
                      viewBox="0 0 24 24"
                      width="16"
                      height="16"
                    >
                      <rect
                        x="3"
                        y="4"
                        width="18"
                        height="16"
                        rx="2"
                        fill="currentColor"
                        opacity="0.2"
                      />
                      <path
                        d="M3 9h18M9 4v16"
                        stroke="currentColor"
                        stroke-width="1.5"
                        fill="none"
                      />
                    </svg>
                    <svg v-else viewBox="0 0 24 24" width="16" height="16">
                      <path
                        d="M6 3h8l4 4v14a1 1 0 0 1-1 1H6a1 1 0 0 1-1-1V4a1 1 0 0 1 1-1z"
                        fill="currentColor"
                        opacity="0.2"
                      />
                      <path
                        d="M14 3v4h4M8 12h8M8 16h5"
                        stroke="currentColor"
                        stroke-width="1.5"
                        fill="none"
                        stroke-linecap="round"
                      />
                    </svg>
                  </span>
                  <span class="file-name">{{ doc.name }}</span>
                  <span class="file-size">{{ formatSize(doc.size) }}</span>
                  <span class="file-date">{{ formatDate(doc.updatedAt) }}</span>
                </button>
                <div class="file-actions">
                  <button
                    type="button"
                    class="more"
                    aria-label="Действия"
                    :aria-expanded="menuOpenId === doc.id"
                    @click="toggleMenu(doc.id, $event)"
                  >
                    <svg
                      viewBox="0 0 24 24"
                      width="16"
                      height="16"
                      aria-hidden="true"
                    >
                      <circle cx="12" cy="6" r="1.5" fill="currentColor" />
                      <circle cx="12" cy="12" r="1.5" fill="currentColor" />
                      <circle cx="12" cy="18" r="1.5" fill="currentColor" />
                    </svg>
                  </button>
                  <div
                    v-if="menuOpenId === doc.id"
                    class="ui-dropdown file-menu"
                    @click.stop
                  >
                    <button type="button" @click="download(doc)">
                      Скачать
                    </button>
                    <button type="button" @click="rename(doc)">
                      Переименовать
                    </button>
                    <button type="button" class="danger" @click="remove(doc)">
                      Удалить
                    </button>
                  </div>
                </div>
              </li>
            </ul>
          </div>
        </template>
      </div>
    </section>
  </div>
</template>

<style scoped lang="scss">
.home {
  height: 100%;
  display: flex;
  flex-direction: column;
  background: var(--bg-0);
}
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
.content {
  flex: 1;
  min-height: 0;
  overflow: auto;
  padding: 28px 32px 64px;
  background: var(--bg-1);
  border-top: 0;
}
.create-block {
  margin-bottom: 32px;
}
.recent-block {
  width: 100%;
}
.section-title {
  margin: 0 0 14px;
  font-family: var(--font-display);
  font-size: 1.05rem;
  letter-spacing: -0.02em;
  color: var(--fg-0);
  font-weight: 600;
}
.tiles {
  display: grid;
  grid-template-columns: repeat(auto-fill, minmax(220px, 1fr));
  gap: 10px;
  max-width: 920px;
}
.tile {
  display: flex;
  align-items: center;
  gap: 12px;
  min-height: 64px;
  padding: 12px 14px;
  border: 1px solid var(--border);
  border-radius: var(--radius-md);
  background: var(--bg-2);
  color: var(--fg-0);
  cursor: pointer;
  text-align: left;
  transition:
    background 0.12s ease,
    border-color 0.12s ease;
}
.tile:hover:not(:disabled) {
  background: var(--bg-hover);
  border-color: var(--fg-3);
}
.tile:focus-visible {
  outline: none;
  box-shadow: var(--focus-ring);
}
.tile:disabled {
  cursor: not-allowed;
  opacity: 0.65;
}
.tile-icon {
  flex-shrink: 0;
  width: 40px;
  height: 40px;
  border-radius: var(--radius-sm);
  display: grid;
  place-items: center;
}
.tile-icon.sheet {
  background: color-mix(in srgb, var(--success) 14%, transparent);
  color: var(--success);
}
.tile-icon.doc {
  background: color-mix(in srgb, var(--accent) 14%, transparent);
  color: var(--accent);
}
.tile-icon.upload {
  background: var(--bg-3);
  color: var(--fg-2);
}
.tile-spinner {
  width: 20px;
  height: 20px;
  border: 2px solid color-mix(in srgb, currentColor 25%, transparent);
  border-top-color: currentColor;
  border-radius: 50%;
  animation: tile-spin 0.7s linear infinite;
}
@keyframes tile-spin {
  to {
    transform: rotate(360deg);
  }
}
.tile-text {
  display: grid;
  gap: 2px;
  min-width: 0;
  flex: 1;
}
.tile-label {
  font-size: 0.9rem;
  font-weight: 600;
  color: var(--fg-0);
  line-height: 1.25;
}
.empty {
  color: var(--fg-3);
  padding: 8px 0 4px;
  font-size: 0.95rem;
  line-height: 1.45;
  max-width: 40ch;
}
.list-head {
  display: grid;
  grid-template-columns: 28px minmax(0, 1fr) 88px 120px 36px;
  gap: 10px;
  align-items: center;
  padding: 6px 0 8px 8px;
  border-bottom: 1px solid var(--border);
  color: var(--fg-3);
  font-size: 0.75rem;
  font-weight: 500;
}
.col-size,
.col-date {
  text-align: right;
}
.group-label {
  margin: 14px 0 2px;
  padding: 0 8px;
  font-size: 0.75rem;
  font-weight: 600;
  color: var(--fg-3);
}
.file-list {
  list-style: none;
  margin: 0;
  padding: 0;
}
.file-item {
  display: grid;
  grid-template-columns: minmax(0, 1fr) 36px;
  align-items: center;
  border-radius: var(--radius-sm);
  position: relative;
  min-height: 42px;
}
.file-item:hover {
  background: var(--bg-hover);
}
.file-main {
  min-width: 0;
  display: grid;
  grid-template-columns: 28px minmax(0, 1fr) 88px 120px;
  align-items: center;
  gap: 10px;
  padding: 6px 0 6px 8px;
  border: 0;
  background: transparent;
  color: inherit;
  text-align: left;
  cursor: pointer;
  font: inherit;
  min-height: 42px;
}
.type-icon {
  width: 28px;
  height: 28px;
  border-radius: var(--radius-sm);
  display: grid;
  place-items: center;
}
.type-icon.sheet {
  background: color-mix(in srgb, var(--success) 14%, transparent);
  color: var(--success);
}
.type-icon.doc {
  background: color-mix(in srgb, var(--accent) 14%, transparent);
  color: var(--accent);
}
.file-name {
  font-weight: 500;
  font-size: 0.9375rem;
  color: var(--fg-0);
  white-space: nowrap;
  overflow: hidden;
  text-overflow: ellipsis;
}
.file-size,
.file-date {
  color: var(--fg-3);
  font-size: 0.8125rem;
  white-space: nowrap;
  text-align: right;
  font-variant-numeric: tabular-nums;
}
.file-actions {
  position: relative;
  display: grid;
  place-items: center;
}
.more {
  width: 32px;
  height: 32px;
  border: 0;
  border-radius: var(--radius-sm);
  background: transparent;
  color: var(--fg-2);
  cursor: pointer;
  display: grid;
  place-items: center;
  padding: 0;
  opacity: 0;
}
.file-item:hover .more,
.more[aria-expanded='true'] {
  opacity: 1;
}
.more:hover,
.more[aria-expanded='true'] {
  background: var(--bg-2);
  color: var(--fg-0);
}
.file-menu {
  right: 0;
  top: calc(100% - 2px);
  min-width: 168px;
}

@media (max-width: 720px) {
  .bar {
    padding-left: 16px;
    padding-right: 16px;
    grid-template-columns: auto 1fr auto;
  }
  .content {
    padding-left: 16px;
    padding-right: 16px;
  }
  .list-head,
  .file-main {
    grid-template-columns: 28px minmax(0, 1fr) 72px 88px;
  }
  .list-head {
    grid-template-columns: 28px minmax(0, 1fr) 72px 88px 36px;
  }
  .more {
    opacity: 1;
  }
}
</style>
