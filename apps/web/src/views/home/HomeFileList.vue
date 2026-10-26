<script setup lang="ts">
import type { DocItem } from '@/api/client';

defineProps<{
  groups: Array<{ key: string; label: string; items: DocItem[] }>;
  empty: boolean;
  menuOpenId: string | null;
}>();

const emit = defineEmits<{
  open: [doc: DocItem];
  'toggle-menu': [id: string, event: Event];
  download: [doc: DocItem];
  rename: [doc: DocItem];
  remove: [doc: DocItem];
}>();

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
</script>

<template>
  <div class="recent-block">
    <h2 class="section-title">Недавние</h2>

    <div v-if="empty" class="empty">
      Пока нет файлов. Создайте таблицу или документ выше — или откройте файл с
      компьютера.
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
            <button type="button" class="file-main" @click="emit('open', doc)">
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
                @click="emit('toggle-menu', doc.id, $event)"
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
                <button type="button" @click="emit('download', doc)">
                  Скачать
                </button>
                <button type="button" @click="emit('rename', doc)">
                  Переименовать
                </button>
                <button
                  type="button"
                  class="danger"
                  @click="emit('remove', doc)"
                >
                  Удалить
                </button>
              </div>
            </div>
          </li>
        </ul>
      </div>
    </template>
  </div>
</template>

<style scoped lang="scss">
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
