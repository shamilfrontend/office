<script setup lang="ts">
import { computed, onMounted, ref } from 'vue';
import { useRouter } from 'vue-router';
import { api, type DocItem } from '@/api/client';
import { uiConfirm } from '@/composables/useConfirm';
import { useAuthStore } from '@/stores/auth';
import { useFileTabsStore } from '@/stores/fileTabs';
import HomeCreateTiles from './home/HomeCreateTiles.vue';
import HomeFileList from './home/HomeFileList.vue';
import HomeTopBar from './home/HomeTopBar.vue';

type FileGroup = { key: string; label: string; items: DocItem[] };

const auth = useAuthStore();
const fileTabs = useFileTabsStore();
const router = useRouter();
const files = ref<DocItem[]>([]);
const query = ref('');
const busy = ref(false);
const uploading = ref(false);
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
  await router.push({ name: 'edit', params: { id: doc.id } });
}

async function create(type: 'SHEET' | 'DOC' | 'SLIDE') {
  busy.value = true;
  try {
    const doc = await api.createFile(type);
    await openDoc(doc);
  } finally {
    busy.value = false;
  }
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
    if (lower.endsWith('.pptx')) {
      typeHint = 'SLIDE';
    } else if (
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
  let format: 'xlsx' | 'docx' | 'pptx' = 'docx';
  if (doc.type === 'SHEET') format = 'xlsx';
  else if (doc.type === 'SLIDE') format = 'pptx';
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

async function logout() {
  profileMenuOpen.value = false;
  const ok = await uiConfirm({
    title: 'Выйти из аккаунта?',
    message: 'Сессия будет завершена на этом устройстве.',
    confirmLabel: 'Выйти',
  });
  if (!ok) return;
  await auth.logout();
  await router.push({ name: 'landing' });
}
</script>

<template>
  <div class="home" @click="closeMenu">
    <HomeTopBar
      :query="query"
      :profile-menu-open="profileMenuOpen"
      :avatar-url="auth.me?.avatarUrl"
      :display-name="auth.me?.displayName"
      :email="auth.me?.email"
      @update:query="query = $event"
      @toggle-profile="toggleProfileMenu"
      @settings="openSettings"
      @logout="logout"
    />

    <section class="content">
      <HomeCreateTiles
        :busy="busy"
        :uploading="uploading"
        @create="create"
        @local-file="onLocalFile"
      />
      <HomeFileList
        :groups="groups"
        :empty="!filtered.length"
        :menu-open-id="menuOpenId"
        @open="openDoc"
        @toggle-menu="toggleMenu"
        @download="download"
        @rename="rename"
        @remove="remove"
      />
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
.content {
  flex: 1;
  min-height: 0;
  overflow: auto;
  padding: 28px 32px 64px;
  background: var(--bg-1);
  border-top: 0;
}

@media (max-width: 720px) {
  .content {
    padding-left: 16px;
    padding-right: 16px;
  }
}
</style>
