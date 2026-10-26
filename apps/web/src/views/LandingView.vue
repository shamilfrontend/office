<script setup lang="ts">
import { computed, onMounted, ref } from 'vue';
import AppLogo from '@/components/AppLogo.vue';
import { useAuthStore } from '@/stores/auth';
import { useRoute, useRouter } from 'vue-router';

const auth = useAuthStore();
const route = useRoute();
const router = useRouter();
const authError = ref('');

const yandexEnabled = computed(() => auth.authConfig.yandexEnabled);

const features = [
  {
    title: 'Таблицы',
    text: 'Ячейки и формулы в браузере. Откройте Excel или CSV и продолжите работу.',
  },
  {
    title: 'Документы',
    text: 'Текст рядом с таблицей. Файл Word можно открыть и скачать обратно.',
  },
  {
    title: 'Свои файлы',
    text: 'У каждого аккаунта своя папка на сервере. Документы не уходят в чужое облако.',
  },
  {
    title: 'AI по желанию',
    text: 'Чат, план и агент: DeepSeek, Ollama или LM Studio. Подключение можно пропустить.',
  },
];

const steps = [
  {
    title: 'Войти через Яндекс',
    text: 'Один аккаунт — и сразу своя папка с файлами.',
  },
  {
    title: 'Создать или открыть',
    text: 'Новая таблица, новый документ или файл с компьютера.',
  },
  {
    title: 'Править',
    text: 'Меняйте сами или попросите AI разобрать числа и текст.',
  },
];

onMounted(() => {
  const err = route.query.authError;
  if (typeof err === 'string' && err) {
    authError.value = err;
    void router.replace({ query: { ...route.query, authError: undefined } });
  }
});

function loginYandex() {
  window.location.href = '/auth/yandex';
}
</script>

<template>
  <div class="landing">
    <header class="bar">
      <a class="brand" href="#top">
        <AppLogo :size="28" />
      </a>
      <nav class="nav" aria-label="Разделы">
        <a href="#features">Возможности</a>
      </nav>
      <div class="bar-actions">
        <button
          v-if="yandexEnabled"
          class="ui-btn ui-btn--primary"
          type="button"
          @click="loginYandex"
        >
          Войти через Яндекс
        </button>
      </div>
    </header>

    <main id="top">
      <section class="hero">
        <div class="hero-copy">
          <h1>Таблицы и документы в браузере</h1>
          <p class="lead">
            Свои файлы на вашем сервере. AI-помощник подключается, только если
            он нужен.
          </p>
          <p v-if="authError" class="error" role="alert">
            {{ authError }}
          </p>
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
        </div>

        <div class="stage" aria-hidden="true">
          <div class="window">
            <div class="window-bar">
              <span class="tab tab--on">Бюджет</span>
              <span class="tab">Отчёт</span>
            </div>
            <div class="window-body">
              <div class="workspace">
                <div class="sheet">
                  <div class="sheet-row sheet-row--head">
                    <span></span>
                    <span>A</span>
                    <span>B</span>
                    <span>C</span>
                  </div>
                  <div class="sheet-row">
                    <span class="row-n">1</span>
                    <span>Статья</span>
                    <span>План</span>
                    <span>Факт</span>
                  </div>
                  <div class="sheet-row">
                    <span class="row-n">2</span>
                    <span>Аренда</span>
                    <span>120</span>
                    <span class="cell--on">118</span>
                  </div>
                  <div class="sheet-row">
                    <span class="row-n">3</span>
                    <span>Зарплата</span>
                    <span>400</span>
                    <span>400</span>
                  </div>
                  <div class="sheet-row">
                    <span class="row-n">4</span>
                    <span>Итого</span>
                    <span>520</span>
                    <span>518</span>
                  </div>
                </div>
                <div class="doc">
                  <p class="doc-title">Квартальный отчёт</p>
                  <p>Расходы держатся в плане. Аренда чуть ниже бюджета.</p>
                </div>
              </div>
              <aside class="chat">
                <div class="chat-tabs">
                  <span class="chat-tab chat-tab--on">Чат</span>
                  <span class="chat-tab">План</span>
                  <span class="chat-tab">Агент</span>
                </div>
                <p class="bubble">Посчитай итог по факту</p>
                <p class="bubble bubble--ai">518 — на 2 меньше плана.</p>
              </aside>
            </div>
          </div>
        </div>
      </section>

      <section id="features" class="section">
        <h2>Возможности</h2>
        <ul class="cards">
          <li v-for="item in features" :key="item.title" class="card">
            <h3>{{ item.title }}</h3>
            <p>{{ item.text }}</p>
          </li>
        </ul>
      </section>

      <section class="section">
        <h2>Как начать</h2>
        <ol class="steps">
          <li v-for="(step, index) in steps" :key="step.title">
            <span class="num">{{ index + 1 }}</span>
            <h3>{{ step.title }}</h3>
            <p>{{ step.text }}</p>
          </li>
        </ol>
      </section>
    </main>
  </div>
</template>

<style scoped lang="scss">
.landing {
  height: 100%;
  overflow: auto;
  scroll-behavior: smooth;
  background: var(--bg-0);
  color: var(--fg-1);
}

.bar {
  position: sticky;
  top: 0;
  z-index: 2;
  display: grid;
  grid-template-columns: 1fr auto 1fr;
  align-items: center;
  gap: 16px;
  height: var(--app-header-height);
  padding: 0 var(--app-header-padding-x);
  border-bottom: 1px solid var(--border);
  background: var(--bg-1);
}

.brand {
  justify-self: start;
  display: inline-flex;
  color: inherit;
  text-decoration: none;
}

.nav {
  display: flex;
  justify-content: center;
}

.nav a {
  color: var(--fg-2);
  text-decoration: none;
  font-weight: 600;
}

.nav a:hover {
  color: var(--fg-0);
}

.nav a:focus-visible {
  outline: none;
  border-radius: var(--radius-sm);
  box-shadow: var(--focus-ring);
}

.bar-actions {
  justify-self: end;
}

.hero,
.section {
  width: min(1120px, calc(100% - 40px));
  margin-inline: auto;
}

.hero {
  display: grid;
  grid-template-columns: minmax(0, 1fr) minmax(0, 1.08fr);
  align-items: center;
  gap: 48px;
  padding: 64px 0 32px;
}

.hero-copy h1,
.section h2 {
  margin: 0;
  font-family: var(--font-display);
  font-weight: 700;
  letter-spacing: -0.03em;
  color: var(--fg-0);
}

.hero-copy h1 {
  font-size: clamp(2.25rem, 4.5vw, 3.5rem);
  line-height: 1.05;
}

.lead,
.card p,
.steps p,
.doc p,
.bubble {
  margin: 0;
  color: var(--fg-2);
  line-height: 1.5;
}

.lead {
  margin-top: 16px;
  max-width: 36ch;
  font-size: 1.15rem;
}

.error {
  margin: 16px 0 0;
  max-width: 42ch;
  color: var(--danger);
  line-height: 1.4;
}

.actions {
  margin-top: 28px;
}

.stage {
  min-width: 0;
}

.window {
  border: 1px solid var(--border);
  border-radius: var(--radius-lg);
  background: var(--bg-1);
  box-shadow: var(--shadow);
  overflow: hidden;
}

.window-bar {
  display: flex;
  gap: 8px;
  align-items: center;
  height: 40px;
  padding: 0 12px;
  border-bottom: 1px solid var(--border);
  background: var(--bg-2);
}

.tab {
  padding: 4px 10px;
  border-radius: var(--radius-sm);
  color: var(--fg-3);
  font-size: 0.85rem;
  font-weight: 600;
}

.tab--on {
  background: var(--bg-0);
  color: var(--fg-0);
}

.window-body {
  display: grid;
  grid-template-columns: minmax(0, 1.4fr) minmax(140px, 0.7fr);
  min-height: 280px;
}

.workspace {
  display: flex;
  flex-direction: column;
  min-width: 0;
  border-right: 1px solid var(--border);
}

.sheet {
  display: grid;
  font-size: 0.82rem;
  font-variant-numeric: tabular-nums;
}

.sheet-row {
  display: grid;
  grid-template-columns: 28px repeat(3, minmax(0, 1fr));
  border-bottom: 1px solid var(--border);
}

.sheet-row span {
  padding: 8px 10px;
  min-width: 0;
  overflow: hidden;
  text-overflow: ellipsis;
  white-space: nowrap;
  border-right: 1px solid var(--border);
  color: var(--fg-1);
}

.sheet-row span:last-child {
  border-right: 0;
}

.sheet-row--head span,
.row-n {
  color: var(--fg-3);
  background: var(--bg-2);
  font-weight: 600;
}

.cell--on {
  box-shadow: inset 0 0 0 1px var(--accent);
  color: var(--fg-0);
}

.doc {
  padding: 16px;
  display: grid;
  gap: 8px;
}

.doc-title {
  margin: 0;
  font-family: var(--font-display);
  font-weight: 700;
  color: var(--fg-0);
}

.chat {
  display: flex;
  flex-direction: column;
  gap: 10px;
  padding: 12px;
  background: var(--bg-0);
}

.chat-tabs {
  display: flex;
  gap: 6px;
}

.chat-tab {
  padding: 4px 8px;
  border-radius: var(--radius-sm);
  color: var(--fg-3);
  font-size: 0.78rem;
  font-weight: 600;
}

.chat-tab--on {
  background: var(--bg-2);
  color: var(--fg-0);
}

.bubble {
  padding: 8px 10px;
  border: 1px solid var(--border);
  border-radius: var(--radius-md);
  background: var(--bg-1);
  font-size: 0.85rem;
}

.bubble--ai {
  background: var(--bg-2);
  color: var(--fg-1);
}

.section {
  padding: 48px 0 8px;
  scroll-margin-top: calc(var(--app-header-height) + 12px);
}

.section:last-child {
  padding-bottom: 72px;
}

.section h2 {
  font-size: 1.75rem;
  line-height: 1.2;
}

.cards,
.steps {
  display: grid;
  gap: 12px;
  margin: 20px 0 0;
  padding: 0;
  list-style: none;
}

.cards {
  grid-template-columns: repeat(2, minmax(0, 1fr));
}

.card,
.steps li {
  padding: 20px;
  border: 1px solid var(--border);
  border-radius: var(--radius-lg);
  background: var(--bg-1);
}

.card h3,
.steps h3 {
  margin: 0 0 8px;
  font-family: var(--font-display);
  font-size: 1.1rem;
  font-weight: 700;
  letter-spacing: -0.02em;
  color: var(--fg-0);
}

.steps {
  grid-template-columns: repeat(3, minmax(0, 1fr));
}

.num {
  display: inline-flex;
  align-items: center;
  justify-content: center;
  width: 28px;
  height: 28px;
  margin-bottom: 14px;
  border-radius: var(--radius-sm);
  background: var(--bg-2);
  color: var(--accent);
  font-family: var(--font-display);
  font-weight: 700;
}

@media (max-width: 900px) {
  .hero {
    grid-template-columns: 1fr;
    gap: 32px;
    padding-top: 40px;
  }

  .steps {
    grid-template-columns: 1fr;
  }
}

@media (max-width: 720px) {
  .bar {
    grid-template-columns: 1fr auto;
    padding-inline: 16px;
  }

  .nav {
    display: none;
  }

  .hero,
  .section {
    width: min(1120px, calc(100% - 32px));
  }

  .cards {
    grid-template-columns: 1fr;
  }

  .window-body {
    grid-template-columns: 1fr;
  }

  .workspace {
    border-right: 0;
    border-bottom: 1px solid var(--border);
  }
}

@media (prefers-reduced-motion: reduce) {
  .landing {
    scroll-behavior: auto;
  }
}
</style>
