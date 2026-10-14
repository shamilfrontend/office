# Office — Univer + AI

Веб-приложение: таблицы и документы на [Univer](https://univer.ai/), вход через Яндекс ID, персональная папка файлов, AI-ассистент (DeepSeek / Ollama / LM Studio).

## Стек

- Frontend: Vue 3, Vite, TypeScript, SCSS
- Backend: NestJS, Prisma, SQLite (локально; можно заменить на PostgreSQL)
- Auth: Яндекс OAuth

## Быстрый старт

```bash
# API
cd apps/api
cp .env.example .env   # уже есть .env с SQLite
npm install
npx prisma migrate dev
npm run start:dev

# Web (другой терминал)
cd apps/web
npm install
npm run dev
```

Откройте http://localhost:5176

- **Войти через Яндекс** — задайте `YANDEX_CLIENT_ID` / `YANDEX_CLIENT_SECRET` в `apps/api/.env`

Если API запущен из терминала Cursor и вход через Яндекс падает с `fetch failed` / `ENOTFOUND`, перезапустите API в обычном терминале (`npm run start:dev` в `apps/api`): IDE-прокси и DNS sandbox часто ломают запросы к `oauth.yandex.ru`.

## Возможности MVP

1. Лендинг → вход → онбординг AI (можно пропустить) → стартовая как Google Docs
2. Создать таблицу/документ, открыть с компьютера (Excel `.xlsx` / Word `.docx`, также CSV/JSON), скачать в Excel/Word
3. Редактор Univer + сворачиваемый AI-чат (Чат / План / Агент)
4. Настройки: профиль и провайдеры

Файлы пользователей: `data/users/{userId}/documents/`  
БД: `data/office.db`

## PostgreSQL (опционально)

В `docker-compose.yml` есть Postgres. Смените `DATABASE_URL` и `provider` в `prisma/schema.prisma` на `postgresql`, затем `npx prisma migrate dev`.
