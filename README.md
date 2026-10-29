# Office — Univer + AI

Веб-приложение: таблицы и документы на [Univer](https://univer.ai/), презентации на [@office-kit/pptx-editor](https://github.com/office-kit/pptx), вход через Яндекс ID, персональная папка файлов, AI-ассистент (DeepSeek / Ollama / LM Studio).

## Стек

- Frontend: Vue 3, Vite, TypeScript, SCSS
- Backend: NestJS, Prisma, PostgreSQL
- Auth: Яндекс OAuth

## Быстрый старт

```bash
# из корня репозитория
docker compose up -d postgres
cp apps/api/.env.example apps/api/.env
npm install --prefix apps/api
npm install --prefix apps/web
npm run db:migrate
npm run dev
```

Откройте http://localhost:5176

Отдельно: `npm run dev:api` или `npm run dev:web`.

- **Войти через Яндекс** — задайте `YANDEX_CLIENT_ID` / `YANDEX_CLIENT_SECRET` в `apps/api/.env`

Если API запущен из терминала Cursor и вход через Яндекс падает с `fetch failed` / `ENOTFOUND`, перезапустите API в обычном терминале (`npm run start:dev` в `apps/api`): IDE-прокси и DNS sandbox часто ломают запросы к `oauth.yandex.ru`.

## Возможности MVP

1. Лендинг → вход → онбординг AI (можно пропустить) → стартовая как Google Docs
2. Создать таблицу, документ или презентацию; открыть с компьютера Excel `.xlsx`, Word `.docx`, PowerPoint `.pptx`, а также CSV/JSON; скачать обратно в Excel, Word или PowerPoint
3. Редактор: Univer для таблиц и документов, редактор слайдов для презентаций. Сворачиваемый AI-чат (Чат / План / Агент). На слайдах агент читает и правит фигуры; в чате и плане слайды только читаются
4. Настройки: профиль и провайдеры

Таблицы и документы хранятся как снимок Univer в JSON. Презентация — файл `.pptx`.  
Файлы пользователей: `data/users/{userId}/documents/`  
БД: PostgreSQL, локально `localhost:5432`, база `office`

## Деплой

| Параметр | Значение |
| --- | --- |
| **Хостинг** | VPS Beget, Ubuntu 26.04, 2 ядра / 4 ГБ RAM / 30 ГБ NVMe |
| **Домен** | [office.shamilfrontend.ru](https://office.shamilfrontend.ru) |
| **Контейнеры** | Docker: web (nginx), api (NestJS), PostgreSQL |
| **Прокси / HTTPS** | Caddy (TLS и reverse proxy к web/api) |
| **CI/CD** | GitHub Actions → SSH на VPS |

Сборка и запуск через Docker Compose на VPS. Caddy терминирует HTTPS и проксирует `/` на web, а API-пути (`/auth/yandex`, `/files`, `/users`, `/ai`, `/health`) на api. PostgreSQL хранит данные в томе `postgres_data`, файлы пользователей — в томе `/data`.

`POSTGRES_PASSWORD`, `YANDEX_CLIENT_ID`, `YANDEX_CLIENT_SECRET`, `JWT_SECRET`, `ENCRYPTION_KEY` и `YANDEX_REDIRECT_URI` — GitHub Secrets; CI пишет их в `/opt/office/.env` на VPS. Compose подставляет `DATABASE_URL` с хостом `postgres`.

### Первый деплой

1. DNS: A-запись `office` → IPv4 VPS. Проверка: `dig +short office.shamilfrontend.ru`
2. На VPS один раз (под пользователем из `DEPLOY_USER`, на Beget это `root`):

```bash
sudo apt-get update
sudo apt-get install -y git ca-certificates curl
curl -fsSL https://get.docker.com | sudo sh
sudo usermod -aG docker "$USER"
sudo ufw allow OpenSSH
sudo ufw allow 80/tcp
sudo ufw allow 443/tcp
sudo ufw --force enable
sudo systemctl disable --now nginx apache2 2>/dev/null || true
```

Перелогинься, проверь `docker compose version`. Публичный ключ от `DEPLOY_SSH_KEY` — в `~/.ssh/authorized_keys`. Ключ с ноутбука в GitHub Secrets не класть: для Actions отдельная пара.

3. GitHub Secrets репозитория `shamilfrontend/office`:

| Secret | Назначение |
| --- | --- |
| `DEPLOY_HOST` | Хост VPS |
| `DEPLOY_PORT` | SSH-порт, обычно `22` |
| `DEPLOY_USER` | Пользователь SSH, на Beget `root` |
| `DEPLOY_SSH_KEY` | Приватный ключ для деплоя |
| `YANDEX_CLIENT_ID` | OAuth-приложение Яндекс ID |
| `YANDEX_CLIENT_SECRET` | Секрет OAuth-приложения |
| `JWT_SECRET` | `openssl rand -hex 32` |
| `ENCRYPTION_KEY` | `openssl rand -hex 32`, шифрует ключи AI-провайдеров |
| `YANDEX_REDIRECT_URI` | `https://office.shamilfrontend.ru/auth/yandex/callback` |
| `POSTGRES_PASSWORD` | `openssl rand -hex 24`, пароль пользователя `office` в PostgreSQL |

4. oauth.yandex.ru: добавь тот же Callback URI. Локальный `http://localhost:3000/auth/yandex/callback` не удаляй.
5. Push в `main` → Actions → SSH → `git clone`/`pull` в `/opt/office` → `.env` → `docker compose up -d --build`
6. Проверка: `https://office.shamilfrontend.ru/health` → `{"ok":true}`

Если TLS не выдался — DNS ещё не дошёл или 80/443 закрыты. Репозиторий клонируется по HTTPS; если он приватный, на VPS нужен deploy key с правом чтения.
