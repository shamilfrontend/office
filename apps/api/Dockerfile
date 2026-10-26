FROM node:22-alpine AS build
RUN apk add --no-cache openssl libc6-compat
WORKDIR /repo

COPY packages/shared ./packages/shared
COPY apps/api ./apps/api

WORKDIR /repo/packages/shared
RUN npm ci && npm run build

WORKDIR /repo/apps/api
RUN npm ci && npx prisma generate && npm run build

FROM node:22-alpine
RUN apk add --no-cache openssl libc6-compat
WORKDIR /repo
ENV NODE_ENV=production
COPY --from=build /repo/packages/shared /repo/packages/shared
COPY --from=build /repo/apps/api /repo/apps/api
WORKDIR /repo/apps/api
EXPOSE 3000
CMD ["sh", "-c", "./node_modules/.bin/prisma migrate deploy && node dist/main.js"]
