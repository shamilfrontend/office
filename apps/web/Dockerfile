FROM node:22-alpine AS build
WORKDIR /repo
ENV NODE_OPTIONS=--max-old-space-size=2048

COPY packages/shared ./packages/shared
COPY apps/web ./apps/web

WORKDIR /repo/packages/shared
RUN npm ci && npm run build

WORKDIR /repo/apps/web
RUN npm ci && npm run build

FROM nginx:1.27-alpine
COPY apps/web/nginx.conf /etc/nginx/conf.d/default.conf
COPY --from=build /repo/apps/web/dist /usr/share/nginx/html
EXPOSE 80
