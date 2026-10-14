import { NestFactory } from '@nestjs/core';
import { ConfigService } from '@nestjs/config';
import { AppModule } from './app.module';

// eslint-disable-next-line @typescript-eslint/no-require-imports
const cookieParser = require('cookie-parser') as () => import('express').RequestHandler;
// eslint-disable-next-line @typescript-eslint/no-require-imports
const express = require('express') as typeof import('express');

// Cursor and some IDE terminals inject HTTP(S)_PROXY to a local sandbox proxy.
// That breaks outbound OAuth (Yandex) with "fetch failed". Opt in with OFFICE_USE_PROXY=true.
if (process.env.OFFICE_USE_PROXY !== 'true') {
  for (const key of [
    'HTTP_PROXY',
    'HTTPS_PROXY',
    'ALL_PROXY',
    'SOCKS_PROXY',
    'SOCKS5_PROXY',
    'http_proxy',
    'https_proxy',
    'all_proxy',
    'socks_proxy',
    'socks5_proxy',
  ]) {
    delete process.env[key];
  }
}

async function bootstrap() {
  const app = await NestFactory.create(AppModule, { bodyParser: false });
  app.use(express.json({ limit: '4mb' }));
  app.use(express.urlencoded({ extended: true, limit: '4mb' }));
  const config = app.get(ConfigService);
  const frontend = config.get<string>('FRONTEND_URL') ?? 'http://localhost:5176';

  app.use(cookieParser());
  app.enableCors({
    origin: frontend,
    credentials: true,
  });

  const port = Number(config.get<string>('PORT') ?? 3000);
  await app.listen(port);
  console.log(`API listening on http://localhost:${port}`);
}
bootstrap();
