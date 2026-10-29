import { Logger } from '@nestjs/common';
import { NestFactory } from '@nestjs/core';
import { ConfigService } from '@nestjs/config';
import cookieParser from 'cookie-parser';
import express from 'express';
import { AppModule } from './app.module';

// Cursor and some IDE terminals inject HTTP(S)_PROXY to a local sandbox proxy.
// That breaks outbound OAuth (Yandex) with "fetch failed". Opt in with OFFICE_USE_PROXY=true.
if (process.env.OFFICE_USE_PROXY !== 'true') {
  [
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
  ].forEach((key) => {
    delete process.env[key];
  });
}

async function bootstrap() {
  const app = await NestFactory.create(AppModule, { bodyParser: false });
  app.use(express.json({ limit: '32mb' }));
  app.use(express.urlencoded({ extended: true, limit: '32mb' }));
  const config = app.get(ConfigService);
  const frontend =
    config.get<string>('FRONTEND_URL') ?? 'http://localhost:5176';

  app.use(cookieParser());
  app.enableCors({
    origin: frontend,
    credentials: true,
  });

  const port = Number(config.get<string>('PORT') ?? 3000);
  await app.listen(port);
  Logger.log(`API listening on http://localhost:${port}`, 'Bootstrap');
}
bootstrap();
