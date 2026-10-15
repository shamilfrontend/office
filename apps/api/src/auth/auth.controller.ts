import {
  Controller,
  Get,
  Post,
  Query,
  Res,
  UseGuards,
  BadRequestException,
  ServiceUnavailableException,
} from '@nestjs/common';
import { ConfigService } from '@nestjs/config';
import type { Response } from 'express';
import type { User } from '@prisma/client';
import { AuthService } from './auth.service';
import { CurrentUser } from './current-user.decorator';
import { JwtAuthGuard } from './jwt-auth.guard';

@Controller('auth')
export class AuthController {
  constructor(
    private readonly auth: AuthService,
    private readonly config: ConfigService,
  ) {}

  @Get('yandex')
  startYandex(@Res() res: Response) {
    const clientId = this.config.get<string>('YANDEX_CLIENT_ID');
    if (!clientId) {
      throw new ServiceUnavailableException(
        'Yandex OAuth is not configured. Set YANDEX_CLIENT_ID.',
      );
    }
    return res.redirect(this.auth.getYandexAuthUrl());
  }

  @Get('yandex/callback')
  async yandexCallback(@Query('code') code: string, @Res() res: Response) {
    if (!code) {
      throw new BadRequestException('Missing code');
    }
    const frontend =
      this.config.get<string>('FRONTEND_URL') ?? 'http://localhost:5176';
    try {
      const user = await this.auth.exchangeYandexCode(code);
      this.auth.issueSession(user, res);
      const me = await this.auth.getMePayload(user);
      const needOnboarding = !me.hasAiProvider && !me.onboardingSkipped;
      return res.redirect(
        `${frontend}/auth/callback?onboarding=${needOnboarding ? '1' : '0'}`,
      );
    } catch (e) {
      const message = e instanceof Error ? e.message : 'auth_failed';
      return res.redirect(
        `${frontend}/?authError=${encodeURIComponent(message)}`,
      );
    }
  }

  @Get('me')
  @UseGuards(JwtAuthGuard)
  me(@CurrentUser() user: User) {
    return this.auth.getMePayload(user);
  }

  @Post('logout')
  logout(@Res({ passthrough: true }) res: Response) {
    this.auth.clearSession(res);
    return { ok: true };
  }

  @Get('config')
  authConfig() {
    return {
      yandexEnabled: Boolean(this.config.get<string>('YANDEX_CLIENT_ID')),
    };
  }
}
