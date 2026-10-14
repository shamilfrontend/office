import { Injectable } from '@nestjs/common';
import { ConfigService } from '@nestjs/config';
import { JwtService } from '@nestjs/jwt';
import type { User } from '@prisma/client';
import type { Response } from 'express';
import { PrismaService } from '../prisma/prisma.service';
import { UserStorageService } from '../storage/user-storage.service';

type YandexProfile = {
  id: string;
  login?: string;
  display_name?: string;
  real_name?: string;
  default_email?: string;
  default_avatar_id?: string;
  is_avatar_empty?: boolean;
};

@Injectable()
export class AuthService {
  constructor(
    private readonly prisma: PrismaService,
    private readonly storage: UserStorageService,
    private readonly jwt: JwtService,
    private readonly config: ConfigService,
  ) {}

  getYandexAuthUrl(state?: string) {
    const clientId = this.config.get<string>('YANDEX_CLIENT_ID');
    const redirectUri = this.config.get<string>('YANDEX_REDIRECT_URI');
    const params = new URLSearchParams({
      response_type: 'code',
      client_id: clientId ?? '',
      redirect_uri: redirectUri ?? '',
      force_confirm: 'yes',
    });
    if (state) params.set('state', state);
    return `https://oauth.yandex.ru/authorize?${params.toString()}`;
  }

  async exchangeYandexCode(code: string) {
    const clientId = this.config.get<string>('YANDEX_CLIENT_ID') ?? '';
    const clientSecret = this.config.get<string>('YANDEX_CLIENT_SECRET') ?? '';
    const redirectUri = this.config.get<string>('YANDEX_REDIRECT_URI') ?? '';

    let tokenRes: globalThis.Response;
    try {
      tokenRes = await fetch('https://oauth.yandex.ru/token', {
        method: 'POST',
        headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
        body: new URLSearchParams({
          grant_type: 'authorization_code',
          code,
          client_id: clientId,
          client_secret: clientSecret,
          redirect_uri: redirectUri,
        }),
      });
    } catch (e) {
      const cause = e instanceof Error && e.cause instanceof Error ? e.cause.message : '';
      throw new Error(
        `Cannot reach oauth.yandex.ru${cause ? `: ${cause}` : ''}. Check network/proxy.`,
      );
    }
    if (!tokenRes.ok) {
      const text = await tokenRes.text();
      throw new Error(`Yandex token error: ${text}`);
    }
    const tokenJson = (await tokenRes.json()) as { access_token: string };

    let profileRes: globalThis.Response;
    try {
      profileRes = await fetch('https://login.yandex.ru/info?format=json', {
        headers: { Authorization: `OAuth ${tokenJson.access_token}` },
      });
    } catch (e) {
      const cause = e instanceof Error && e.cause instanceof Error ? e.cause.message : '';
      throw new Error(
        `Cannot reach login.yandex.ru${cause ? `: ${cause}` : ''}. Check network/proxy.`,
      );
    }
    if (!profileRes.ok) {
      throw new Error('Failed to fetch Yandex profile');
    }
    const profile = (await profileRes.json()) as YandexProfile;
    return this.upsertFromYandex(profile);
  }

  async upsertFromYandex(profile: YandexProfile) {
    const yandexId = String(profile.id);
    const displayName =
      profile.display_name || profile.real_name || profile.login || 'User';
    const avatarUrl =
      !profile.is_avatar_empty && profile.default_avatar_id
        ? `https://avatars.yandex.net/get-yapic/${profile.default_avatar_id}/islands-200`
        : null;

    let user = await this.prisma.user.findUnique({ where: { yandexId } });
    if (!user) {
      user = await this.prisma.user.create({
        data: {
          yandexId,
          email: profile.default_email ?? null,
          displayName,
          avatarUrl,
          storagePath: '',
        },
      });
      const storagePath = await this.storage.ensureUserFolder(user.id);
      user = await this.prisma.user.update({
        where: { id: user.id },
        data: { storagePath },
      });
    } else {
      await this.storage.ensureUserFolder(user.id);
      user = await this.prisma.user.update({
        where: { id: user.id },
        data: {
          email: profile.default_email ?? user.email,
          displayName,
          avatarUrl: avatarUrl ?? user.avatarUrl,
        },
      });
    }
    return user;
  }

  issueSession(user: User, res: Response) {
    const token = this.jwt.sign({
      sub: user.id,
      yandexId: user.yandexId,
    });
    const cookieName = this.config.get<string>('COOKIE_NAME') ?? 'office_session';
    res.cookie(cookieName, token, {
      httpOnly: true,
      sameSite: 'lax',
      secure: false,
      path: '/',
      maxAge: 7 * 24 * 60 * 60 * 1000,
    });
    return token;
  }

  clearSession(res: Response) {
    const cookieName = this.config.get<string>('COOKIE_NAME') ?? 'office_session';
    res.clearCookie(cookieName);
  }

  async getMePayload(user: User) {
    const providers = await this.prisma.aiProviderConfig.findMany({
      where: { userId: user.id, enabled: true },
      select: {
        id: true,
        provider: true,
        baseUrl: true,
        model: true,
        enabled: true,
      },
    });
    return {
      id: user.id,
      email: user.email,
      displayName: user.displayName,
      avatarUrl: user.avatarUrl,
      onboardingSkipped: user.onboardingSkipped,
      hasAiProvider: providers.length > 0,
      providers,
    };
  }
}
