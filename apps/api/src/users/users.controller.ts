import { Body, Controller, Patch, UseGuards } from '@nestjs/common';
import type { User } from '@prisma/client';
import { CurrentUser } from '../auth/current-user.decorator';
import { JwtAuthGuard } from '../auth/jwt-auth.guard';
import { PrismaService } from '../prisma/prisma.service';
import { AuthService } from '../auth/auth.service';

@Controller('users')
@UseGuards(JwtAuthGuard)
export class UsersController {
  constructor(
    private readonly prisma: PrismaService,
    private readonly auth: AuthService,
  ) {}

  @Patch('onboarding')
  async skipOnboarding(
    @CurrentUser() user: User,
    @Body() body: { skipped?: boolean },
  ) {
    const updated = await this.prisma.user.update({
      where: { id: user.id },
      data: { onboardingSkipped: body.skipped ?? true },
    });
    return this.auth.getMePayload(updated);
  }
}
