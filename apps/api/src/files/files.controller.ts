import {
  Body,
  Controller,
  Delete,
  Get,
  Param,
  Patch,
  Post,
  Put,
  Query,
  Res,
  UploadedFile,
  UseGuards,
  UseInterceptors,
} from '@nestjs/common';
import { FileInterceptor } from '@nestjs/platform-express';
import type { User } from '@prisma/client';
import type { Response } from 'express';
import type { DocType } from '@office/shared';
import { CurrentUser } from '../auth/current-user.decorator';
import { JwtAuthGuard } from '../auth/jwt-auth.guard';
import type { MulterFile } from '../types/uploaded-file';
import { FilesService } from './files.service';
import type { DownloadFormat } from './files.service';

@Controller('files')
@UseGuards(JwtAuthGuard)
export class FilesController {
  constructor(private readonly files: FilesService) {}

  @Get()
  list(@CurrentUser() user: User) {
    return this.files.list(user.id);
  }

  @Post('upload')
  @UseInterceptors(
    FileInterceptor('file', { limits: { fileSize: 50 * 1024 * 1024 } }),
  )
  upload(
    @CurrentUser() user: User,
    @UploadedFile() file: MulterFile,
    @Body() body: { type?: string },
  ) {
    return this.files.uploadFromComputer(user, file, body?.type);
  }

  @Post()
  create(
    @CurrentUser() user: User,
    @Body() body: { type: DocType; name?: string },
  ) {
    return this.files.create(user, body.type, body.name);
  }

  @Get(':id/download')
  async download(
    @CurrentUser() user: User,
    @Param('id') id: string,
    @Query('format') format: DownloadFormat | undefined,
    @Res() res: Response,
  ) {
    const allowed: DownloadFormat[] = ['json', 'csv', 'xlsx', 'docx'];
    const fmt =
      format && allowed.includes(format) ? format : ('json' as DownloadFormat);
    const file = await this.files.download(user.id, id, fmt);
    res.setHeader('Content-Type', file.contentType);
    res.setHeader(
      'Content-Disposition',
      `attachment; filename*=UTF-8''${encodeURIComponent(file.filename)}`,
    );
    res.send(file.body);
  }

  @Get(':id')
  get(@CurrentUser() user: User, @Param('id') id: string) {
    return this.files.get(user.id, id);
  }

  @Put(':id')
  save(
    @CurrentUser() user: User,
    @Param('id') id: string,
    @Body() body: { snapshot: unknown; name?: string },
  ) {
    return this.files.save(user.id, id, body.snapshot, body.name);
  }

  @Patch(':id')
  rename(
    @CurrentUser() user: User,
    @Param('id') id: string,
    @Body() body: { name: string },
  ) {
    return this.files.rename(user.id, id, body.name);
  }

  @Delete(':id')
  remove(@CurrentUser() user: User, @Param('id') id: string) {
    return this.files.remove(user.id, id);
  }
}
