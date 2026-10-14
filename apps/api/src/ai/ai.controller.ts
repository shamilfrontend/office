import {
  Body,
  Controller,
  Delete,
  Get,
  Param,
  Patch,
  Post,
  Query,
  Res,
  UploadedFile,
  UseGuards,
  UseInterceptors,
} from '@nestjs/common';
import { FileInterceptor } from '@nestjs/platform-express';
import type { User } from '@prisma/client';
import type { Response } from 'express';
import { CurrentUser } from '../auth/current-user.decorator';
import { JwtAuthGuard } from '../auth/jwt-auth.guard';
import { AiService } from './ai.service';
import { ProviderKind } from './ai.providers';
import OpenAI from 'openai';

@Controller('ai')
@UseGuards(JwtAuthGuard)
export class AiController {
  constructor(private readonly ai: AiService) {}

  @Get('defaults')
  defaults() {
    return this.ai.defaults();
  }

  @Get('providers')
  list(@CurrentUser() user: User) {
    return this.ai.list(user.id);
  }

  @Post('providers')
  upsert(
    @CurrentUser() user: User,
    @Body()
    body: {
      id?: string;
      provider: ProviderKind;
      baseUrl?: string;
      apiKey?: string;
      model?: string | null;
      enabled?: boolean;
    },
  ) {
    return this.ai.upsert(user, body);
  }

  @Delete('providers/:id')
  remove(@CurrentUser() user: User, @Param('id') id: string) {
    return this.ai.remove(user.id, id);
  }

  @Get('providers/:id/models')
  listModels(@CurrentUser() user: User, @Param('id') id: string) {
    return this.ai.listProviderModels(user.id, id);
  }

  @Post('providers/test')
  test(
    @Body()
    body: { provider: ProviderKind; baseUrl?: string; apiKey?: string },
  ) {
    return this.ai.testConnection(body);
  }

  @Get('chats')
  listChats(
    @CurrentUser() user: User,
    @Query('open') open?: string,
    @Query('documentId') documentId?: string,
  ) {
    if (open === '0') return this.ai.listClosedChats(user.id, documentId);
    return this.ai.listOpenChats(user.id, documentId);
  }

  @Post('chats')
  createChat(
    @CurrentUser() user: User,
    @Body()
    body: {
      mode?: string;
      providerId?: string;
      model?: string;
      documentId?: string;
    },
  ) {
    return this.ai.createChat(user.id, body ?? {});
  }

  @Patch('chats/:id')
  updateChat(
    @CurrentUser() user: User,
    @Param('id') id: string,
    @Body()
    body: {
      title?: string;
      mode?: string;
      providerId?: string | null;
      model?: string | null;
      documentId?: string | null;
      planTitle?: string | null;
      planMarkdown?: string | null;
      pinnedOpen?: boolean;
      sortOrder?: number;
    },
  ) {
    return this.ai.updateChat(user.id, id, body);
  }

  @Post('chats/:id/close')
  closeChat(@CurrentUser() user: User, @Param('id') id: string) {
    return this.ai.closeChat(user.id, id);
  }

  @Delete('chats/:id')
  deleteChat(@CurrentUser() user: User, @Param('id') id: string) {
    return this.ai.deleteChat(user.id, id);
  }

  @Post('chats/:id/messages')
  addMessage(
    @CurrentUser() user: User,
    @Param('id') id: string,
    @Body()
    body: {
      role: string;
      content: string;
      toolName?: string;
      toolCallId?: string;
      toolStatus?: string;
      toolCallsJson?: string;
    },
  ) {
    return this.ai.addMessage(user.id, id, body);
  }

  @Post('attachments/extract')
  @UseInterceptors(
    FileInterceptor('file', { limits: { fileSize: 15 * 1024 * 1024 } }),
  )
  extract(
    @UploadedFile() file: Express.Multer.File,
  ) {
    return this.ai.extractAttachment(file);
  }

  @Post('chat')
  async chat(
    @CurrentUser() user: User,
    @Body()
    body: {
      providerId?: string;
      model?: string;
      mode: 'chat' | 'plan' | 'agent';
      messages: Array<{
        role: 'system' | 'user' | 'assistant' | 'tool';
        content: string;
        tool_call_id?: string;
        name?: string;
      }>;
      tools?: OpenAI.Chat.Completions.ChatCompletionTool[];
      tool_choice?: 'auto' | 'none' | 'required';
    },
    @Res() res: Response,
  ) {
    res.setHeader('Content-Type', 'text/event-stream');
    res.setHeader('Cache-Control', 'no-cache');
    res.setHeader('Connection', 'keep-alive');
    res.flushHeaders?.();

    await this.ai.chatStream(user.id, body, (event) => {
      res.write(`data: ${JSON.stringify(event)}\n\n`);
    });
    res.end();
  }
}
