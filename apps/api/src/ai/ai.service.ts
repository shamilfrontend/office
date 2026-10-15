import {
  BadRequestException,
  Injectable,
  NotFoundException,
} from '@nestjs/common';
import type { User } from '@prisma/client';
import OpenAI from 'openai';
import { CryptoService } from '../crypto/crypto.service';
import { PrismaService } from '../prisma/prisma.service';
import {
  LLM_ERROR_HINTS,
  PROVIDER_DEFAULTS,
  ProviderKind,
} from './ai.providers';
import { decodeUploadFilename } from '../files/decode-filename';
import { extractAttachmentText } from './extract-text';

@Injectable()
export class AiService {
  constructor(
    private readonly prisma: PrismaService,
    private readonly crypto: CryptoService,
  ) {}

  defaults() {
    return PROVIDER_DEFAULTS;
  }

  list(userId: string) {
    return this.prisma.aiProviderConfig.findMany({
      where: { userId },
      select: {
        id: true,
        provider: true,
        baseUrl: true,
        model: true,
        enabled: true,
        createdAt: true,
        updatedAt: true,
        encryptedApiKey: true,
      },
    }).then((rows) =>
      rows.map(({ encryptedApiKey, ...rest }) => ({
        ...rest,
        hasApiKey: Boolean(encryptedApiKey),
      })),
    );
  }

  async upsert(
    user: User,
    body: {
      id?: string;
      provider: ProviderKind;
      baseUrl?: string;
      apiKey?: string;
      model?: string | null;
      enabled?: boolean;
    },
  ) {
    const defaults = PROVIDER_DEFAULTS[body.provider];
    if (!defaults) throw new BadRequestException('Unknown provider');
    const baseUrl = (body.baseUrl || defaults.baseUrl).replace(/\/$/, '');

    if (body.id) {
      const existing = await this.prisma.aiProviderConfig.findFirst({
        where: { id: body.id, userId: user.id },
      });
      if (!existing) throw new NotFoundException();
      const willClearKey = body.apiKey !== undefined && !body.apiKey;
      const willHaveKey = body.apiKey
        ? true
        : willClearKey
          ? false
          : Boolean(existing.encryptedApiKey);
      if (defaults.needsKey && !willHaveKey) {
        throw new BadRequestException(
          'API-ключ обязателен для этого провайдера.',
        );
      }
      const data = {
        provider: body.provider,
        baseUrl,
        ...(body.model !== undefined ? { model: body.model } : {}),
        enabled: body.enabled ?? true,
        ...(body.apiKey !== undefined
          ? {
              encryptedApiKey: body.apiKey
                ? this.crypto.encrypt(body.apiKey)
                : null,
            }
          : {}),
      };
      return this.prisma.aiProviderConfig.update({
        where: { id: body.id },
        data,
        select: {
          id: true,
          provider: true,
          baseUrl: true,
          model: true,
          enabled: true,
        },
      });
    }

    if (defaults.needsKey && !body.apiKey) {
      throw new BadRequestException(
        'API-ключ обязателен для этого провайдера.',
      );
    }

    return this.prisma.aiProviderConfig.create({
      data: {
        userId: user.id,
        provider: body.provider,
        baseUrl,
        ...(body.model !== undefined ? { model: body.model } : {}),
        enabled: body.enabled ?? true,
        ...(body.apiKey
          ? { encryptedApiKey: this.crypto.encrypt(body.apiKey) }
          : {}),
      },
      select: {
        id: true,
        provider: true,
        baseUrl: true,
        model: true,
        enabled: true,
      },
    });
  }

  async remove(userId: string, id: string) {
    const existing = await this.prisma.aiProviderConfig.findFirst({
      where: { id, userId },
    });
    if (!existing) throw new NotFoundException();
    await this.prisma.aiProviderConfig.delete({ where: { id } });
    return { ok: true };
  }

  private clientFromConfig(cfg: {
    provider: string;
    baseUrl: string;
    encryptedApiKey: string | null;
  }) {
    const defaults = PROVIDER_DEFAULTS[cfg.provider as ProviderKind];
    if (!cfg.encryptedApiKey) {
      if (defaults?.needsKey) {
        throw new BadRequestException(
          'API-ключ не задан. Добавьте ключ в Настройках → Провайдеры.',
        );
      }
      return new OpenAI({
        apiKey: 'ollama',
        baseURL: cfg.baseUrl,
        timeout: 120_000,
      });
    }
    return new OpenAI({
      apiKey: this.crypto.decrypt(cfg.encryptedApiKey),
      baseURL: cfg.baseUrl,
      timeout: 120_000,
    });
  }

  async testConnection(body: {
    provider: ProviderKind;
    baseUrl?: string;
    apiKey?: string;
  }) {
    const defaults = PROVIDER_DEFAULTS[body.provider];
    if (!defaults) {
      return { ok: false, error: 'Unknown provider', models: [] };
    }
    if (defaults.needsKey && !body.apiKey) {
      return {
        ok: false,
        error: 'API-ключ обязателен для этого провайдера.',
        models: [],
      };
    }
    const baseUrl = (body.baseUrl || defaults.baseUrl).replace(/\/$/, '');
    const client = new OpenAI({
      apiKey: body.apiKey || 'ollama',
      baseURL: baseUrl,
      timeout: 30_000,
    });
    try {
      const models = await client.models.list();
      const ids = models.data.map((m) => m.id);
      return { ok: true, models: ids };
    } catch (e) {
      return {
        ok: false,
        error: this.formatError(e),
        models: [],
      };
    }
  }

  async listProviderModels(userId: string, id: string) {
    const cfg = await this.prisma.aiProviderConfig.findFirst({
      where: { id, userId },
    });
    if (!cfg) throw new NotFoundException();
    try {
      const client = this.clientFromConfig(cfg);
      const models = await client.models.list();
      return { models: models.data.map((m) => m.id) };
    } catch (e) {
      if (e instanceof BadRequestException) throw e;
      throw new BadRequestException(this.formatError(e));
    }
  }

  private chatSelect = {
    id: true,
    title: true,
    mode: true,
    providerId: true,
    model: true,
    documentId: true,
    planTitle: true,
    planMarkdown: true,
    pinnedOpen: true,
    sortOrder: true,
    createdAt: true,
    updatedAt: true,
    messages: {
      orderBy: { createdAt: 'asc' as const },
      select: {
        id: true,
        role: true,
        content: true,
        toolName: true,
        toolCallId: true,
        toolStatus: true,
        toolCallsJson: true,
        createdAt: true,
      },
    },
  };

  private chatSummarySelect = {
    id: true,
    title: true,
    mode: true,
    providerId: true,
    model: true,
    documentId: true,
    planTitle: true,
    planMarkdown: true,
    pinnedOpen: true,
    sortOrder: true,
    createdAt: true,
    updatedAt: true,
  };

  async listOpenChats(userId: string, documentId?: string) {
    if (!documentId) return [];
    const rows = await this.prisma.chat.findMany({
      where: { userId, pinnedOpen: true, documentId },
      orderBy: { sortOrder: 'asc' },
      select: this.chatSelect,
    });
    if (rows.length) return rows;
    const created = await this.createChat(userId, { documentId });
    return [created];
  }

  async listClosedChats(userId: string, documentId?: string) {
    if (!documentId) return [];
    return this.prisma.chat.findMany({
      where: { userId, pinnedOpen: false, documentId },
      orderBy: { updatedAt: 'desc' },
      select: this.chatSummarySelect,
    });
  }

  async createChat(
    userId: string,
    body: {
      mode?: string;
      providerId?: string;
      model?: string;
      documentId?: string;
    },
  ) {
    const max = await this.prisma.chat.aggregate({
      where: { userId },
      _max: { sortOrder: true },
    });
    return this.prisma.chat.create({
      data: {
        userId,
        mode: body.mode || 'agent',
        providerId: body.providerId,
        model: body.model,
        documentId: body.documentId,
        pinnedOpen: true,
        sortOrder: (max._max.sortOrder ?? 0) + 1,
      },
      select: this.chatSelect,
    });
  }

  async updateChat(
    userId: string,
    id: string,
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
    await this.ownChat(userId, id);
    return this.prisma.chat.update({
      where: { id },
      data: body,
      select: this.chatSelect,
    });
  }

  async closeChat(userId: string, id: string) {
    const existing = await this.ownChat(userId, id);
    if (!existing.pinnedOpen) {
      return { ok: true as const, chat: null };
    }
    await this.prisma.chat.update({
      where: { id },
      data: { pinnedOpen: false },
    });
    const documentId = existing.documentId ?? undefined;
    if (!documentId) return { ok: true as const, chat: null };
    const left = await this.prisma.chat.count({
      where: { userId, pinnedOpen: true, documentId },
    });
    if (!left) {
      const created = await this.createChat(userId, { documentId });
      return { ok: true as const, chat: created };
    }
    return { ok: true as const, chat: null };
  }

  async deleteChat(userId: string, id: string) {
    const existing = await this.ownChat(userId, id);
    const documentId = existing.documentId ?? undefined;
    await this.prisma.chat.delete({ where: { id } });
    if (!documentId) return { ok: true, chat: null };
    const left = await this.prisma.chat.count({
      where: { userId, pinnedOpen: true, documentId },
    });
    if (!left) {
      const created = await this.createChat(userId, { documentId });
      return { ok: true, chat: created };
    }
    return { ok: true, chat: null };
  }

  async addMessage(
    userId: string,
    chatId: string,
    body: {
      role: string;
      content: string;
      toolName?: string;
      toolCallId?: string;
      toolStatus?: string;
      toolCallsJson?: string;
    },
  ) {
    await this.ownChat(userId, chatId);
    return this.prisma.message.create({
      data: {
        chatId,
        role: body.role,
        content: body.content ?? '',
        toolName: body.toolName,
        toolCallId: body.toolCallId,
        toolStatus: body.toolStatus,
        toolCallsJson: body.toolCallsJson,
      },
      select: {
        id: true,
        role: true,
        content: true,
        toolName: true,
        toolCallId: true,
        toolStatus: true,
        toolCallsJson: true,
        createdAt: true,
      },
    });
  }

  async extractAttachment(file: Express.Multer.File) {
    if (!file) throw new BadRequestException('Файл не передан');
    const name = decodeUploadFilename(file.originalname || 'file');
    const text = await extractAttachmentText(name, file.buffer);
    return { name, text };
  }

  private async ownChat(userId: string, id: string) {
    const chat = await this.prisma.chat.findFirst({ where: { id, userId } });
    if (!chat) throw new NotFoundException();
    return chat;
  }

  async chatStream(
    userId: string,
    body: {
      providerId?: string;
      model?: string;
      mode: 'chat' | 'plan' | 'agent';
      messages: OpenAI.Chat.Completions.ChatCompletionMessageParam[];
      tools: OpenAI.Chat.Completions.ChatCompletionTool[];
      tool_choice: 'auto' | 'none' | 'required';
    },
    onEvent: (event: Record<string, unknown>) => void,
    signal?: AbortSignal,
  ) {
    if (signal?.aborted) return;
    let cfg;
    try {
      cfg = await this.resolveProvider(userId, body.providerId);
    } catch (e) {
      onEvent({
        type: 'error',
        error:
          e instanceof BadRequestException
            ? String(e.message)
            : 'AI provider is not configured',
      });
      return;
    }
    const model = body.model || cfg.model;
    if (!model) {
      onEvent({
        type: 'error',
        error: 'Выберите модель в чате ассистента.',
      });
      return;
    }
    let client: OpenAI;
    try {
      client = this.clientFromConfig(cfg);
    } catch (e) {
      onEvent({
        type: 'error',
        error:
          e instanceof BadRequestException
            ? String(e.message)
            : this.formatError(e),
      });
      return;
    }
    const systemExtra =
      body.mode === 'plan'
        ? 'Ты помощник в офисном редакторе. Всегда отвечай пользователю на русском языке. Режим План: только читай документ через tools и создавай план (create_plan/update_plan). Не изменяй документ. Если пользователь говорит «здесь/это/выделенное», используй get_selection или контекст выделения. Для таблиц сначала read_range/read_formulas/get_selection; другой лист — activate_sheet. В плане указывай конкретные диапазоны/ячейки. Диаграммы (charts) tools не поддерживают — отмечай как ручной шаг.'
        : body.mode === 'agent'
          ? 'Ты помощник в офисном редакторе. Всегда отвечай пользователю на русском языке. Режим Агент: можешь читать и изменять документ через tools. Если пользователь говорит «здесь/это/выделенное», опирайся на selection / get_selection. Для таблиц: другие листы — сначала activate_sheet; insert_sheet/rename_sheet/delete_sheet/set_freeze для структуры; set_cells — литералы; set_formulas — формулы (EN-имена, аргументы через запятую, ведущий =, напр. =SUM(A1:A10); пустые ячейки в матрице можно оставлять — они пропускаются); copy_formulas — размножить формулы вниз с автосдвигом ссылок (предпочтительнее для 10+ одинаковых строк); read_formulas читает формулы и значения; insert_image — картинка по https URL; диаграммы (charts) недоступны как tool — скажи пользователю вставить вручную; для оформления — format_range, set_column_widths (ширина в пикселях, типично 80–220), set_number_format, set_borders, set_row_heights; для объединения — merge_range; sort_range — сортировка; create_filter/set_column_filter/remove_filter/get_filter — автофильтр; search_in_sheet / replace_in_sheet — поиск и замена литералов. Сначала read_range/read_formulas/get_selection. При выполнении плана вызывай update_plan по мере прогресса (status: completed). Для документов: get_doc_outline/search_in_document; структура — insert_paragraph, set_heading, set_list; format_text по выделению или paragraphIndex. После серии tools дай краткий итог.'
          : 'Ты помощник в офисном редакторе. Всегда отвечай пользователю на русском языке. Режим Чат: читай открытый документ через tools (сначала get_workbook_outline/get_doc_outline или get_selection, затем read_range / read_formulas / read_document / search_in_document; для другого листа — activate_sheet). Не изменяй документ. Если пользователь говорит «здесь/это/выделенное», используй get_selection. Если tools доступны, не говори, что нет доступа к файлу — сначала прочитай его. Предлагай правки текстом.';

    let attempt = 0;
    while (attempt < 3) {
      if (signal?.aborted) return;
      attempt++;
      try {
        const stream = await client.chat.completions.create(
          {
            model,
            stream: true,
            stream_options: { include_usage: true },
            messages: [
              { role: 'system', content: systemExtra },
              ...body.messages,
            ],
            tools: body.tools.length ? body.tools : undefined,
            tool_choice: body.tools.length ? body.tool_choice : undefined,
          },
          { signal },
        );

        let full = '';
        const toolCalls: Array<{
          id: string;
          name: string;
          arguments: string;
        }> = [];
        let usage: {
          promptTokens: number;
          completionTokens: number;
          totalTokens: number;
        } | null = null;

        for await (const chunk of stream) {
          if (signal?.aborted) return;
          const chunkUsage = (
            chunk as {
              usage?: {
                prompt_tokens?: number;
                completion_tokens?: number;
                total_tokens?: number;
              } | null;
            }
          ).usage;
          if (chunkUsage) {
            usage = {
              promptTokens: chunkUsage.prompt_tokens ?? 0,
              completionTokens: chunkUsage.completion_tokens ?? 0,
              totalTokens: chunkUsage.total_tokens ?? 0,
            };
          }
          const delta = chunk.choices[0]?.delta;
          if (delta?.content) {
            full += delta.content;
            onEvent({ type: 'delta', content: delta.content });
          }
          if (delta?.tool_calls) {
            for (const tc of delta.tool_calls) {
              const idx = tc.index ?? 0;
              if (!toolCalls[idx]) {
                toolCalls[idx] = {
                  id: tc.id ?? `call_${idx}`,
                  name: tc.function?.name ?? '',
                  arguments: tc.function?.arguments ?? '',
                };
              } else {
                if (tc.id) toolCalls[idx].id = tc.id;
                if (tc.function?.name) toolCalls[idx].name += tc.function.name;
                if (tc.function?.arguments)
                  toolCalls[idx].arguments += tc.function.arguments;
              }
            }
          }
        }

        if (signal?.aborted) return;
        if (usage) {
          onEvent({
            type: 'usage',
            promptTokens: usage.promptTokens,
            completionTokens: usage.completionTokens,
            totalTokens: usage.totalTokens,
          });
        }
        onEvent({
          type: 'done',
          content: full,
          toolCalls: toolCalls.filter(Boolean),
        });
        return;
      } catch (e) {
        if (signal?.aborted) return;
        const name = (e as { name?: string }).name;
        if (name === 'AbortError' || name === 'APIUserAbortError') return;
        const status = (e as { status?: number }).status;
        if ((status === 408 || status === 429) && attempt < 3 && !signal?.aborted) {
          await new Promise((r) => setTimeout(r, 500 * attempt));
          if (signal?.aborted) return;
          continue;
        }
        onEvent({ type: 'error', error: this.formatError(e) });
        return;
      }
    }
  }

  private async resolveProvider(userId: string, providerId?: string) {
    const cfg = providerId
      ? await this.prisma.aiProviderConfig.findFirst({
          where: { id: providerId, userId, enabled: true },
        })
      : await this.prisma.aiProviderConfig.findFirst({
          where: { userId, enabled: true },
          orderBy: { updatedAt: 'desc' },
        });
    if (!cfg) {
      throw new BadRequestException('AI provider is not configured');
    }
    return cfg;
  }

  private formatError(e: unknown) {
    const err = e as {
      status?: number;
      message?: string;
      error?: { message?: string } | string;
    };
    const providerMsg =
      typeof err.error === 'object' && err.error?.message
        ? err.error.message
        : typeof err.error === 'string'
          ? err.error
          : err.message;
    if (err.status && LLM_ERROR_HINTS[err.status]) {
      const hint = LLM_ERROR_HINTS[err.status];
      if (providerMsg && !providerMsg.includes(hint)) {
        return `${hint} ${providerMsg}`;
      }
      return hint;
    }
    if (err.message?.includes('ECONNREFUSED')) {
      return 'Сеть недоступна. Проверьте адрес endpoint и что локальный сервер запущен.';
    }
    const protocolText = `${providerMsg || ''} ${err.message || ''}`;
    if (
      protocolText.includes("role 'tool'") ||
      protocolText.includes('tool_calls')
    ) {
      return 'Сбой истории tools. Начните новый чат или повторите запрос.';
    }
    return providerMsg || 'Ошибка запроса к провайдеру';
  }
}
