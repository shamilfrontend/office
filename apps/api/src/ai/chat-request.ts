import { BadRequestException } from '@nestjs/common';
import type OpenAI from 'openai';
import {
  type ChatMode,
  type DocType,
  ALL_TOOL_NAMES,
  toolsForMode,
} from './office-tools.catalog';

const MAX_MESSAGES = 80;
const MAX_CONTENT_CHARS = 32_000;
const MAX_TOTAL_CHARS = 200_000;
const MAX_ARGS_CHARS = 8_000;

const MODES = new Set<ChatMode>(['chat', 'plan', 'agent']);
const DOC_TYPES = new Set<DocType>(['SHEET', 'DOC']);
const ROLES = new Set(['system', 'user', 'assistant', 'tool']);
const TOOL_CHOICES = new Set(['auto', 'none', 'required']);

export type SanitizedChatRequest = {
  providerId?: string;
  model?: string;
  mode: ChatMode;
  docType: DocType;
  hasDocument: boolean;
  buildingPlan: boolean;
  messages: OpenAI.Chat.Completions.ChatCompletionMessageParam[];
  tools: OpenAI.Chat.Completions.ChatCompletionTool[];
  tool_choice: 'auto' | 'none' | 'required';
};

function asString(value: unknown, field: string): string {
  if (typeof value !== 'string') {
    throw new BadRequestException(`Поле ${field} должно быть строкой`);
  }
  return value;
}

function softClipContent(content: string, max: number): string {
  if (content.length <= max) return content;
  return JSON.stringify({
    truncated: true,
    preview: content.slice(0, max),
    originalChars: content.length,
  });
}

function messageContentChars(
  msg: OpenAI.Chat.Completions.ChatCompletionMessageParam,
): number {
  let n = 0;
  if (typeof msg.content === 'string') n += msg.content.length;
  if (msg.role === 'assistant' && 'tool_calls' in msg && msg.tool_calls) {
    msg.tool_calls.forEach((tc) => {
      if (
        tc.type === 'function' &&
        typeof tc.function?.arguments === 'string'
      ) {
        n += tc.function.arguments.length;
      }
    });
  }
  return n;
}

function softTrimMessages(
  messages: OpenAI.Chat.Completions.ChatCompletionMessageParam[],
): OpenAI.Chat.Completions.ChatCompletionMessageParam[] {
  const totalChars = (list: typeof messages) =>
    list.reduce((sum, m) => sum + messageContentChars(m), 0);

  if (
    messages.length <= MAX_MESSAGES &&
    totalChars(messages) <= MAX_TOTAL_CHARS
  ) {
    return messages;
  }

  const turns: OpenAI.Chat.Completions.ChatCompletionMessageParam[][] = [];
  let current: OpenAI.Chat.Completions.ChatCompletionMessageParam[] = [];
  messages.forEach((msg) => {
    if (msg.role === 'user' && current.length) {
      turns.push(current);
      current = [];
    }
    current.push(msg);
  });
  if (current.length) turns.push(current);

  let dropped = 0;
  while (
    turns.length > 1 &&
    (turns.reduce((n, t) => n + t.length, 0) > MAX_MESSAGES ||
      turns.reduce((n, t) => n + totalChars(t), 0) > MAX_TOTAL_CHARS)
  ) {
    dropped += turns.shift()!.length;
  }

  const kept = turns.flat();
  if (!dropped) return kept;
  return [
    {
      role: 'user',
      content: `[Пропущены ранние сообщения: ${dropped}]`,
    },
    ...kept,
  ];
}

function sanitizeToolCalls(
  raw: unknown,
): OpenAI.Chat.Completions.ChatCompletionMessageToolCall[] | undefined {
  if (!Array.isArray(raw) || raw.length === 0) return undefined;
  const out = raw.reduce<
    OpenAI.Chat.Completions.ChatCompletionMessageToolCall[]
  >((acc, item) => {
    if (!item || typeof item !== 'object') return acc;
    const tc = item as Record<string, unknown>;
    const fn = tc.function;
    if (!fn || typeof fn !== 'object') return acc;
    const f = fn as Record<string, unknown>;
    const name = typeof f.name === 'string' ? f.name : '';
    if (!name || !ALL_TOOL_NAMES.has(name)) return acc;
    const id =
      typeof tc.id === 'string' && tc.id.trim() ? tc.id : `call_${acc.length}`;
    let args =
      typeof f.arguments === 'string'
        ? f.arguments
        : JSON.stringify(f.arguments ?? {});
    args = softClipContent(args, MAX_ARGS_CHARS);
    acc.push({
      id,
      type: 'function',
      function: { name, arguments: args },
    });
    return acc;
  }, []);
  return out.length ? out : undefined;
}

function sanitizeMessage(
  raw: unknown,
  index: number,
): OpenAI.Chat.Completions.ChatCompletionMessageParam {
  if (!raw || typeof raw !== 'object') {
    throw new BadRequestException(`Сообщение #${index + 1} некорректно`);
  }
  const msg = raw as Record<string, unknown>;
  const role = asString(msg.role, `messages[${index}].role`);
  if (!ROLES.has(role)) {
    throw new BadRequestException(`Недопустимая роль сообщения: ${role}`);
  }

  let content = '';
  if (msg.content === null || msg.content === undefined) {
    content = '';
  } else if (typeof msg.content === 'string') {
    content = msg.content;
  } else {
    throw new BadRequestException(
      `messages[${index}].content должно быть строкой`,
    );
  }
  content = softClipContent(content, MAX_CONTENT_CHARS);

  if (role === 'tool') {
    const toolCallId =
      typeof msg.tool_call_id === 'string' ? msg.tool_call_id.trim() : '';
    if (!toolCallId) {
      throw new BadRequestException(
        `Сообщение tool #${index + 1} без tool_call_id`,
      );
    }
    const name = typeof msg.name === 'string' ? msg.name : undefined;
    return {
      role: 'tool',
      content,
      tool_call_id: toolCallId,
      ...(name ? { name } : {}),
    };
  }

  if (role === 'assistant') {
    const toolCalls = sanitizeToolCalls(msg.tool_calls);
    return {
      role: 'assistant',
      content,
      ...(toolCalls ? { tool_calls: toolCalls } : {}),
    };
  }

  if (role === 'system') {
    return { role: 'system', content };
  }

  return { role: 'user', content };
}

export function sanitizeChatRequest(body: unknown): SanitizedChatRequest {
  if (!body || typeof body !== 'object') {
    throw new BadRequestException('Тело запроса некорректно');
  }
  const b = body as Record<string, unknown>;

  const modeRaw = b.mode;
  if (typeof modeRaw !== 'string' || !MODES.has(modeRaw as ChatMode)) {
    throw new BadRequestException('mode должен быть chat, plan или agent');
  }
  const mode = modeRaw as ChatMode;

  const docType: DocType =
    typeof b.docType === 'string' && DOC_TYPES.has(b.docType as DocType)
      ? (b.docType as DocType)
      : 'SHEET';

  const hasDocument = b.hasDocument !== false && b.hasDocument !== 0;
  const buildingPlan = Boolean(b.buildingPlan);

  if (!Array.isArray(b.messages)) {
    throw new BadRequestException('messages должен быть массивом');
  }

  const messages = softTrimMessages(
    b.messages.map((m, i) => sanitizeMessage(m, i)),
  );

  const tools = toolsForMode(mode, buildingPlan, hasDocument, docType);

  let toolChoice: 'auto' | 'none' | 'required' = 'auto';
  if (typeof b.tool_choice === 'string' && TOOL_CHOICES.has(b.tool_choice)) {
    toolChoice = b.tool_choice as 'auto' | 'none' | 'required';
  }
  if (!tools.length) {
    toolChoice = 'none';
  }

  const providerId =
    typeof b.providerId === 'string' && b.providerId.trim()
      ? b.providerId.trim()
      : undefined;
  const model =
    typeof b.model === 'string' && b.model.trim() ? b.model.trim() : undefined;

  return {
    providerId,
    model,
    mode,
    docType,
    hasDocument,
    buildingPlan,
    messages,
    tools,
    tool_choice: toolChoice,
  };
}
