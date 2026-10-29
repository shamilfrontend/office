import {
  slideToolNamesForMode,
  systemPromptForMode,
  type DocType,
  type OfficeToolDef,
} from '@office/shared';
import { toolsForMode, type ChatMode } from '@/univer/officeTools';
import {
  MESSAGE_FRAMING_TOKENS,
  chatTokenCalibrator,
  estimateMessageTokens,
  estimateTextTokens,
  type TokenEstimateCalibrator,
} from './estimateTokens';
import { effectiveReservedTokens, getModelContextInfo } from './modelContext';

type ContextUsageCategoryId = 'system' | 'tools' | 'conversation';

type ContextUsageCategory = {
  id: ContextUsageCategoryId;
  label: string;
  tokens: number;
  color: string;
};

export type ChatContextUsage = {
  contextWindow: number;
  usedTokens: number;
  reservedTokens: number;
  percentFull: number;
  categories: ContextUsageCategory[];
};

export const RESERVE_COLOR = '#38BDF8';

const CONTEXT_USAGE_META: Record<
  ContextUsageCategoryId,
  { label: string; color: string }
> = {
  system: { label: 'Системный промпт', color: '#9CA3AF' },
  tools: { label: 'Определения инструментов', color: '#C4B5FD' },
  conversation: { label: 'Диалог', color: '#FB923C' },
};

function wireMessageTokens(msg: Record<string, unknown>): number {
  let tokens = 0;
  if (typeof msg.content === 'string') {
    tokens += estimateMessageTokens(msg.content);
  } else if (msg.content == null && msg.role === 'assistant') {
    tokens += MESSAGE_FRAMING_TOKENS;
  }
  const toolCalls = msg.tool_calls;
  if (Array.isArray(toolCalls)) {
    toolCalls.forEach((tc) => {
      const fn = (tc as { function?: { name?: string; arguments?: string } })
        ?.function;
      if (!fn) return;
      tokens += estimateTextTokens(fn.name || '');
      tokens += estimateTextTokens(fn.arguments || '');
      tokens += MESSAGE_FRAMING_TOKENS;
    });
  }
  if (typeof msg.name === 'string') {
    tokens += estimateTextTokens(msg.name);
  }
  return tokens;
}

function buildPendingUserContent(opts: {
  draft: string;
  attachments: Array<{ name: string; text: string }>;
  includeDoc: boolean;
  docName?: string;
  docType?: DocType | null;
  selectionContext?: string;
}): string {
  const display = opts.draft.trim();
  const contextParts: string[] = [];
  if (opts.includeDoc && opts.docName) {
    contextParts.push(`Текущий документ: ${opts.docName} (${opts.docType})`);
    if (opts.selectionContext) contextParts.push(opts.selectionContext);
  }
  opts.attachments.forEach((file) => {
    contextParts.push(`Файл ${file.name}:\n${file.text}`);
  });
  const contextBlock = contextParts.length
    ? `\n\n---\nКонтекст:\n${contextParts.join('\n\n')}`
    : '';
  if (!display && !contextBlock) return '';
  return contextParts.length ? `${display}${contextBlock}` : display;
}

export function estimateChatContextUsage(opts: {
  mode: ChatMode;
  model: string;
  wireMessages: Array<Record<string, unknown>>;
  draft: string;
  attachments: Array<{ name: string; text: string }>;
  includeDoc: boolean;
  docName?: string;
  docType?: DocType | null;
  selectionContext?: string;
  buildingPlan?: boolean;
  hasDocument?: boolean;
  slideTools?: OfficeToolDef[];
  calibrator?: TokenEstimateCalibrator;
}): ChatContextUsage {
  const calibrator = opts.calibrator ?? chatTokenCalibrator;
  const { contextWindow, reservedOutputTokenSpace } = getModelContextInfo(
    opts.model,
  );
  const reservedTokens = effectiveReservedTokens(
    contextWindow,
    reservedOutputTokenSpace,
  );

  const hasDocument = opts.hasDocument !== false;
  const docType: DocType =
    opts.docType === 'DOC' || opts.docType === 'SLIDE' ? opts.docType : 'SHEET';
  const systemTokens = estimateMessageTokens(
    systemPromptForMode(opts.mode, docType),
  );

  const catalogTools = toolsForMode(
    opts.mode,
    Boolean(opts.buildingPlan),
    hasDocument,
    docType,
  );
  const allowedSlideTools = new Set(slideToolNamesForMode(opts.mode));
  const slideTools =
    docType === 'SLIDE'
      ? (opts.slideTools ?? []).filter((tool) =>
          allowedSlideTools.has(tool.function.name),
        )
      : [];
  const tools = [...slideTools, ...catalogTools];
  const toolsTokens = tools.length
    ? estimateTextTokens(JSON.stringify(tools))
    : 0;

  const pending = buildPendingUserContent({
    draft: opts.draft,
    attachments: opts.attachments,
    includeDoc: opts.includeDoc,
    docName: opts.docName,
    docType: opts.docType,
    selectionContext: opts.selectionContext,
  });
  const conversationTokens =
    opts.wireMessages.reduce((sum, msg) => sum + wireMessageTokens(msg), 0) +
    (pending ? estimateMessageTokens(pending) : 0);

  const raw: { id: ContextUsageCategoryId; tokens: number }[] = [
    { id: 'system', tokens: systemTokens },
    { id: 'tools', tokens: toolsTokens },
    { id: 'conversation', tokens: conversationTokens },
  ];

  const categories: ContextUsageCategory[] = raw
    .map(({ id, tokens }) => ({
      id,
      label: CONTEXT_USAGE_META[id].label,
      color: CONTEXT_USAGE_META[id].color,
      tokens: calibrator.apply(tokens),
    }))
    .filter((c) => c.tokens > 0);

  const usedTokens = categories.reduce((sum, c) => sum + c.tokens, 0);
  const percentFull =
    contextWindow > 0
      ? Math.min(100, Math.round((usedTokens / contextWindow) * 100))
      : 0;

  return {
    contextWindow,
    usedTokens,
    reservedTokens: calibrator.apply(reservedTokens),
    percentFull,
    categories,
  };
}
