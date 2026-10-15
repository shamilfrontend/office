import { toolsForMode, type ChatMode } from '@/univer/officeTools';
import {
  MESSAGE_FRAMING_TOKENS,
  chatTokenCalibrator,
  estimateMessageTokens,
  estimateTextTokens,
  type TokenEstimateCalibrator,
} from './estimateTokens';
import { effectiveReservedTokens, getModelContextInfo } from './modelContext';

export type ContextUsageCategoryId = 'system' | 'tools' | 'conversation';

export type ContextUsageCategory = {
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

/** Зеркало system prompt из apps/api/src/ai/ai.service.ts */
export function systemPromptForMode(mode: ChatMode): string {
  if (mode === 'plan') {
    return 'Ты помощник в офисном редакторе. Всегда отвечай пользователю на русском языке. Режим План: только читай документ через tools и создавай план (create_plan/update_plan). Не изменяй документ. Если пользователь говорит «здесь/это/выделенное», используй get_selection или контекст выделения. Для таблиц сначала read_range/read_formulas/get_selection; другой лист — activate_sheet.';
  }
  if (mode === 'agent') {
    return 'Ты помощник в офисном редакторе. Всегда отвечай пользователю на русском языке. Режим Агент: можешь читать и изменять документ через tools. Если пользователь говорит «здесь/это/выделенное», опирайся на selection / get_selection. Для таблиц: другие листы — сначала activate_sheet; insert_sheet/rename_sheet/delete_sheet/set_freeze для структуры; set_cells — литералы; set_formulas — формулы (EN-имена, аргументы через запятую, ведущий =, напр. =SUM(A1:A10)); read_formulas читает формулы и значения; для оформления — format_range, set_column_widths, set_number_format, set_borders, set_row_heights; add_conditional_highlight/get_conditional_formatting/clear_conditional_formatting — условное форматирование (в т.ч. text_contains/duplicate); set_list_validation/set_number_validation/get_data_validations/clear_data_validation — валидация списка или числа; для объединения — merge_range; sort_range — сортировка (by.column: 0-based внутри диапазона; буква колонки → относительный индекс); create_filter/set_column_filter/remove_filter/get_filter — автофильтр (set_column_filter.column: буква или абсолютный 0-based индекс листа); search_in_sheet / replace_in_sheet — поиск и замена литералов (формулы не трогает); сначала read_range/read_formulas/get_selection. Для документов: get_doc_outline/search_in_document; структура — insert_paragraph, set_heading, set_list; format_text по выделению или paragraphIndex. После серии tools дай краткий итог.';
  }
  return 'Ты помощник в офисном редакторе. Всегда отвечай пользователю на русском языке. Режим Чат: читай открытый документ через tools (сначала get_workbook_outline/get_doc_outline или get_selection, затем read_range / read_formulas / read_document / search_in_document; для другого листа — activate_sheet). Не изменяй документ. Если пользователь говорит «здесь/это/выделенное», используй get_selection. Если tools доступны, не говори, что нет доступа к файлу — сначала прочитай его. Предлагай правки текстом.';
}

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
  docType?: 'SHEET' | 'DOC' | null;
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
  docType?: 'SHEET' | 'DOC' | null;
  selectionContext?: string;
  buildingPlan?: boolean;
  hasDocument?: boolean;
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

  const systemTokens = estimateMessageTokens(systemPromptForMode(opts.mode));

  const hasDocument = opts.hasDocument !== false;
  const docType = opts.docType === 'DOC' ? 'DOC' : 'SHEET';
  const tools = toolsForMode(
    opts.mode,
    Boolean(opts.buildingPlan),
    hasDocument,
    docType,
  );
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
