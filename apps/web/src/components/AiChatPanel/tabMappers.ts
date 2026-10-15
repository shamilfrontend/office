import { parsePlanMarkdown, planAllTodosDone } from '@office/shared';
import type { ChatRecord } from '@/api/client';
import { asMode, visibleUser } from './formatters';
import type { Tab, ToolCall, UiMsg } from './types';

type ProviderInfo = { id: string; model?: string | null };

export function planWasExecuted(record: ChatRecord): boolean {
  if (record.planExecuted === true) return true;
  if (record.planExecuted === false) return false;
  let lastCreate = -1;
  let lastRun = -1;
  record.messages.forEach((message, index) => {
    if (message.role === 'tool' && message.toolName === 'create_plan')
      lastCreate = index;
    if (
      message.role === 'user' &&
      message.content.startsWith('Сделать по плану')
    ) {
      lastRun = index;
    }
  });
  if (lastRun < 0 || lastRun < lastCreate) return false;
  const assistant = [...record.messages.slice(lastRun + 1)]
    .reverse()
    .find((message) => message.role === 'assistant');
  if (!assistant) return false;
  return !/Остановлено|Отклонено/.test(assistant.content);
}

export function messageFromRecord(row: ChatRecord['messages'][number]): UiMsg {
  if (row.role === 'tool') {
    return {
      key: row.id,
      role: 'tool',
      name: row.toolName || 'tool',
      content: row.content.slice(0, 500),
      full: row.content,
      status: row.toolStatus === 'error' ? 'error' : 'ok',
      toolCallId: row.toolCallId || '',
    };
  }
  if (row.role === 'assistant') {
    let toolCalls: ToolCall[] | undefined;
    if (row.toolCallsJson) {
      try {
        toolCalls = JSON.parse(row.toolCallsJson) as ToolCall[];
      } catch {
        toolCalls = undefined;
      }
    }
    return { key: row.id, role: 'assistant', content: row.content, toolCalls };
  }
  return {
    key: row.id,
    role: 'user',
    content: visibleUser(row.content),
    wire: row.content,
  };
}

export function tabFromRecord(
  record: ChatRecord,
  providers: ProviderInfo[],
): Tab {
  const providerId = record.providerId || providers[0]?.id || '';
  const providerModel =
    providers.find((item) => item.id === providerId)?.model ||
    providers[0]?.model ||
    '';
  const plan = record.planMarkdown
    ? parsePlanMarkdown(record.planMarkdown)
    : null;
  return {
    id: record.id,
    title: record.title || 'Новый чат',
    mode: asMode(record.mode),
    providerId,
    model: record.model || providerModel || '',
    draft: '',
    attachments: [],
    includeDoc: true,
    messages: record.messages.map(messageFromRecord),
    plan,
    planTitle: record.planTitle,
    planMarkdown: record.planMarkdown,
    planEdited: false,
    planRunState:
      planWasExecuted(record) || planAllTodosDone(plan) ? 'done' : 'idle',
    streaming: false,
    tokenRate: null,
    tokenMeta: null,
    showCheckpoint: false,
    wroteSinceCheckpoint: false,
    canRetry: false,
    retryReason: null,
    executingPlan: false,
    abort: null,
    allowedTools: new Set(),
    pendingApprove: null,
  };
}
