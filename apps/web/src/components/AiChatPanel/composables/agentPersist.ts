import { api } from '@/api/client';
import type { Tab, ToolCall } from '../types';

export function setToolRow(
  tab: Tab,
  key: string,
  content: string,
  full: string,
  status: 'ok' | 'error',
) {
  const row = tab.messages.find(
    (item) => item.key === key && item.role === 'tool',
  );
  if (row && row.role === 'tool') {
    row.content = content;
    row.full = full;
    row.status = status;
  }
}

export async function persistAssistant(
  tab: Tab,
  assistant: { content: string; toolCalls: ToolCall[] },
) {
  if (!assistant.content && !assistant.toolCalls.length) return;
  await api.addChatMessage(tab.id, {
    role: 'assistant',
    content: assistant.content,
    toolCallsJson: assistant.toolCalls.length
      ? JSON.stringify(assistant.toolCalls)
      : undefined,
  });
}
