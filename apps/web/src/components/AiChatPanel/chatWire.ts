import type { Tab } from './types';

const TOOL_WIRE_MAX = 4000;
const ARGS_WIRE_MAX = 2000;
const MSG_WIRE_MAX = 32_000;
const WIRE_MAX_MESSAGES = 60;
const WIRE_MAX_TOTAL_CHARS = 140_000;

function clipWireText(text: string, max: number): string {
  if (text.length <= max) return text;
  return JSON.stringify({
    truncated: true,
    preview: text.slice(0, max),
    originalChars: text.length,
  });
}

function wireContentChars(msg: Record<string, unknown>): number {
  let n = 0;
  if (typeof msg.content === 'string') n += msg.content.length;
  const toolCalls = msg.tool_calls;
  if (Array.isArray(toolCalls)) {
    toolCalls.forEach((tc) => {
      const args = (tc as { function?: { arguments?: string } })?.function
        ?.arguments;
      if (typeof args === 'string') n += args.length;
    });
  }
  return n;
}

export function trimWireForProvider(wire: Array<Record<string, unknown>>) {
  const totalChars = () =>
    wire.reduce((sum, m) => sum + wireContentChars(m), 0);
  if (
    wire.length <= WIRE_MAX_MESSAGES &&
    totalChars() <= WIRE_MAX_TOTAL_CHARS
  ) {
    return wire;
  }

  const turns: Array<Array<Record<string, unknown>>> = [];
  let current: Array<Record<string, unknown>> = [];
  wire.forEach((msg) => {
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
    (turns.reduce((n, t) => n + t.length, 0) > WIRE_MAX_MESSAGES ||
      turns.reduce(
        (n, t) => n + t.reduce((s, m) => s + wireContentChars(m), 0),
        0,
      ) > WIRE_MAX_TOTAL_CHARS)
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

function closeOpenToolCalls(
  wire: Array<Record<string, unknown>>,
  openToolCallIds: Set<string> | null,
) {
  if (!openToolCallIds?.size) return;
  openToolCallIds.forEach((id) => {
    wire.push({
      role: 'tool',
      tool_call_id: id,
      content: JSON.stringify({ error: 'tool result missing' }),
    });
  });
  openToolCallIds.clear();
}

export function toWire(tab: Tab) {
  const wire: Array<Record<string, unknown>> = [];
  let openToolCallIds: Set<string> | null = null;
  tab.messages.forEach((message) => {
    if (message.role === 'user') {
      closeOpenToolCalls(wire, openToolCallIds);
      openToolCallIds = null;
      wire.push({
        role: 'user',
        content: clipWireText(message.wire, MSG_WIRE_MAX),
      });
    }
    if (message.role === 'assistant') {
      closeOpenToolCalls(wire, openToolCallIds);
      const toolCalls = message.toolCalls?.length
        ? message.toolCalls.map((call) => ({
            id: call.id,
            type: 'function' as const,
            function: {
              name: call.name,
              arguments: clipWireText(call.arguments, ARGS_WIRE_MAX),
            },
          }))
        : undefined;
      openToolCallIds = toolCalls
        ? new Set(toolCalls.map((call) => call.id))
        : null;
      const content = message.content
        ? clipWireText(message.content, MSG_WIRE_MAX)
        : null;
      wire.push({
        role: 'assistant',
        content,
        ...(toolCalls ? { tool_calls: toolCalls } : {}),
      });
    }
    if (message.role === 'tool' && message.status !== 'pending') {
      if (message.toolCallId && openToolCallIds?.has(message.toolCallId)) {
        openToolCallIds.delete(message.toolCallId);
        wire.push({
          role: 'tool',
          tool_call_id: message.toolCallId,
          name: message.name,
          content: clipWireText(message.full || message.content, TOOL_WIRE_MAX),
        });
      }
    }
  });
  closeOpenToolCalls(wire, openToolCallIds);
  return wire;
}
