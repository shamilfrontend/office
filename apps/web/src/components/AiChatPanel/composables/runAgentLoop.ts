import { estimateChatContextUsage } from '@/ai/estimateChatContextUsage';
import { chatTokenCalibrator } from '@/ai/estimateTokens';
import { api } from '@/api/client';
import { isToolAutoAllowed } from '@/composables/useAiToolPermissions';
import {
  parsePlanMarkdown,
  planAllTodosDone,
  planToMarkdown,
} from '@office/shared';
import {
  isDocumentWriteTool,
  needsApprove,
  runOfficeTool,
  toolsForMode,
} from '@/univer/officeTools';
import { toWire, trimWireForProvider } from '../chatWire';
import { compact, toolResultStatus } from '../formatters';
import { streamChat } from '../streamChat';
import type { LoopStopReason, Tab, ToolCall } from '../types';
import type {
  AgentChatEmit,
  AgentChatProps,
  ProviderInfo,
  PushSessionLog,
} from './agentTypes';
import { persistAssistant, setToolRow } from './agentPersist';

const TOKEN_RATE_MS = 200;

export type AgentLoopContext = {
  tabId: string;
  props: AgentChatProps;
  emit: AgentChatEmit;
  providers: ProviderInfo[];
  findTab: (id: string) => Tab | undefined;
  patchTab: (tab: Tab, body: Record<string, unknown>) => Promise<void>;
  pushSessionLog: PushSessionLog;
  askApprove: (
    tab: Tab,
    name: string,
    args: Record<string, unknown>,
  ) => Promise<boolean>;
  flushStreamHtml: (key: string, content: string) => void;
  scheduleStreamHtml: (key: string, content: string) => void;
  refreshContextUsage: () => void;
  getLastTokenRateAt: () => number;
  setLastTokenRateAt: (value: number) => void;
};

/* eslint-disable no-await-in-loop, no-restricted-syntax -- ordered agent/tool stream processing */
export async function runAgentLoop(
  ctx: AgentLoopContext,
): Promise<LoopStopReason> {
  const {
    tabId,
    props,
    emit,
    providers,
    findTab,
    patchTab,
    pushSessionLog,
    askApprove,
    flushStreamHtml,
    scheduleStreamHtml,
    refreshContextUsage,
    getLastTokenRateAt,
    setLastTokenRateAt,
  } = ctx;

  const tab = findTab(tabId);
  if (!tab) return null;
  const signal = tab.abort?.signal;
  let steps = 0;
  let lastToolSig = '';
  let repeats = 0;
  let consecutiveToolErrors = 0;

  while (steps < 25) {
    if (signal?.aborted) return 'aborted';
    steps += 1;
    const hasDocument = Boolean(props.documentId || props.docName);
    const docType = props.docType === 'DOC' ? 'DOC' : 'SHEET';
    const tools = toolsForMode(
      tab.mode,
      tab.executingPlan,
      hasDocument,
      docType,
    );
    if (tools.length && !props.univerApi) {
      const content = 'Редактор ещё загружается, повторите.';
      tab.messages.push({
        key: crypto.randomUUID(),
        role: 'assistant',
        content,
      });
      pushSessionLog(tab, {
        kind: 'system',
        name: 'editor',
        content,
        status: 'error',
      });
      return 'editor';
    }
    const started = performance.now();
    let chars = 0;
    const assistant: { content: string; toolCalls: ToolCall[] } = {
      content: '',
      toolCalls: [],
    };
    let assistantKey = '';
    let usagePrompt: number | null = null;
    let usageCompletion: number | null = null;
    const messages = trimWireForProvider(toWire(tab));
    const heuristicPrompt = JSON.stringify(messages).length / 4;
    const promptEstimate = estimateChatContextUsage({
      mode: tab.mode,
      model: tab.model,
      wireMessages: messages,
      draft: '',
      attachments: [],
      includeDoc: false,
      buildingPlan: tab.executingPlan,
      hasDocument: Boolean(props.documentId || props.docName),
      docType: props.docType,
      calibrator: chatTokenCalibrator,
    }).usedTokens;
    const rateClock = { lastAt: getLastTokenRateAt() };

    try {
      await streamChat(
        {
          mode: tab.mode,
          messages,
          docType,
          hasDocument,
          buildingPlan: tab.executingPlan,
          providerId: tab.providerId || providers[0]?.id,
          model: tab.model || undefined,
        },
        (ev) => {
          if (ev.type === 'delta' && typeof ev.content === 'string') {
            assistant.content += ev.content;
            chars += ev.content.length;
            const now = performance.now();
            if (now - rateClock.lastAt >= TOKEN_RATE_MS) {
              rateClock.lastAt = now;
              const seconds = Math.max(0.2, (now - started) / 1000);
              tab.tokenRate = Math.round(chars / 4 / seconds);
            }
            if (!assistantKey) {
              assistantKey = crypto.randomUUID();
              tab.messages.push({
                key: assistantKey,
                role: 'assistant',
                content: assistant.content,
              });
              flushStreamHtml(assistantKey, assistant.content);
            } else {
              const row = tab.messages.find(
                (item) =>
                  item.key === assistantKey && item.role === 'assistant',
              );
              if (row && row.role === 'assistant')
                row.content = assistant.content;
              scheduleStreamHtml(assistantKey, assistant.content);
            }
          }
          if (ev.type === 'usage') {
            usagePrompt =
              typeof ev.promptTokens === 'number'
                ? ev.promptTokens
                : usagePrompt;
            usageCompletion =
              typeof ev.completionTokens === 'number'
                ? ev.completionTokens
                : usageCompletion;
            if (
              typeof ev.completionTokens === 'number' &&
              Number.isFinite(ev.completionTokens)
            ) {
              const seconds = Math.max(
                0.2,
                (performance.now() - started) / 1000,
              );
              tab.tokenRate = Math.round(ev.completionTokens / seconds);
            }
          }
          if (ev.type === 'done') {
            assistant.toolCalls = (ev.toolCalls as ToolCall[]) || [];
            if (assistant.toolCalls.length && !assistantKey) {
              assistantKey = crypto.randomUUID();
              tab.messages.push({
                key: assistantKey,
                role: 'assistant',
                content: assistant.content,
                toolCalls: assistant.toolCalls,
              });
            }
            if (assistantKey) flushStreamHtml(assistantKey, assistant.content);
          }
          if (ev.type === 'error')
            throw new Error(String(ev.error || 'AI error'));
        },
        signal ?? new AbortController().signal,
      );
      setLastTokenRateAt(rateClock.lastAt);
    } catch (error) {
      setLastTokenRateAt(rateClock.lastAt);
      const aborted =
        (error instanceof DOMException && error.name === 'AbortError') ||
        Boolean(signal?.aborted);
      if (aborted) {
        if (assistantKey) flushStreamHtml(assistantKey, assistant.content);
        await persistAssistant(tab, assistant);
        pushSessionLog(tab, {
          kind: 'system',
          name: 'abort',
          content: 'Стрим остановлен',
          status: 'error',
        });
        return 'aborted';
      }
      throw error;
    }

    if (signal?.aborted) return 'aborted';

    const promptTokens = usagePrompt ?? heuristicPrompt;
    const completionTokens = usageCompletion ?? assistant.content.length / 4;
    if (usagePrompt != null) {
      const rawEstimate = Math.round(
        promptEstimate / Math.max(0.01, chatTokenCalibrator.getFactor()),
      );
      chatTokenCalibrator.observe(rawEstimate, usagePrompt);
      refreshContextUsage();
    }
    const rate = tab.tokenRate ?? 0;
    const usageTag = usagePrompt != null || usageCompletion != null ? '' : ' ≈';
    tab.tokenMeta = `${compact(promptTokens)} → ${compact(completionTokens)} · всего ${compact(promptTokens + completionTokens)}${usageTag} · ${rate} ток/с`;

    if (assistant.toolCalls.length) {
      if (!assistantKey) {
        assistantKey = crypto.randomUUID();
        tab.messages.push({
          key: assistantKey,
          role: 'assistant',
          content: assistant.content,
          toolCalls: assistant.toolCalls,
        });
      } else {
        const row = tab.messages.find(
          (item) => item.key === assistantKey && item.role === 'assistant',
        );
        if (row && row.role === 'assistant')
          row.toolCalls = assistant.toolCalls;
      }
    }
    if (assistant.content || assistant.toolCalls.length) {
      pushSessionLog(tab, {
        kind: 'assistant',
        content: [
          assistant.content,
          assistant.toolCalls.length
            ? `toolCalls: ${JSON.stringify(assistant.toolCalls)}`
            : '',
        ]
          .filter(Boolean)
          .join('\n'),
      });
    }
    await persistAssistant(tab, assistant);
    if (!assistant.toolCalls.length) break;

    const sig = JSON.stringify(
      assistant.toolCalls.map((call) => [call.name, call.arguments]),
    );
    if (sig === lastToolSig) {
      repeats += 1;
      if (repeats >= 2) {
        const content = 'Остановка: повторяющиеся tool-calls.';
        tab.messages.push({
          key: crypto.randomUUID(),
          role: 'assistant',
          content,
        });
        pushSessionLog(tab, {
          kind: 'system',
          name: 'loop',
          content,
          status: 'error',
        });
        return 'loop';
      }
    } else {
      repeats = 0;
      lastToolSig = sig;
    }

    for (const call of assistant.toolCalls) {
      if (signal?.aborted) return 'aborted';
      let args: Record<string, unknown> = {};
      try {
        args = JSON.parse(call.arguments || '{}') as Record<string, unknown>;
      } catch {
        args = {};
      }
      const toolKey = crypto.randomUUID();
      tab.messages.push({
        key: toolKey,
        role: 'tool',
        name: call.name,
        content: 'Выполняется…',
        full: '',
        status: 'pending',
        toolCallId: call.id,
      });
      const apiRef = props.univerApi ?? null;
      if (
        !tab.allowedTools.has(call.name) &&
        !isToolAutoAllowed(call.name) &&
        needsApprove(call.name, args, apiRef)
      ) {
        const ok = await askApprove(tab, call.name, args);
        if (signal?.aborted) return 'aborted';
        if (!ok) {
          const rejected = 'Отклонено пользователем';
          setToolRow(tab, toolKey, rejected, rejected, 'error');
          pushSessionLog(tab, {
            kind: 'tool',
            name: call.name,
            content: JSON.stringify({ error: 'rejected by user', args }),
            status: 'error',
          });
          await api.addChatMessage(tab.id, {
            role: 'tool',
            content: JSON.stringify({ error: 'rejected by user' }),
            toolName: call.name,
            toolCallId: call.id,
            toolStatus: 'error',
          });
          return 'rejected';
        }
      }
      if (signal?.aborted) return 'aborted';
      // Keep runtime plan hydrated for update_plan (reload / «Сделать по плану» from markdown).
      if (
        (call.name === 'update_plan' || call.name === 'create_plan') &&
        !tab.plan &&
        tab.planMarkdown
      ) {
        tab.plan = parsePlanMarkdown(tab.planMarkdown);
      }
      let result: string;
      try {
        result = await runOfficeTool(call.name, args, {
          api: apiRef,
          docType: props.docType === 'DOC' ? 'DOC' : 'SHEET',
          plan: tab.plan,
          setPlan: (plan) => {
            tab.plan = plan;
          },
        });
      } catch (error) {
        result = JSON.stringify({
          error: error instanceof Error ? error.message : 'tool failed',
          hint: 'Исправьте аргументы tool и повторите вызов; не повторяйте тот же ошибочный payload.',
        });
      }
      if (
        (call.name === 'create_plan' || call.name === 'update_plan') &&
        tab.plan
      ) {
        if (call.name === 'create_plan') {
          tab.planEdited = false;
          tab.planRunState = 'idle';
          tab.planTitle = tab.plan.name;
          tab.planMarkdown = planToMarkdown(tab.plan);
          void patchTab(tab, {
            planTitle: tab.planTitle,
            planMarkdown: tab.planMarkdown,
            planExecuted: false,
          });
        } else if (tab.executingPlan || !tab.planEdited) {
          tab.planTitle = tab.plan.name;
          tab.planMarkdown = planToMarkdown(tab.plan);
          if (planAllTodosDone(tab.plan)) tab.planRunState = 'done';
          void patchTab(tab, {
            planTitle: tab.planTitle,
            planMarkdown: tab.planMarkdown,
          });
        }
      }
      const status = toolResultStatus(result);
      setToolRow(tab, toolKey, result.slice(0, 500), result, status);
      if (status !== 'error' && isDocumentWriteTool(call.name)) {
        tab.wroteSinceCheckpoint = true;
        emit('documentChanged');
      }
      pushSessionLog(tab, {
        kind: 'tool',
        name: call.name,
        content: `args: ${JSON.stringify(args)}\nresult: ${result}`,
        status,
      });
      await api.addChatMessage(tab.id, {
        role: 'tool',
        content: result,
        toolName: call.name,
        toolCallId: call.id,
        toolStatus: status,
      });
      if (status === 'error') {
        consecutiveToolErrors += 1;
        // During plan execution allow more recoveries — model often fixes empty formulas etc.
        const errorBudget = tab.executingPlan ? 5 : 3;
        if (consecutiveToolErrors >= errorBudget) {
          const content =
            'Остановка: несколько tool-вызовов подряд завершились ошибкой. Уточните задачу или проверьте выделение. Подробности — в логе сессии.';
          tab.messages.push({
            key: crypto.randomUUID(),
            role: 'assistant',
            content,
          });
          pushSessionLog(tab, {
            kind: 'system',
            name: 'tool-errors',
            content,
            status: 'error',
          });
          return 'tools';
        }
      } else {
        consecutiveToolErrors = 0;
      }
    }
  }
  return null;
}
/* eslint-enable no-await-in-loop, no-restricted-syntax */
