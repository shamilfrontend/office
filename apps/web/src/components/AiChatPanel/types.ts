import type { Plan } from '@office/shared';
import type { ChatMode, ImpactSummary } from '@/univer/officeTools';

export type ToolCall = { id: string; name: string; arguments: string };

export type UiMsg =
  | { key: string; role: 'user'; content: string; wire: string }
  | { key: string; role: 'assistant'; content: string; toolCalls?: ToolCall[] }
  | {
      key: string;
      role: 'tool';
      name: string;
      content: string;
      full: string;
      status: 'ok' | 'error' | 'pending';
      toolCallId: string;
    };

export type ModelOption = {
  providerId: string;
  provider: string;
  model: string;
};

export type SessionLogEntry = {
  ts: string;
  chatId: string;
  chatTitle: string;
  kind: 'user' | 'assistant' | 'tool' | 'system';
  name?: string;
  content: string;
  status?: 'ok' | 'error' | 'pending';
};

export type PlanRunState = 'idle' | 'running' | 'done';

export type Tab = {
  id: string;
  title: string;
  mode: ChatMode;
  providerId: string;
  model: string;
  draft: string;
  attachments: Array<{ name: string; text: string }>;
  includeDoc: boolean;
  messages: UiMsg[];
  plan: Plan | null;
  planTitle: string | null;
  planMarkdown: string | null;
  planEdited: boolean;
  planRunState: PlanRunState;
  streaming: boolean;
  tokenRate: number | null;
  tokenMeta: string | null;
  showCheckpoint: boolean;
  wroteSinceCheckpoint: boolean;
  canRetry: boolean;
  retryReason: string | null;
  executingPlan: boolean;
  abort: AbortController | null;
  allowedTools: Set<string>;
  pendingApprove: {
    name: string;
    args: Record<string, unknown>;
    summary: ImpactSummary;
    resolve: (ok: boolean) => void;
  } | null;
};

export type LoopStopReason =
  'rejected' | 'aborted' | 'editor' | 'tools' | 'stream' | 'loop' | null;
