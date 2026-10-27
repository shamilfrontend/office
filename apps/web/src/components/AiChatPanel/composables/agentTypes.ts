import type { ComputedRef, Ref } from 'vue';
import type { OfficeToolDef } from '@office/shared';
import type { UniverApi } from '@/univer/useUniver';
import type { SessionLogEntry, Tab } from '../types';

export type EditorDocType = 'SHEET' | 'DOC' | 'SLIDE';

export type AgentChatProps = {
  docType?: EditorDocType | null;
  docName?: string;
  documentId?: string;
  univerApi?: UniverApi | null;
  slideReady?: boolean;
  getSlideTools?: () => OfficeToolDef[];
  runSlideTool?: (name: string, input: unknown) => Promise<unknown>;
  getSelectionContext?: () => string;
};

export type AgentChatEmit = {
  (e: 'checkpoint'): void;
  (e: 'documentChanged'): void;
};

export type ProviderInfo = {
  id: string;
  provider?: string;
  model?: string | null;
};

export type PushSessionLog = (
  tab: Tab | null | undefined,
  entry: Omit<SessionLogEntry, 'ts' | 'chatId' | 'chatTitle'> & {
    chatId?: string;
    chatTitle?: string;
  },
) => void;

export type UseAgentChatOptions = {
  props: AgentChatProps;
  emit: AgentChatEmit;
  active: ComputedRef<Tab | null>;
  hasProvider: ComputedRef<boolean>;
  providers: ComputedRef<ProviderInfo[]>;
  attachError: Ref<string>;
  approveExpanded: Ref<boolean>;
  findTab: (id: string) => Tab | undefined;
  patchTab: (tab: Tab, body: Record<string, unknown>) => Promise<void>;
  pushSessionLog: PushSessionLog;
  toggleModelMenu: () => void | Promise<void>;
  resizeComposer: () => void;
  refreshContextUsage: () => void;
  clearStreamHtml: () => void;
  flushStreamHtml: (key: string, content: string) => void;
  scheduleStreamHtml: (key: string, content: string) => void;
};
