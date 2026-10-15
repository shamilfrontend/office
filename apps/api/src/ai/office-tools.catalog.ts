import type OpenAI from 'openai';
import {
  ALL_TOOL_NAMES,
  OFFICE_TOOLS as SHARED_TOOLS,
  toolsForMode as sharedToolsForMode,
  type ChatMode,
  type DocType,
  type OfficeToolDef,
} from '@office/shared';

export type { ChatMode, DocType };

/** Re-export shared catalog typed for OpenAI SDK. */
export const OFFICE_TOOLS =
  SHARED_TOOLS as OpenAI.Chat.Completions.ChatCompletionTool[];

export { ALL_TOOL_NAMES };

export function toolsForMode(
  mode: ChatMode,
  buildingPlan = false,
  hasDocument = true,
  docType: DocType = 'SHEET',
): OpenAI.Chat.Completions.ChatCompletionTool[] {
  return sharedToolsForMode(
    mode,
    buildingPlan,
    hasDocument,
    docType,
  ) as OpenAI.Chat.Completions.ChatCompletionTool[];
}

export type { OfficeToolDef };
