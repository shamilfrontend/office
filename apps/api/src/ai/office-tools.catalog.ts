import type OpenAI from 'openai';
import {
  ALL_TOOL_NAMES,
  slideToolNamesForMode,
  toolsForMode as sharedToolsForMode,
  type ChatMode,
  type DocType,
} from '@office/shared';

export type { ChatMode, DocType };

export { ALL_TOOL_NAMES, slideToolNamesForMode };

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
