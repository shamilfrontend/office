import { ALL_TOOL_NAMES, SLIDE_TOOL_NAMES, type DocType } from './office-tools';

export type ToolTarget = 'slide' | 'office' | 'reject' | 'unknown';

const PLAN_TOOL_NAMES = new Set(['create_plan', 'update_plan']);

/**
 * Where a tool call may run.
 * Slide tools stay in the presentation editor. Sheet and document tools
 * never run while a presentation is open, and slide tools never run in Univer.
 */
export function resolveToolTarget(docType: DocType, name: string): ToolTarget {
  if (!ALL_TOOL_NAMES.has(name)) return 'unknown';
  if (docType === 'SLIDE') {
    if (SLIDE_TOOL_NAMES.has(name)) return 'slide';
    if (PLAN_TOOL_NAMES.has(name)) return 'office';
    return 'reject';
  }
  if (SLIDE_TOOL_NAMES.has(name)) return 'reject';
  return 'office';
}

export function toolTargetResult(
  target: 'reject' | 'unknown',
  name: string,
): string {
  if (target === 'unknown') {
    return JSON.stringify({ error: `unknown tool ${name}` });
  }
  return JSON.stringify({
    error: 'tool is not available for this document',
  });
}

export async function dispatchOfficeTool(input: {
  docType: DocType;
  name: string;
  args: Record<string, unknown>;
  runSlide: (name: string, args: Record<string, unknown>) => Promise<unknown>;
  runOffice: (
    name: string,
    args: Record<string, unknown>,
    docType: 'SHEET' | 'DOC',
  ) => Promise<string>;
}): Promise<string> {
  const target = resolveToolTarget(input.docType, input.name);
  if (target === 'unknown' || target === 'reject') {
    return toolTargetResult(target, input.name);
  }
  if (target === 'slide') {
    const value = await input.runSlide(input.name, input.args);
    return JSON.stringify(value ?? null);
  }
  const officeType = input.docType === 'DOC' ? 'DOC' : 'SHEET';
  return input.runOffice(input.name, input.args, officeType);
}
