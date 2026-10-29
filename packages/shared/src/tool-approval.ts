import { estimateA1Area } from './a1';
import { SLIDE_WRITE_TOOL_NAMES } from './office-tools';

const SLIDE_WRITE_TOOLS = new Set<string>(SLIDE_WRITE_TOOL_NAMES);

const DATA_APPROVE_CELLS = 20;
const FORMAT_APPROVE_CELLS = 50;
const LAYOUT_APPROVE_COUNT = 5;
const DOC_APPROVE_CHARS = 200;

export type ApproveProbe = {
  /** Cell count for a range argument, including the active selection when it is empty. */
  area?: (rangeArg: unknown) => number;
  docSelectionLength?: number;
};

function gridCellCount(grid: unknown[][] | undefined): number {
  return grid?.reduce((n, row) => n + (row?.length ?? 0), 0) ?? 0;
}

/** Whether a tool call must be confirmed before it changes the open file. */
export function needsApprove(
  name: string,
  args: Record<string, unknown>,
  probe: ApproveProbe = {},
): boolean {
  if (SLIDE_WRITE_TOOLS.has(name)) return true;
  if (
    name === 'clear_range' ||
    name === 'merge_range' ||
    name === 'replace_text' ||
    name === 'replace_in_sheet' ||
    name === 'delete_sheet' ||
    name === 'insert_sheet' ||
    name === 'rename_sheet' ||
    name === 'set_freeze' ||
    name === 'sort_range' ||
    name === 'create_filter' ||
    name === 'remove_filter' ||
    name === 'set_column_filter' ||
    name === 'add_conditional_highlight' ||
    name === 'clear_conditional_formatting' ||
    name === 'set_list_validation' ||
    name === 'set_number_validation' ||
    name === 'clear_data_validation'
  ) {
    return true;
  }
  if (name === 'set_cells') {
    return (
      gridCellCount(args.values as unknown[][] | undefined) > DATA_APPROVE_CELLS
    );
  }
  if (name === 'set_formulas') {
    return (
      gridCellCount(args.formulas as unknown[][] | undefined) >
      DATA_APPROVE_CELLS
    );
  }
  if (name === 'copy_formulas') {
    const times = Math.max(1, Number(args.times) || 1);
    return (
      times > 3 ||
      estimateA1Area(String(args.source || '')) * times > DATA_APPROVE_CELLS
    );
  }
  if (name === 'insert_image') return true;
  if (
    name === 'format_range' ||
    name === 'set_number_format' ||
    name === 'set_borders'
  ) {
    const area = probe.area ? probe.area(args.range) : 0;
    return area > FORMAT_APPROVE_CELLS;
  }
  if (name === 'set_column_widths') {
    const widths = args.widths as unknown[] | undefined;
    if (Array.isArray(widths) && widths.length > LAYOUT_APPROVE_COUNT)
      return true;
    const num = Number(args.numColumns);
    return Number.isFinite(num) && num > LAYOUT_APPROVE_COUNT;
  }
  if (name === 'set_row_heights') {
    const heights = args.heights as unknown[] | undefined;
    if (Array.isArray(heights) && heights.length > LAYOUT_APPROVE_COUNT)
      return true;
    const num = Number(args.numRows);
    return Number.isFinite(num) && num > LAYOUT_APPROVE_COUNT;
  }
  if (name === 'replace_selection') return true;
  if (name === 'set_heading' || name === 'set_list') return true;
  if (name === 'insert_text') {
    const text = String(args.text ?? '');
    const at = args.at != null ? String(args.at) : 'end';
    return text.length > DOC_APPROVE_CHARS || at === 'selection';
  }
  if (name === 'insert_paragraph') {
    return String(args.text ?? '').length > DOC_APPROVE_CHARS;
  }
  if (name === 'format_text') {
    return (probe.docSelectionLength ?? 0) > DOC_APPROVE_CHARS;
  }
  return false;
}
