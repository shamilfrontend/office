import {
  BorderStyleTypes,
  BorderType,
  ICommandService,
  NamedStyleType,
  PresetListType,
} from '@univerjs/core';
import {
  BulletListCommand,
  CheckListCommand,
  ListOperationCommand,
  OrderListCommand,
  SetParagraphNamedStyleCommand,
} from '@univerjs/docs-ui';
import {
  adjustFormulaRefs,
  collectFormulaCells,
  indexToColLetters,
  mergePlanTodos,
  normalizePlanStatus,
  WRITE_TOOL_NAMES,
} from '@office/shared';
import {
  a1FromRowCol,
  a1Range,
  estimateA1Area,
  getActiveWorkbook,
  getDocActiveOffsets,
  getDocSelection,
  getSelectionInfo,
  getSelectionInjector,
  parseA1Cell,
  resolveColumnIndex,
  type CfRuleLike,
  type DataValidationRuleLike,
  type SheetFilterHandle,
  type SheetHandle,
  type SheetRange,
  type SheetSortBy,
} from './selection';
import type { DocParagraph, UniverApi, UniverDocument } from './useUniver';

export { toolsForMode } from '@office/shared';
export type { ChatMode } from '@office/shared';

const READ_MAX_ROWS = 80;
const READ_MAX_COLS = 40;
const READ_MAX_CELLS = 1200;
const SEARCH_MAX_HITS = 50;
const DATA_APPROVE_CELLS = 20;
const FORMAT_APPROVE_CELLS = 50;
const LAYOUT_APPROVE_COUNT = 5;
const DOC_APPROVE_CHARS = 200;

export type ImpactSummary = { title: string; detail: string };

export function isDocumentWriteTool(name: string): boolean {
  return WRITE_TOOL_NAMES.has(name);
}

function gridCellCount(grid: unknown[][] | undefined): number {
  return grid?.reduce((n, row) => n + (row?.length ?? 0), 0) ?? 0;
}

function resolvedRangeA1(api: UniverApi | null, rangeArg: unknown): string {
  const a1 = rangeArg != null ? String(rangeArg).trim() : '';
  if (a1) return a1;
  const wb = getActiveWorkbook(api);
  if (!wb) return '';
  const sheet = wb.getActiveSheet();
  const active = sheet.getActiveRange?.() ?? wb.getActiveRange?.();
  return active?.getA1Notation?.() ?? '';
}

function resolvedArea(api: UniverApi | null, rangeArg: unknown): number {
  const a1 = resolvedRangeA1(api, rangeArg);
  if (!a1) return 0;
  return estimateA1Area(a1);
}

export function needsApprove(
  name: string,
  args: Record<string, unknown>,
  api: UniverApi | null,
): boolean {
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
    return resolvedArea(api, args.range) > FORMAT_APPROVE_CELLS;
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
    const sel = getDocSelection(api);
    const len = sel?.text?.length ?? 0;
    return len > DOC_APPROVE_CHARS;
  }
  return false;
}

export function describeImpact(
  name: string,
  args: Record<string, unknown>,
  api: UniverApi | null,
): ImpactSummary {
  const title = name;
  if (name === 'set_cells') {
    const cells = gridCellCount(args.values as unknown[][] | undefined);
    const start = args.start != null ? String(args.start) : 'активная ячейка';
    return { title, detail: `~${cells} ячеек от ${start}` };
  }
  if (name === 'set_formulas') {
    const cells = gridCellCount(args.formulas as unknown[][] | undefined);
    const start = args.start != null ? String(args.start) : 'активная ячейка';
    return { title, detail: `~${cells} формул от ${start}` };
  }
  if (name === 'copy_formulas') {
    const times = Math.max(1, Number(args.times) || 1);
    return {
      title,
      detail: `${String(args.source || '?')} → ${String(args.destination || '?')} ×${times}`,
    };
  }
  if (name === 'insert_image') {
    return {
      title,
      detail: String(args.start || 'A1'),
    };
  }
  if (
    name === 'clear_range' ||
    name === 'merge_range' ||
    name === 'format_range' ||
    name === 'set_number_format' ||
    name === 'set_borders'
  ) {
    const a1 = resolvedRangeA1(api, args.range) || 'выделение';
    const area = a1 === 'выделение' ? 0 : estimateA1Area(a1);
    return {
      title,
      detail: area > 0 ? `${a1} (~${area} ячеек)` : a1,
    };
  }
  if (name === 'set_column_widths') {
    const widths = args.widths as unknown[] | undefined;
    const n = Array.isArray(widths)
      ? widths.length
      : Number(args.numColumns) || 0;
    return { title, detail: `~${n} столбцов` };
  }
  if (name === 'set_row_heights') {
    const heights = args.heights as unknown[] | undefined;
    const n = Array.isArray(heights)
      ? heights.length
      : Number(args.numRows) || 0;
    return { title, detail: `~${n} строк` };
  }
  if (name === 'replace_text' || name === 'replace_in_sheet') {
    const search = String(args.search ?? '').slice(0, 40);
    return { title, detail: `замена «${search}»` };
  }
  if (name === 'replace_selection' || name === 'insert_text') {
    const text = String(args.text ?? '');
    return { title, detail: `~${text.length} символов` };
  }
  if (name === 'format_text') {
    const sel = getDocSelection(api);
    const len = sel?.text?.length ?? 0;
    return { title, detail: `выделение ~${len} символов` };
  }
  if (name === 'insert_paragraph') {
    return { title, detail: `~${String(args.text ?? '').length} символов` };
  }
  if (name === 'set_heading') {
    return {
      title,
      detail: `стиль ${String(args.style ?? '')}${
        args.paragraphIndex != null ? `, ¶${args.paragraphIndex}` : ''
      }`,
    };
  }
  if (name === 'set_list') {
    return {
      title,
      detail: `список ${String(args.type ?? '')}${
        args.paragraphIndex != null ? `, ¶${args.paragraphIndex}` : ''
      }`,
    };
  }
  if (
    name === 'delete_sheet' ||
    name === 'rename_sheet' ||
    name === 'insert_sheet'
  ) {
    let sheetName = 'активный лист';
    if (args.name != null) sheetName = String(args.name);
    else if (args.newName != null) sheetName = String(args.newName);
    return { title, detail: sheetName };
  }
  if (name === 'set_freeze') {
    if (args.clear) return { title, detail: 'снять закрепление' };
    const rows = args.rows != null ? Number(args.rows) : undefined;
    const cols = args.columns != null ? Number(args.columns) : undefined;
    const parts: string[] = [];
    if (Number.isFinite(rows)) parts.push(`${rows} строк`);
    if (Number.isFinite(cols)) parts.push(`${cols} столбцов`);
    return { title, detail: parts.length ? parts.join(', ') : 'закрепление' };
  }
  if (name === 'sort_range') {
    const a1 = resolvedRangeA1(api, args.range) || 'выделение';
    const area = a1 === 'выделение' ? 0 : estimateA1Area(a1);
    return {
      title,
      detail: area > 0 ? `${a1} (~${area} ячеек)` : a1,
    };
  }
  if (name === 'create_filter') {
    const a1 = resolvedRangeA1(api, args.range) || 'выделение';
    return { title, detail: a1 };
  }
  if (name === 'remove_filter') {
    return { title, detail: 'активный фильтр' };
  }
  if (name === 'set_column_filter') {
    const col = args.column != null ? String(args.column) : '?';
    if (args.clear) return { title, detail: `очистить колонку ${col}` };
    const values = Array.isArray(args.values) ? args.values.length : 0;
    return { title, detail: `колонка ${col}, ${values} значений` };
  }
  if (name === 'add_conditional_highlight') {
    const a1 = resolvedRangeA1(api, args.range) || 'выделение';
    const condition = String(args.condition ?? '');
    const extra =
      condition === 'text_contains' && args.text != null
        ? ` «${String(args.text)}»`
        : '';
    return {
      title,
      detail: `${condition}${extra} на ${a1}`,
    };
  }
  if (name === 'clear_conditional_formatting') {
    if (args.all) return { title, detail: 'все правила листа' };
    if (args.cfId != null)
      return { title, detail: `cfId ${String(args.cfId)}` };
    const a1 = resolvedRangeA1(api, args.range) || 'выделение';
    return { title, detail: a1 };
  }
  if (name === 'set_list_validation') {
    const a1 = resolvedRangeA1(api, args.range) || 'выделение';
    const n = Array.isArray(args.values) ? args.values.length : 0;
    return { title, detail: `${a1}, ${n} значений` };
  }
  if (name === 'set_number_validation') {
    const a1 = resolvedRangeA1(api, args.range) || 'выделение';
    const op = String(args.op ?? '');
    let bound = '';
    if (op === 'between') bound = ` [${args.min}, ${args.max}]`;
    else if (args.value != null) bound = ` ${args.value}`;
    return { title, detail: `${a1}, ${op}${bound}` };
  }
  if (name === 'clear_data_validation') {
    const a1 = resolvedRangeA1(api, args.range) || 'выделение';
    return { title, detail: a1 };
  }
  return { title, detail: '' };
}

function mapBorderType(value: string): BorderType | null {
  const map: Record<string, BorderType> = {
    all: BorderType.ALL,
    outside: BorderType.OUTSIDE,
    inside: BorderType.INSIDE,
    top: BorderType.TOP,
    bottom: BorderType.BOTTOM,
    left: BorderType.LEFT,
    right: BorderType.RIGHT,
    horizontal: BorderType.HORIZONTAL,
    vertical: BorderType.VERTICAL,
    none: BorderType.NONE,
  };
  return map[value] ?? null;
}

function mapBorderStyle(value: string): BorderStyleTypes | null {
  const map: Record<string, BorderStyleTypes> = {
    none: BorderStyleTypes.NONE,
    thin: BorderStyleTypes.THIN,
    medium: BorderStyleTypes.MEDIUM,
    thick: BorderStyleTypes.THICK,
    dashed: BorderStyleTypes.DASHED,
    dotted: BorderStyleTypes.DOTTED,
    double: BorderStyleTypes.DOUBLE,
  };
  return map[value] ?? null;
}

const HEADING_STYLE_MAP: Record<string, NamedStyleType> = {
  normal: NamedStyleType.NORMAL_TEXT,
  title: NamedStyleType.TITLE,
  subtitle: NamedStyleType.SUBTITLE,
  h1: NamedStyleType.HEADING_1,
  h2: NamedStyleType.HEADING_2,
  h3: NamedStyleType.HEADING_3,
  h4: NamedStyleType.HEADING_4,
  h5: NamedStyleType.HEADING_5,
};

const HEADING_STYLE_LABEL: Partial<Record<NamedStyleType, string>> = {
  [NamedStyleType.NORMAL_TEXT]: 'normal',
  [NamedStyleType.TITLE]: 'title',
  [NamedStyleType.SUBTITLE]: 'subtitle',
  [NamedStyleType.HEADING_1]: 'h1',
  [NamedStyleType.HEADING_2]: 'h2',
  [NamedStyleType.HEADING_3]: 'h3',
  [NamedStyleType.HEADING_4]: 'h4',
  [NamedStyleType.HEADING_5]: 'h5',
};

function isHeadingStyle(style: NamedStyleType | undefined): boolean {
  if (style == null) return false;
  return (
    style === NamedStyleType.TITLE ||
    style === NamedStyleType.SUBTITLE ||
    style === NamedStyleType.HEADING_1 ||
    style === NamedStyleType.HEADING_2 ||
    style === NamedStyleType.HEADING_3 ||
    style === NamedStyleType.HEADING_4 ||
    style === NamedStyleType.HEADING_5
  );
}

function resolveDocParagraph(
  doc: UniverDocument,
  paragraphIndexArg: unknown,
): { paragraph: DocParagraph; index: number } | { error: string } {
  const paragraphs =
    typeof doc.getParagraphs === 'function' ? doc.getParagraphs() : [];
  if (!paragraphs.length) return { error: 'no paragraphs' };

  if (paragraphIndexArg != null && String(paragraphIndexArg).trim() !== '') {
    const index = Number(paragraphIndexArg);
    if (!Number.isInteger(index) || index < 0 || index >= paragraphs.length) {
      return { error: `invalid paragraphIndex: ${String(paragraphIndexArg)}` };
    }
    return { paragraph: paragraphs[index], index };
  }

  const offsets = getDocActiveOffsets();
  if (offsets) {
    const mid = Math.floor((offsets.startOffset + offsets.endOffset) / 2);
    const matched = paragraphs.findIndex((p) => {
      const info = p.getInfo?.();
      return Boolean(info && mid >= info.startOffset && mid <= info.endOffset);
    });
    if (matched >= 0) {
      const info = paragraphs[matched].getInfo?.();
      return {
        paragraph: paragraphs[matched],
        index: info?.paragraphIndex ?? matched,
      };
    }
  }

  return { paragraph: paragraphs[0], index: 0 };
}

async function executeDocCommand(
  id: string,
  params?: Record<string, unknown>,
): Promise<boolean> {
  const injector = getSelectionInjector();
  if (!injector) return false;
  try {
    const commandService = injector.get(ICommandService) as {
      executeCommand: (
        commandId: string,
        params?: Record<string, unknown>,
      ) => boolean | Promise<boolean>;
    };
    return Boolean(await commandService.executeCommand(id, params));
  } catch {
    return false;
  }
}

function resolveSheetByName(
  wb: NonNullable<ReturnType<typeof getActiveWorkbook>>,
  nameArg: unknown,
  allowActiveFallback: boolean,
): { sheet: SheetHandle } | { error: string } {
  const name = nameArg != null ? String(nameArg).trim() : '';
  if (!name) {
    if (!allowActiveFallback) return { error: 'sheet name required' };
    return { sheet: wb.getActiveSheet() };
  }
  const byName = wb.getSheetByName?.(name);
  if (byName) return { sheet: byName };
  const match = wb.getSheets().find((s) => s.getSheetName() === name);
  if (match) return { sheet: match };
  return { error: `sheet not found: ${name}` };
}

function resolveSheetRangeA1(
  sheet: SheetHandle,
  wb: NonNullable<ReturnType<typeof getActiveWorkbook>>,
  rangeArg: unknown,
): { a1: string; range: SheetRange } | { error: string } {
  let a1 = rangeArg != null ? String(rangeArg).trim() : '';
  if (!a1) {
    const active = sheet.getActiveRange?.() ?? wb.getActiveRange?.();
    a1 = active?.getA1Notation?.() ?? '';
  }
  if (!a1) return { error: 'no range; pass range or select cells' };
  try {
    return { a1, range: sheet.getRange(a1) };
  } catch (error) {
    return {
      error: error instanceof Error ? error.message : 'invalid range',
    };
  }
}

function getRangeStartColumn(range: SheetRange, a1: string): number {
  const bounds = range.getRange?.();
  if (bounds && typeof bounds.startColumn === 'number')
    return bounds.startColumn;
  if (typeof range.getColumn === 'function') {
    const col = range.getColumn();
    if (typeof col === 'number' && Number.isFinite(col)) return col;
  }
  const cell =
    a1
      .replace(/^[^!]*!/, '')
      .split(':')[0]
      ?.trim() ?? '';
  const withRow = /^\$?[A-Za-z]+$/.test(cell) ? `${cell}1` : cell;
  return parseA1Cell(withRow)?.col ?? 0;
}

function resolveSortColumnRef(
  column: unknown,
  startColumn: number,
): number | { error: string } {
  if (typeof column === 'number' && Number.isFinite(column)) {
    const n = Math.floor(column);
    return n >= 0 ? n : { error: `invalid sort column: ${column}` };
  }
  const text = String(column ?? '').trim();
  if (!text) return { error: 'sort column required' };
  if (/^[A-Za-z]+$/.test(text)) {
    const abs = resolveColumnIndex(text);
    if (abs == null) return { error: `invalid sort column: ${text}` };
    const relative = abs - startColumn;
    return relative >= 0 ? relative : { error: `column ${text} outside range` };
  }
  const n = resolveColumnIndex(text);
  if (n == null) return { error: `invalid sort column: ${text}` };
  return n;
}

function parseSortBy(
  byArg: unknown,
  startColumn: number,
): SheetSortBy | { error: string } {
  const one = (
    item: unknown,
  ): number | { column: number; ascending?: boolean } | { error: string } => {
    if (typeof item === 'number' || typeof item === 'string') {
      return resolveSortColumnRef(item, startColumn);
    }
    if (item && typeof item === 'object') {
      const obj = item as { column?: unknown; ascending?: unknown };
      const col = resolveSortColumnRef(obj.column, startColumn);
      if (typeof col !== 'number') return col;
      if (obj.ascending === undefined) return { column: col };
      return { column: col, ascending: Boolean(obj.ascending) };
    }
    return { error: 'invalid sort by' };
  };
  if (Array.isArray(byArg)) {
    if (!byArg.length) return { error: 'by required' };
    const out: Array<{ column: number; ascending?: boolean }> = [];
    const error = byArg.reduce<{ error: string } | null>((err, item) => {
      if (err) return err;
      const parsed = one(item);
      if (typeof parsed === 'object' && 'error' in parsed) return parsed;
      if (typeof parsed === 'number') out.push({ column: parsed });
      else out.push(parsed);
      return null;
    }, null);
    if (error) return error;
    return out;
  }
  return one(byArg);
}

function serializeFilter(filter: SheetFilterHandle): {
  range: string | null;
  criteria: Array<{ column: number; criteria: unknown }>;
  filteredOutRows: number[];
} {
  const rangeObj = filter.getRange?.();
  let rangeA1: string | null = null;
  let startCol = 0;
  let endCol = -1;
  if (rangeObj) {
    if (
      'getA1Notation' in rangeObj &&
      typeof rangeObj.getA1Notation === 'function'
    ) {
      rangeA1 = rangeObj.getA1Notation() ?? null;
    }
    if ('getRange' in rangeObj && typeof rangeObj.getRange === 'function') {
      const bounds = rangeObj.getRange();
      if (bounds) {
        startCol = bounds.startColumn;
        endCol = bounds.endColumn;
      }
    }
  }
  const criteria: Array<{ column: number; criteria: unknown }> = [];
  if (endCol >= startCol && filter.getColumnFilterCriteria) {
    Array.from(
      { length: endCol - startCol + 1 },
      (unused, i) => startCol + i,
    ).forEach((col) => {
      const c = filter.getColumnFilterCriteria?.(col);
      if (c) criteria.push({ column: col, criteria: c });
    });
  }
  const rawOut = filter.getFilteredOutRows?.();
  let filteredOutRows: number[] = [];
  if (Array.isArray(rawOut)) filteredOutRows = rawOut;
  else if (rawOut) filteredOutRows = Array.from(rawOut);
  return { range: rangeA1, criteria, filteredOutRows };
}

function mapHorizontalAlign(
  value: string,
): 'left' | 'center' | 'normal' | null {
  if (value === 'left' || value === 'center') return value;
  if (value === 'right') return 'normal';
  return null;
}

/** Case-insensitive substring replace (all occurrences); not regex. */
function replaceAllCaseInsensitive(
  text: string,
  search: string,
  replacement: string,
): string {
  if (!search) return text;
  const lower = text.toLowerCase();
  const needle = search.toLowerCase();
  let out = '';
  let i = 0;
  while (i < text.length) {
    const at = lower.indexOf(needle, i);
    if (at < 0) {
      out += text.slice(i);
      break;
    }
    out += text.slice(i, at) + replacement;
    i = at + search.length;
  }
  return out;
}

function a1FromIRange(r: {
  startRow: number;
  startColumn: number;
  endRow: number;
  endColumn: number;
}): string {
  const start = a1FromRowCol(r.startRow, r.startColumn);
  const end = a1FromRowCol(r.endRow, r.endColumn);
  return start === end ? start : `${start}:${end}`;
}

function summarizeCfRule(rule: CfRuleLike): {
  cfId: string | null;
  ranges: string[];
  condition: string;
} {
  const cfg = rule.rule;
  let condition = 'unknown';
  if (cfg && typeof cfg === 'object') {
    const parts = [cfg.subType, cfg.operator, cfg.type].filter(Boolean);
    condition = parts.length ? parts.map(String).join(':') : 'unknown';
  }
  return {
    cfId: rule.cfId ?? null,
    ranges: (rule.ranges ?? []).map(a1FromIRange),
    condition,
  };
}

function serializeDataValidation(rule: DataValidationRuleLike): {
  uid: string | null;
  type: string | null;
  criteria: [string | undefined, string | undefined, string | undefined] | null;
  ranges: string[];
} {
  const ranges =
    rule
      .getRanges?.()
      .map((r) => r.getA1Notation?.() ?? '')
      .filter(Boolean) ?? [];
  return {
    uid: rule.rule?.uid ?? rule.uid ?? null,
    type: rule.getCriteriaType?.() ?? rule.rule?.type ?? null,
    criteria: rule.getCriteriaValues?.() ?? null,
    ranges,
  };
}

function truncateMatrix(values: unknown[][]): {
  values: unknown[][];
  truncated: boolean;
  rows: number;
  cols: number;
} {
  const rows = values.length;
  const cols = values.reduce((m, row) => Math.max(m, row?.length ?? 0), 0);
  const cells = rows * Math.max(cols, 1);
  if (
    rows <= READ_MAX_ROWS &&
    cols <= READ_MAX_COLS &&
    cells <= READ_MAX_CELLS
  ) {
    return { values, truncated: false, rows, cols };
  }
  return {
    values: values
      .slice(0, READ_MAX_ROWS)
      .map((row) => (row ?? []).slice(0, READ_MAX_COLS)),
    truncated: true,
    rows,
    cols,
  };
}

export async function runOfficeTool(
  name: string,
  args: Record<string, unknown>,
  ctx: {
    api: UniverApi | null;
    docType: 'SHEET' | 'DOC';
    plan: {
      name: string;
      overview: string;
      todos: Array<{ id: string; content: string; status?: string }>;
    } | null;
    setPlan: (p: typeof ctx.plan) => void;
  },
): Promise<string> {
  if (name === 'create_plan') {
    const rawTodos =
      (args.todos as Array<{
        id?: string;
        content?: string;
        status?: string;
      }>) || [];
    const plan = {
      name: String(args.name || 'План'),
      overview: String(args.overview || ''),
      todos: rawTodos.map((t, i) => ({
        id: String(t.id || i + 1),
        content: String(t.content || ''),
        ...(t.status ? { status: normalizePlanStatus(t.status) } : {}),
      })),
    };
    ctx.setPlan(plan);
    return JSON.stringify({ ok: true, plan });
  }
  if (name === 'update_plan') {
    if (!ctx.plan) {
      return JSON.stringify({
        error: 'no plan',
        hint: 'План не загружен в runtime. Повторите update_plan после создания плана или перезапустите «Сделать по плану».',
      });
    }
    const todos = Array.isArray(args.todos)
      ? (args.todos as Array<{
          id?: string;
          content?: string;
          status?: string;
        }>)
      : null;
    if (!todos?.length) {
      return JSON.stringify({ error: 'todos required (non-empty array)' });
    }
    const next = mergePlanTodos(ctx.plan, todos);
    ctx.setPlan(next);
    return JSON.stringify({ ok: true, plan: next });
  }

  if (!ctx.api) return JSON.stringify({ error: 'editor not ready' });

  if (name === 'get_selection') {
    return JSON.stringify(getSelectionInfo(ctx.api, ctx.docType));
  }

  if (ctx.docType === 'SHEET') {
    const wb = getActiveWorkbook(ctx.api);
    if (!wb) return JSON.stringify({ error: 'no workbook' });
    const sheet = wb.getActiveSheet();

    if (name === 'get_workbook_outline') {
      let usedRange = '';
      try {
        usedRange = sheet.getDataRange().getA1Notation();
      } catch {
        usedRange = '';
      }
      return JSON.stringify({
        sheets: wb.getSheets().map((s) => s.getSheetName()),
        active: sheet.getSheetName(),
        rows: sheet.getMaxRows(),
        cols: sheet.getMaxColumns(),
        usedRange,
        frozenRows: sheet.getFrozenRows?.() ?? 0,
        frozenColumns: sheet.getFrozenColumns?.() ?? 0,
      });
    }
    if (name === 'activate_sheet') {
      try {
        const resolved = resolveSheetByName(wb, args.name, false);
        if ('error' in resolved) return JSON.stringify(resolved);
        if (resolved.sheet.activate) resolved.sheet.activate();
        else if (wb.setActiveSheet) wb.setActiveSheet(resolved.sheet);
        else return JSON.stringify({ error: 'activate_sheet not supported' });
        return JSON.stringify({
          ok: true,
          active: resolved.sheet.getSheetName(),
          sheets: wb.getSheets().map((s) => s.getSheetName()),
        });
      } catch (error) {
        return JSON.stringify({
          error:
            error instanceof Error ? error.message : 'activate_sheet failed',
        });
      }
    }
    if (name === 'insert_sheet') {
      try {
        if (!wb.insertSheet)
          return JSON.stringify({ error: 'insert_sheet not supported' });
        const sheetName =
          args.name != null && String(args.name).trim()
            ? String(args.name).trim()
            : undefined;
        const index =
          args.index != null && Number.isFinite(Number(args.index))
            ? Number(args.index)
            : undefined;
        const created = wb.insertSheet(
          sheetName,
          index != null ? { index } : undefined,
        );
        const active = created ?? wb.getActiveSheet();
        return JSON.stringify({
          ok: true,
          sheet: active.getSheetName(),
          sheets: wb.getSheets().map((s) => s.getSheetName()),
        });
      } catch (error) {
        return JSON.stringify({
          error: error instanceof Error ? error.message : 'insert_sheet failed',
        });
      }
    }
    if (name === 'rename_sheet') {
      try {
        const newName = String(args.newName ?? '').trim();
        if (!newName) return JSON.stringify({ error: 'newName required' });
        const resolved = resolveSheetByName(wb, args.name, true);
        if ('error' in resolved) return JSON.stringify(resolved);
        if (!resolved.sheet.setName) {
          return JSON.stringify({ error: 'rename_sheet not supported' });
        }
        const oldName = resolved.sheet.getSheetName();
        resolved.sheet.setName(newName);
        return JSON.stringify({
          ok: true,
          from: oldName,
          to: resolved.sheet.getSheetName(),
          sheets: wb.getSheets().map((s) => s.getSheetName()),
        });
      } catch (error) {
        return JSON.stringify({
          error: error instanceof Error ? error.message : 'rename_sheet failed',
        });
      }
    }
    if (name === 'delete_sheet') {
      try {
        if (!wb.deleteSheet)
          return JSON.stringify({ error: 'delete_sheet not supported' });
        const resolved = resolveSheetByName(wb, args.name, true);
        if ('error' in resolved) return JSON.stringify(resolved);
        const deletedName = resolved.sheet.getSheetName();
        if (wb.getSheets().length <= 1) {
          return JSON.stringify({ error: 'cannot delete the only sheet' });
        }
        const ok = wb.deleteSheet(resolved.sheet);
        if (!ok) return JSON.stringify({ error: 'delete_sheet failed' });
        return JSON.stringify({
          ok: true,
          deleted: deletedName,
          active: wb.getActiveSheet().getSheetName(),
          sheets: wb.getSheets().map((s) => s.getSheetName()),
        });
      } catch (error) {
        return JSON.stringify({
          error: error instanceof Error ? error.message : 'delete_sheet failed',
        });
      }
    }
    if (name === 'set_freeze') {
      try {
        if (args.clear) {
          if (!sheet.cancelFreeze) {
            return JSON.stringify({ error: 'set_freeze not supported' });
          }
          sheet.cancelFreeze();
          return JSON.stringify({
            ok: true,
            cleared: true,
            sheet: sheet.getSheetName(),
          });
        }
        const rows =
          args.rows != null && Number.isFinite(Number(args.rows))
            ? Math.max(0, Math.floor(Number(args.rows)))
            : undefined;
        const columns =
          args.columns != null && Number.isFinite(Number(args.columns))
            ? Math.max(0, Math.floor(Number(args.columns)))
            : undefined;
        if (rows == null && columns == null) {
          return JSON.stringify({
            error: 'pass rows and/or columns, or clear: true',
          });
        }
        if (rows != null) {
          if (!sheet.setFrozenRows) {
            return JSON.stringify({ error: 'set_freeze not supported' });
          }
          sheet.setFrozenRows(rows);
        }
        if (columns != null) {
          if (!sheet.setFrozenColumns) {
            return JSON.stringify({ error: 'set_freeze not supported' });
          }
          sheet.setFrozenColumns(columns);
        }
        return JSON.stringify({
          ok: true,
          sheet: sheet.getSheetName(),
          frozenRows: sheet.getFrozenRows?.() ?? rows ?? 0,
          frozenColumns: sheet.getFrozenColumns?.() ?? columns ?? 0,
        });
      } catch (error) {
        return JSON.stringify({
          error: error instanceof Error ? error.message : 'set_freeze failed',
        });
      }
    }
    if (name === 'get_filter') {
      try {
        const filter = sheet.getFilter?.() ?? null;
        if (!filter) {
          return JSON.stringify({
            ok: true,
            sheet: sheet.getSheetName(),
            filter: null,
          });
        }
        return JSON.stringify({
          ok: true,
          sheet: sheet.getSheetName(),
          filter: serializeFilter(filter),
        });
      } catch (error) {
        return JSON.stringify({
          error: error instanceof Error ? error.message : 'get_filter failed',
        });
      }
    }
    if (name === 'sort_range') {
      try {
        const resolved = resolveSheetRangeA1(sheet, wb, args.range);
        if ('error' in resolved) return JSON.stringify(resolved);
        const { a1, range } = resolved;
        if (!range.sort) {
          return JSON.stringify({ error: 'sort_range not supported' });
        }
        if (args.by == null) return JSON.stringify({ error: 'by required' });
        const startColumn = getRangeStartColumn(range, a1);
        const by = parseSortBy(args.by, startColumn);
        if (typeof by === 'object' && 'error' in by && !Array.isArray(by)) {
          return JSON.stringify(by);
        }
        range.sort(by as SheetSortBy);
        return JSON.stringify({
          ok: true,
          sheet: sheet.getSheetName(),
          range: range.getA1Notation(),
          by,
        });
      } catch (error) {
        return JSON.stringify({
          error: error instanceof Error ? error.message : 'sort_range failed',
        });
      }
    }
    if (name === 'create_filter') {
      try {
        const resolved = resolveSheetRangeA1(sheet, wb, args.range);
        if ('error' in resolved) return JSON.stringify(resolved);
        const { range } = resolved;
        if (!range.createFilter) {
          return JSON.stringify({ error: 'create_filter not supported' });
        }
        const replace = args.replace !== false;
        let filter = range.createFilter();
        if (!filter && replace) {
          const existing = sheet.getFilter?.() ?? range.getFilter?.();
          existing?.remove?.();
          filter = range.createFilter();
        }
        if (!filter) {
          return JSON.stringify({
            error:
              'create_filter failed; filter may already exist (try replace: true)',
          });
        }
        return JSON.stringify({
          ok: true,
          sheet: sheet.getSheetName(),
          filter: serializeFilter(filter),
        });
      } catch (error) {
        return JSON.stringify({
          error:
            error instanceof Error ? error.message : 'create_filter failed',
        });
      }
    }
    if (name === 'remove_filter') {
      try {
        const filter = sheet.getFilter?.() ?? null;
        if (!filter) {
          return JSON.stringify({
            ok: true,
            sheet: sheet.getSheetName(),
            removed: false,
            reason: 'no filter',
          });
        }
        if (!filter.remove) {
          return JSON.stringify({ error: 'remove_filter not supported' });
        }
        const ok = filter.remove();
        return JSON.stringify({
          ok: true,
          sheet: sheet.getSheetName(),
          removed: ok !== false,
        });
      } catch (error) {
        return JSON.stringify({
          error:
            error instanceof Error ? error.message : 'remove_filter failed',
        });
      }
    }
    if (name === 'set_column_filter') {
      try {
        const filter = sheet.getFilter?.() ?? null;
        if (!filter) {
          return JSON.stringify({
            error: 'no filter; call create_filter first',
          });
        }
        const column = resolveColumnIndex(args.column as string | number);
        if (column == null) {
          return JSON.stringify({ error: 'invalid column' });
        }
        if (args.clear) {
          if (!filter.removeColumnFilterCriteria) {
            return JSON.stringify({ error: 'set_column_filter not supported' });
          }
          filter.removeColumnFilterCriteria(column);
          return JSON.stringify({
            ok: true,
            sheet: sheet.getSheetName(),
            column,
            cleared: true,
            filter: serializeFilter(filter),
          });
        }
        if (!filter.setColumnFilterCriteria) {
          return JSON.stringify({ error: 'set_column_filter not supported' });
        }
        const values = Array.isArray(args.values)
          ? args.values.map((v) => String(v))
          : null;
        if (!values?.length) {
          return JSON.stringify({
            error: 'pass values[] or clear: true',
          });
        }
        filter.setColumnFilterCriteria(column, {
          colId: column,
          filters: { filters: values },
        });
        return JSON.stringify({
          ok: true,
          sheet: sheet.getSheetName(),
          column,
          values,
          filter: serializeFilter(filter),
        });
      } catch (error) {
        return JSON.stringify({
          error:
            error instanceof Error ? error.message : 'set_column_filter failed',
        });
      }
    }
    if (name === 'get_conditional_formatting') {
      try {
        const a1 = args.range != null ? String(args.range).trim() : '';
        let rules: CfRuleLike[] = [];
        if (a1) {
          const range = sheet.getRange(a1);
          if (!range.getConditionalFormattingRules) {
            return JSON.stringify({
              error: 'get_conditional_formatting not supported',
            });
          }
          rules = range.getConditionalFormattingRules() ?? [];
        } else {
          if (!sheet.getConditionalFormattingRules) {
            return JSON.stringify({
              error: 'get_conditional_formatting not supported',
            });
          }
          rules = sheet.getConditionalFormattingRules() ?? [];
        }
        return JSON.stringify({
          ok: true,
          sheet: sheet.getSheetName(),
          range: a1 || null,
          rules: rules.map(summarizeCfRule),
        });
      } catch (error) {
        return JSON.stringify({
          error:
            error instanceof Error
              ? error.message
              : 'get_conditional_formatting failed',
        });
      }
    }
    if (name === 'add_conditional_highlight') {
      try {
        const resolved = resolveSheetRangeA1(sheet, wb, args.range);
        if ('error' in resolved) return JSON.stringify(resolved);
        const { range } = resolved;
        if (
          !sheet.newConditionalFormattingRule ||
          !sheet.addConditionalFormattingRule
        ) {
          return JSON.stringify({
            error: 'add_conditional_highlight not supported',
          });
        }
        const condition = String(args.condition ?? '').trim();
        let builder = sheet.newConditionalFormattingRule();
        if (condition === 'empty') builder = builder.whenCellEmpty();
        else if (condition === 'not_empty')
          builder = builder.whenCellNotEmpty();
        else if (condition === 'number_between') {
          const min = Number(args.min);
          const max = Number(args.max);
          if (!Number.isFinite(min) || !Number.isFinite(max)) {
            return JSON.stringify({
              error: 'number_between requires numeric min and max',
            });
          }
          builder = builder.whenNumberBetween(min, max);
        } else if (condition === 'text_contains') {
          const text = args.text != null ? String(args.text) : '';
          if (!text) {
            return JSON.stringify({
              error: 'text_contains requires text',
            });
          }
          builder = builder.whenTextContains(text);
        } else if (condition === 'duplicate') {
          builder = builder.setDuplicateValues();
        } else {
          return JSON.stringify({
            error:
              'condition must be empty | not_empty | number_between | text_contains | duplicate',
          });
        }
        const bg =
          args.backgroundColor != null
            ? String(args.backgroundColor)
            : '#FFEB9C';
        builder = builder.setBackground(bg);
        if (args.fontColor != null) {
          builder = builder.setFontColor(String(args.fontColor));
        }
        const iRange = range.getRange?.();
        if (!iRange) {
          return JSON.stringify({ error: 'range bounds unavailable' });
        }
        const rule = builder.setRanges([iRange]).build();
        sheet.addConditionalFormattingRule(rule);
        return JSON.stringify({
          ok: true,
          sheet: sheet.getSheetName(),
          range: range.getA1Notation(),
          rule: summarizeCfRule(rule),
        });
      } catch (error) {
        return JSON.stringify({
          error:
            error instanceof Error
              ? error.message
              : 'add_conditional_highlight failed',
        });
      }
    }
    if (name === 'clear_conditional_formatting') {
      try {
        if (args.all) {
          if (!sheet.clearConditionalFormatRules) {
            return JSON.stringify({
              error: 'clear_conditional_formatting not supported',
            });
          }
          sheet.clearConditionalFormatRules();
          return JSON.stringify({
            ok: true,
            sheet: sheet.getSheetName(),
            cleared: 'all',
          });
        }
        if (args.cfId != null) {
          const cfId = String(args.cfId);
          if (!sheet.deleteConditionalFormattingRule) {
            return JSON.stringify({
              error: 'clear_conditional_formatting not supported',
            });
          }
          sheet.deleteConditionalFormattingRule(cfId);
          return JSON.stringify({
            ok: true,
            sheet: sheet.getSheetName(),
            cleared: cfId,
          });
        }
        const resolved = resolveSheetRangeA1(sheet, wb, args.range);
        if ('error' in resolved) return JSON.stringify(resolved);
        const { range } = resolved;
        if (!range.clearConditionalFormatRules) {
          return JSON.stringify({
            error: 'clear_conditional_formatting not supported',
          });
        }
        range.clearConditionalFormatRules();
        return JSON.stringify({
          ok: true,
          sheet: sheet.getSheetName(),
          range: range.getA1Notation(),
          cleared: 'range',
        });
      } catch (error) {
        return JSON.stringify({
          error:
            error instanceof Error
              ? error.message
              : 'clear_conditional_formatting failed',
        });
      }
    }
    if (name === 'get_data_validations') {
      try {
        const a1 = args.range != null ? String(args.range).trim() : '';
        let rules: DataValidationRuleLike[] = [];
        if (a1) {
          const range = sheet.getRange(a1);
          if (!range.getDataValidations) {
            return JSON.stringify({
              error: 'get_data_validations not supported',
            });
          }
          rules = range.getDataValidations() ?? [];
        } else {
          if (!sheet.getDataValidations) {
            return JSON.stringify({
              error: 'get_data_validations not supported',
            });
          }
          rules = sheet.getDataValidations() ?? [];
        }
        return JSON.stringify({
          ok: true,
          sheet: sheet.getSheetName(),
          range: a1 || null,
          rules: rules.map(serializeDataValidation),
        });
      } catch (error) {
        return JSON.stringify({
          error:
            error instanceof Error
              ? error.message
              : 'get_data_validations failed',
        });
      }
    }
    if (name === 'set_list_validation') {
      try {
        const resolved = resolveSheetRangeA1(sheet, wb, args.range);
        if ('error' in resolved) return JSON.stringify(resolved);
        const { range } = resolved;
        if (!ctx.api.newDataValidation || !range.setDataValidation) {
          return JSON.stringify({ error: 'set_list_validation not supported' });
        }
        const values = Array.isArray(args.values)
          ? args.values.map((v) => String(v))
          : [];
        if (!values.length) {
          return JSON.stringify({ error: 'values required (non-empty array)' });
        }
        const showDropdown = args.showDropdown !== false;
        const allowBlank = args.allowBlank !== false;
        const errorMsg = args.error != null ? String(args.error) : undefined;
        const rule = ctx.api
          .newDataValidation()
          .requireValueInList(values, false, showDropdown)
          .setOptions({
            allowBlank,
            showDropDown: showDropdown,
            showErrorMessage: Boolean(errorMsg),
            ...(errorMsg ? { error: errorMsg } : {}),
          })
          .build();
        range.setDataValidation(rule);
        return JSON.stringify({
          ok: true,
          sheet: sheet.getSheetName(),
          range: range.getA1Notation(),
          values,
        });
      } catch (error) {
        return JSON.stringify({
          error:
            error instanceof Error
              ? error.message
              : 'set_list_validation failed',
        });
      }
    }
    if (name === 'set_number_validation') {
      try {
        const resolved = resolveSheetRangeA1(sheet, wb, args.range);
        if ('error' in resolved) return JSON.stringify(resolved);
        const { range } = resolved;
        if (!ctx.api.newDataValidation || !range.setDataValidation) {
          return JSON.stringify({
            error: 'set_number_validation not supported',
          });
        }
        const op = String(args.op ?? '').trim();
        const integer = Boolean(args.integer);
        let builder = ctx.api.newDataValidation();
        if (op === 'between') {
          const min = Number(args.min);
          const max = Number(args.max);
          if (!Number.isFinite(min) || !Number.isFinite(max)) {
            return JSON.stringify({
              error: 'between requires numeric min and max',
            });
          }
          builder = builder.requireNumberBetween(min, max, integer);
        } else if (op === 'greater_than') {
          const value = Number(args.value);
          if (!Number.isFinite(value)) {
            return JSON.stringify({
              error: 'greater_than requires numeric value',
            });
          }
          builder = builder.requireNumberGreaterThan(value, integer);
        } else if (op === 'less_than') {
          const value = Number(args.value);
          if (!Number.isFinite(value)) {
            return JSON.stringify({
              error: 'less_than requires numeric value',
            });
          }
          builder = builder.requireNumberLessThan(value, integer);
        } else {
          return JSON.stringify({
            error: 'op must be between | greater_than | less_than',
          });
        }
        const allowBlank = args.allowBlank !== false;
        const errorMsg = args.error != null ? String(args.error) : undefined;
        const rule = builder
          .setOptions({
            allowBlank,
            showErrorMessage: Boolean(errorMsg),
            ...(errorMsg ? { error: errorMsg } : {}),
          })
          .build();
        range.setDataValidation(rule);
        return JSON.stringify({
          ok: true,
          sheet: sheet.getSheetName(),
          range: range.getA1Notation(),
          op,
          integer,
        });
      } catch (error) {
        return JSON.stringify({
          error:
            error instanceof Error
              ? error.message
              : 'set_number_validation failed',
        });
      }
    }
    if (name === 'clear_data_validation') {
      try {
        const resolved = resolveSheetRangeA1(sheet, wb, args.range);
        if ('error' in resolved) return JSON.stringify(resolved);
        const { range } = resolved;
        if (!range.setDataValidation) {
          return JSON.stringify({
            error: 'clear_data_validation not supported',
          });
        }
        range.setDataValidation(null);
        return JSON.stringify({
          ok: true,
          sheet: sheet.getSheetName(),
          range: range.getA1Notation(),
          cleared: true,
        });
      } catch (error) {
        return JSON.stringify({
          error:
            error instanceof Error
              ? error.message
              : 'clear_data_validation failed',
        });
      }
    }
    if (name === 'read_range') {
      try {
        let a1 = args.range != null ? String(args.range).trim() : '';
        if (!a1) {
          const active = sheet.getActiveRange?.() ?? wb.getActiveRange?.();
          a1 = active?.getA1Notation?.() ?? '';
        }
        if (!a1)
          return JSON.stringify({
            error: 'no range; pass range or select cells',
          });
        const range = sheet.getRange(a1);
        const raw = range.getValues();
        const { values, truncated, rows, cols } = truncateMatrix(
          raw as unknown[][],
        );
        return JSON.stringify({
          sheet: sheet.getSheetName(),
          range: range.getA1Notation(),
          values,
          truncated,
          size: { rows, cols },
        });
      } catch (error) {
        return JSON.stringify({
          error: error instanceof Error ? error.message : 'read_range failed',
        });
      }
    }
    if (name === 'read_formulas') {
      try {
        let a1 = args.range != null ? String(args.range).trim() : '';
        if (!a1) {
          const active = sheet.getActiveRange?.() ?? wb.getActiveRange?.();
          a1 = active?.getA1Notation?.() ?? '';
        }
        if (!a1)
          return JSON.stringify({
            error: 'no range; pass range or select cells',
          });
        const range = sheet.getRange(a1);
        const rawFormulas = range.getFormulas();
        const rawValues = range.getValues();
        const formulasTrunc = truncateMatrix(rawFormulas as unknown[][]);
        const valuesTrunc = truncateMatrix(rawValues as unknown[][]);
        return JSON.stringify({
          sheet: sheet.getSheetName(),
          range: range.getA1Notation(),
          formulas: formulasTrunc.values,
          values: valuesTrunc.values,
          truncated: formulasTrunc.truncated || valuesTrunc.truncated,
          size: {
            rows: Math.max(formulasTrunc.rows, valuesTrunc.rows),
            cols: Math.max(formulasTrunc.cols, valuesTrunc.cols),
          },
        });
      } catch (error) {
        return JSON.stringify({
          error:
            error instanceof Error ? error.message : 'read_formulas failed',
        });
      }
    }
    if (name === 'search_in_sheet') {
      try {
        const q = String(args.query || '').toLowerCase();
        if (!q) return JSON.stringify({ error: 'empty query' });
        let searchRange = 'A1:Z200';
        let originRow = 0;
        let originCol = 0;
        try {
          const data = sheet.getDataRange();
          searchRange = data.getA1Notation();
          const bounds = data.getRange?.();
          if (bounds) {
            originRow = bounds.startRow;
            originCol = bounds.startColumn;
          } else {
            const startCell = searchRange.split(':')[0];
            const parsed = parseA1Cell(startCell);
            if (parsed) {
              originRow = parsed.row;
              originCol = parsed.col;
            }
          }
        } catch {
          /* fallback A1:Z200 */
        }
        const values = sheet.getRange(searchRange).getValues() as unknown[][];
        const hits: Array<{ a1: string; v: unknown }> = [];
        values.forEach((row, r) =>
          row.forEach((v, c) => {
            if (
              String(v ?? '')
                .toLowerCase()
                .includes(q)
            ) {
              hits.push({ a1: a1FromRowCol(originRow + r, originCol + c), v });
            }
          }),
        );
        return JSON.stringify({
          sheet: sheet.getSheetName(),
          range: searchRange,
          hits: hits.slice(0, SEARCH_MAX_HITS),
          truncated: hits.length > SEARCH_MAX_HITS,
        });
      } catch (error) {
        return JSON.stringify({
          error:
            error instanceof Error ? error.message : 'search_in_sheet failed',
        });
      }
    }
    if (name === 'replace_in_sheet') {
      try {
        const search = String(args.search ?? '');
        if (!search) return JSON.stringify({ error: 'empty search' });
        const replacement =
          args.replacement != null ? String(args.replacement) : '';
        const q = search.toLowerCase();

        let searchRange = 'A1:Z200';
        let originRow = 0;
        let originCol = 0;
        const rangeArg = args.range != null ? String(args.range).trim() : '';
        if (rangeArg) {
          searchRange = rangeArg;
          const startCell = searchRange.replace(/^[^!]*!/, '').split(':')[0];
          const parsed = parseA1Cell(
            /^\$?[A-Za-z]+$/.test(startCell) ? `${startCell}1` : startCell,
          );
          if (parsed) {
            originRow = parsed.row;
            originCol = parsed.col;
          }
          try {
            const bounds = sheet.getRange(searchRange).getRange?.();
            if (bounds) {
              originRow = bounds.startRow;
              originCol = bounds.startColumn;
            }
          } catch {
            /* keep parsed origin */
          }
        } else {
          try {
            const data = sheet.getDataRange();
            searchRange = data.getA1Notation();
            const bounds = data.getRange?.();
            if (bounds) {
              originRow = bounds.startRow;
              originCol = bounds.startColumn;
            } else {
              const startCell = searchRange.split(':')[0];
              const parsed = parseA1Cell(startCell);
              if (parsed) {
                originRow = parsed.row;
                originCol = parsed.col;
              }
            }
          } catch {
            /* fallback A1:Z200 */
          }
        }

        const range = sheet.getRange(searchRange);
        const values = range.getValues() as unknown[][];
        const formulas = (range.getFormulas?.() ?? []) as string[][];
        const hits: Array<{ a1: string; from: string; to: string }> = [];

        values.forEach((row, r) =>
          row.forEach((v, c) => {
            const formula = String(formulas[r]?.[c] ?? '').trim();
            if (formula) return;
            const from = String(v ?? '');
            if (!from.toLowerCase().includes(q)) return;
            const to = replaceAllCaseInsensitive(from, search, replacement);
            if (to === from) return;
            hits.push({
              a1: a1FromRowCol(originRow + r, originCol + c),
              from,
              to,
            });
          }),
        );

        const truncated = hits.length > SEARCH_MAX_HITS;
        const toWrite = hits.slice(0, SEARCH_MAX_HITS);
        toWrite.forEach((cell) => {
          sheet.getRange(cell.a1).setValues([[cell.to]]);
        });
        return JSON.stringify({
          ok: true,
          sheet: sheet.getSheetName(),
          range: searchRange,
          replaced: toWrite.length,
          cells: toWrite,
          truncated,
        });
      } catch (error) {
        return JSON.stringify({
          error:
            error instanceof Error ? error.message : 'replace_in_sheet failed',
        });
      }
    }
    if (name === 'set_cells') {
      try {
        const values = args.values as unknown[][];
        if (!Array.isArray(values) || !values.length) {
          return JSON.stringify({
            error: 'values must be a non-empty 2D array',
          });
        }
        const numRows = values.length;
        const numCols = Math.max(
          ...values.map((row) => (Array.isArray(row) ? row.length : 0)),
          0,
        );
        if (!numCols) return JSON.stringify({ error: 'values rows are empty' });

        let startRow = 0;
        let startCol = 0;
        if (args.start != null && String(args.start).trim()) {
          const parsed = parseA1Cell(String(args.start));
          if (!parsed || parsed.row < 0 || parsed.col < 0) {
            return JSON.stringify({
              error: `invalid start cell: ${String(args.start)}`,
            });
          }
          startRow = parsed.row;
          startCol = parsed.col;
        } else {
          const cell = sheet.getActiveCell?.() ?? wb.getActiveCell?.();
          const a1 = cell?.getA1Notation?.();
          const parsed = a1 ? parseA1Cell(a1) : null;
          if (!parsed) {
            return JSON.stringify({ error: 'no start and no active cell' });
          }
          startRow = parsed.row;
          startCol = parsed.col;
        }

        const normalized = values.map((row) => {
          const r = Array.isArray(row) ? [...row] : [];
          while (r.length < numCols) r.push('');
          return r.slice(0, numCols);
        });
        const written = a1Range(startRow, startCol, numRows, numCols);
        sheet
          .getRange(startRow, startCol, numRows, numCols)
          .setValues(normalized);
        return JSON.stringify({
          ok: true,
          sheet: sheet.getSheetName(),
          range: written,
          cells: numRows * numCols,
        });
      } catch (error) {
        return JSON.stringify({
          error: error instanceof Error ? error.message : 'set_cells failed',
        });
      }
    }
    if (name === 'set_formulas') {
      try {
        const formulas = args.formulas as unknown[][];
        if (!Array.isArray(formulas) || !formulas.length) {
          return JSON.stringify({
            error: 'formulas must be a non-empty 2D array',
          });
        }
        const numRows = formulas.length;
        const numCols = Math.max(
          ...formulas.map((row) => (Array.isArray(row) ? row.length : 0)),
          0,
        );
        if (!numCols)
          return JSON.stringify({ error: 'formulas rows are empty' });

        let startRow = 0;
        let startCol = 0;
        if (args.start != null && String(args.start).trim()) {
          const parsed = parseA1Cell(String(args.start));
          if (!parsed || parsed.row < 0 || parsed.col < 0) {
            return JSON.stringify({
              error: `invalid start cell: ${String(args.start)}`,
            });
          }
          startRow = parsed.row;
          startCol = parsed.col;
        } else {
          const cell = sheet.getActiveCell?.() ?? wb.getActiveCell?.();
          const a1 = cell?.getA1Notation?.();
          const parsed = a1 ? parseA1Cell(a1) : null;
          if (!parsed) {
            return JSON.stringify({ error: 'no start and no active cell' });
          }
          startRow = parsed.row;
          startCol = parsed.col;
        }

        const cells = collectFormulaCells(formulas, startRow, startCol);
        if (!cells.length) {
          return JSON.stringify({
            error: 'no valid formulas in matrix',
            hint: 'Передайте непустые формулы; пустые ячейки можно опустить или вызвать set_formulas только для заполненных позиций.',
          });
        }

        const denseOk = cells.length === numRows * numCols;
        const written = a1Range(startRow, startCol, numRows, numCols);
        try {
          if (denseOk) {
            const dense: string[][] = Array.from(
              { length: numRows },
              (unusedRow, r) => {
                void unusedRow;
                return Array.from({ length: numCols }, (unusedCol, c) => {
                  void unusedCol;
                  const hit = cells.find(
                    (cell) =>
                      cell.row === startRow + r && cell.col === startCol + c,
                  );
                  return hit?.formula ?? '';
                });
              },
            );
            sheet
              .getRange(startRow, startCol, numRows, numCols)
              .setFormulas(dense);
          } else {
            cells.forEach((cell) => {
              sheet
                .getRange(cell.row, cell.col, 1, 1)
                .setFormulas([[cell.formula]]);
            });
          }
        } catch (writeError) {
          const failed = cells.find((cell) => {
            try {
              sheet
                .getRange(cell.row, cell.col, 1, 1)
                .setFormulas([[cell.formula]]);
              return false;
            } catch {
              return true;
            }
          });
          if (failed) {
            return JSON.stringify({
              error:
                writeError instanceof Error
                  ? writeError.message
                  : 'set_formulas failed',
              hint: `Проблемная ячейка ${a1FromRowCol(failed.row, failed.col)}: ${failed.formula}. Используйте EN-имена функций и запятые в аргументах.`,
              range: written,
            });
          }
        }
        return JSON.stringify({
          ok: true,
          sheet: sheet.getSheetName(),
          range: written,
          cells: cells.length,
          skippedEmpty: numRows * numCols - cells.length,
        });
      } catch (error) {
        return JSON.stringify({
          error: error instanceof Error ? error.message : 'set_formulas failed',
          hint: 'Проверьте start (A1) и формулу: ведущий =, EN-имена (SUM/IF/EDATE), аргументы через запятую.',
        });
      }
    }
    if (name === 'copy_formulas') {
      try {
        const source = String(args.source || '').trim();
        const destination = String(args.destination || '').trim();
        if (!source) return JSON.stringify({ error: 'source required' });
        if (!destination)
          return JSON.stringify({ error: 'destination required' });
        const times = Math.max(
          1,
          Math.min(200, Math.floor(Number(args.times) || 1)),
        );
        const srcRange = sheet.getRange(source);
        const srcFormulas = (srcRange.getFormulas?.() ?? []) as string[][];
        const srcValues = (srcRange.getValues?.() ?? []) as unknown[][];
        const numRows = Math.max(srcFormulas.length, srcValues.length);
        const numCols = Math.max(
          ...srcFormulas.map((r) => r?.length ?? 0),
          ...srcValues.map((r) => (Array.isArray(r) ? r.length : 0)),
          0,
        );
        if (!numRows || !numCols) {
          return JSON.stringify({ error: 'source range is empty' });
        }
        const destParsed = parseA1Cell(
          destination.includes(':') ? destination.split(':')[0] : destination,
        );
        if (!destParsed) {
          return JSON.stringify({
            error: `invalid destination: ${destination}`,
          });
        }
        const srcA1 = srcRange.getA1Notation?.() || source;
        const srcStart = parseA1Cell(
          srcA1.replace(/^[^!]*!/, '').split(':')[0],
        );
        const srcStartRow = srcStart?.row ?? destParsed.row;
        const srcStartCol = srcStart?.col ?? destParsed.col;

        let writtenCells = 0;
        Array.from({ length: times }).forEach((unusedT, t) => {
          void unusedT;
          const baseRow = destParsed.row + t * numRows;
          const baseCol = destParsed.col;
          const dRow = baseRow - srcStartRow;
          const dCol = baseCol - srcStartCol;
          const out: string[][] = Array.from(
            { length: numRows },
            (unusedR, r) => {
              void unusedR;
              return Array.from({ length: numCols }, (unusedC, c) => {
                void unusedC;
                const raw = String(srcFormulas[r]?.[c] ?? '').trim();
                if (!raw || raw === '=') return '';
                const withEq = raw.startsWith('=') ? raw : `=${raw}`;
                writtenCells += 1;
                return adjustFormulaRefs(withEq, dRow, dCol);
              });
            },
          );
          // Write non-empty cells only
          out.forEach((rowOut, r) => {
            rowOut.forEach((f, c) => {
              if (!f) return;
              sheet.getRange(baseRow + r, baseCol + c, 1, 1).setFormulas([[f]]);
            });
          });
        });
        if (!writtenCells) {
          return JSON.stringify({
            error: 'no formulas in source range',
            hint: 'Убедитесь, что source содержит формулы (не только значения).',
          });
        }
        return JSON.stringify({
          ok: true,
          sheet: sheet.getSheetName(),
          source: srcA1,
          destination,
          times,
          cells: writtenCells,
        });
      } catch (error) {
        return JSON.stringify({
          error:
            error instanceof Error ? error.message : 'copy_formulas failed',
        });
      }
    }
    if (name === 'insert_image') {
      try {
        const url = String(args.url || '').trim();
        if (!url) return JSON.stringify({ error: 'url required' });
        if (!/^https?:\/\//i.test(url)) {
          return JSON.stringify({ error: 'url must be http(s)' });
        }
        let col = 0;
        let row = 0;
        if (args.start != null && String(args.start).trim()) {
          const parsed = parseA1Cell(String(args.start));
          if (!parsed) {
            return JSON.stringify({
              error: `invalid start cell: ${String(args.start)}`,
            });
          }
          col = parsed.col;
          row = parsed.row;
        }
        const insert = (
          sheet as SheetHandle & {
            insertImage?: (
              url: string,
              column?: number,
              row?: number,
            ) => Promise<boolean>;
          }
        ).insertImage;
        if (typeof insert !== 'function') {
          return JSON.stringify({
            error: 'insertImage not supported',
            hint: 'Drawing preset may be unavailable; insert the image manually.',
          });
        }
        const ok = await insert.call(sheet, url, col, row);
        if (!ok) {
          return JSON.stringify({ error: 'insert_image failed' });
        }
        return JSON.stringify({
          ok: true,
          sheet: sheet.getSheetName(),
          url,
          start: a1FromRowCol(row, col),
        });
      } catch (error) {
        return JSON.stringify({
          error: error instanceof Error ? error.message : 'insert_image failed',
        });
      }
    }
    if (name === 'clear_range') {
      try {
        const a1 = String(args.range || '');
        if (!a1) return JSON.stringify({ error: 'range required' });
        sheet.getRange(a1).clear();
        return JSON.stringify({
          ok: true,
          range: a1,
          sheet: sheet.getSheetName(),
        });
      } catch (error) {
        return JSON.stringify({
          error: error instanceof Error ? error.message : 'clear_range failed',
        });
      }
    }
    if (name === 'format_range') {
      try {
        const resolved = resolveSheetRangeA1(sheet, wb, args.range);
        if ('error' in resolved)
          return JSON.stringify({ error: resolved.error });
        const { range } = resolved;
        const applied: string[] = [];
        if (typeof args.bold === 'boolean') {
          if (typeof range.setFontWeight !== 'function') {
            return JSON.stringify({ error: 'setFontWeight unavailable' });
          }
          range.setFontWeight(args.bold ? 'bold' : 'normal');
          applied.push(args.bold ? 'bold' : 'normal');
        }
        if (args.fontSize != null) {
          if (typeof range.setFontSize !== 'function') {
            return JSON.stringify({ error: 'setFontSize unavailable' });
          }
          range.setFontSize(Number(args.fontSize));
          applied.push(`fontSize:${Number(args.fontSize)}`);
        }
        if (args.fontColor != null) {
          if (typeof range.setFontColor !== 'function') {
            return JSON.stringify({ error: 'setFontColor unavailable' });
          }
          range.setFontColor(String(args.fontColor));
          applied.push(`fontColor:${String(args.fontColor)}`);
        }
        if (args.backgroundColor != null) {
          if (typeof range.setBackgroundColor !== 'function') {
            return JSON.stringify({ error: 'setBackgroundColor unavailable' });
          }
          range.setBackgroundColor(String(args.backgroundColor));
          applied.push(`backgroundColor:${String(args.backgroundColor)}`);
        }
        if (args.horizontalAlign != null) {
          const align = mapHorizontalAlign(String(args.horizontalAlign));
          if (!align) {
            return JSON.stringify({
              error: 'horizontalAlign must be left|center|right',
            });
          }
          if (typeof range.setHorizontalAlignment !== 'function') {
            return JSON.stringify({
              error: 'setHorizontalAlignment unavailable',
            });
          }
          range.setHorizontalAlignment(align);
          applied.push(`align:${String(args.horizontalAlign)}`);
        }
        if (typeof args.wrap === 'boolean') {
          if (typeof range.setWrap !== 'function') {
            return JSON.stringify({ error: 'setWrap unavailable' });
          }
          range.setWrap(args.wrap);
          applied.push(`wrap:${args.wrap}`);
        }
        if (!applied.length) {
          return JSON.stringify({ error: 'no style properties provided' });
        }
        return JSON.stringify({
          ok: true,
          sheet: sheet.getSheetName(),
          range: range.getA1Notation(),
          applied,
        });
      } catch (error) {
        return JSON.stringify({
          error: error instanceof Error ? error.message : 'format_range failed',
        });
      }
    }
    if (name === 'set_column_widths') {
      try {
        const normalizePxWidth = (raw: number) => {
          // Model often sends Excel-like character widths (8–32); Univer expects pixels.
          let width = raw;
          if (width > 0 && width < 40) width *= 8;
          return Math.max(40, Math.round(width));
        };
        const applied: Array<{ column: string; index: number; width: number }> =
          [];
        const widths = args.widths as
          Array<{ column: string | number; width: number }> | undefined;
        if (Array.isArray(widths) && widths.length) {
          const invalid = widths.find((item) => {
            const index = resolveColumnIndex(item.column);
            const rawWidth = Number(item.width);
            return index == null || !Number.isFinite(rawWidth) || rawWidth <= 0;
          });
          if (invalid) {
            return JSON.stringify({
              error: `invalid column/width: ${JSON.stringify(invalid)}`,
            });
          }
          const { setColumnWidth } = sheet;
          if (typeof setColumnWidth !== 'function') {
            return JSON.stringify({ error: 'setColumnWidth unavailable' });
          }
          widths.forEach((item) => {
            const index = resolveColumnIndex(item.column)!;
            const width = normalizePxWidth(Number(item.width));
            setColumnWidth.call(sheet, index, width);
            applied.push({ column: indexToColLetters(index), index, width });
          });
        } else if (
          args.startColumn != null &&
          args.numColumns != null &&
          args.width != null
        ) {
          const start = resolveColumnIndex(args.startColumn as string | number);
          const num = Number(args.numColumns);
          const rawWidth = Number(args.width);
          if (
            start == null ||
            !Number.isFinite(num) ||
            num < 1 ||
            !Number.isFinite(rawWidth) ||
            rawWidth <= 0
          ) {
            return JSON.stringify({
              error: 'invalid startColumn/numColumns/width',
            });
          }
          const width = normalizePxWidth(rawWidth);
          const { setColumnWidths, setColumnWidth } = sheet;
          if (typeof setColumnWidths === 'function') {
            setColumnWidths.call(sheet, start, num, width);
          } else if (typeof setColumnWidth === 'function') {
            Array.from({ length: num }).forEach((unused, i) => {
              void unused;
              setColumnWidth.call(sheet, start + i, width);
            });
          } else {
            return JSON.stringify({ error: 'setColumnWidth unavailable' });
          }
          Array.from({ length: num }).forEach((unused, i) => {
            void unused;
            applied.push({
              column: indexToColLetters(start + i),
              index: start + i,
              width,
            });
          });
        } else {
          return JSON.stringify({
            error: 'provide widths[] or startColumn+numColumns+width',
          });
        }
        return JSON.stringify({
          ok: true,
          sheet: sheet.getSheetName(),
          columns: applied,
        });
      } catch (error) {
        return JSON.stringify({
          error:
            error instanceof Error ? error.message : 'set_column_widths failed',
        });
      }
    }
    if (name === 'set_number_format') {
      try {
        const pattern = String(args.pattern ?? '');
        if (!pattern) return JSON.stringify({ error: 'pattern required' });
        const resolved = resolveSheetRangeA1(sheet, wb, args.range);
        if ('error' in resolved)
          return JSON.stringify({ error: resolved.error });
        const { range } = resolved;
        if (typeof range.setNumberFormat !== 'function') {
          return JSON.stringify({ error: 'setNumberFormat unavailable' });
        }
        range.setNumberFormat(pattern);
        return JSON.stringify({
          ok: true,
          sheet: sheet.getSheetName(),
          range: range.getA1Notation(),
          pattern,
        });
      } catch (error) {
        return JSON.stringify({
          error:
            error instanceof Error ? error.message : 'set_number_format failed',
        });
      }
    }
    if (name === 'merge_range') {
      try {
        const mode = String(args.mode || 'merge');
        const resolved = resolveSheetRangeA1(sheet, wb, args.range);
        if ('error' in resolved)
          return JSON.stringify({ error: resolved.error });
        const { range } = resolved;
        if (mode === 'unmerge') {
          if (typeof range.breakApart !== 'function') {
            return JSON.stringify({ error: 'breakApart unavailable' });
          }
          range.breakApart();
        } else if (mode === 'across') {
          if (typeof range.mergeAcross !== 'function') {
            return JSON.stringify({ error: 'mergeAcross unavailable' });
          }
          range.mergeAcross({ isForceMerge: true });
        } else if (mode === 'vertically') {
          if (typeof range.mergeVertically !== 'function') {
            return JSON.stringify({ error: 'mergeVertically unavailable' });
          }
          range.mergeVertically({ isForceMerge: true });
        } else if (mode === 'merge') {
          if (typeof range.merge !== 'function') {
            return JSON.stringify({ error: 'merge unavailable' });
          }
          range.merge({ isForceMerge: true });
        } else {
          return JSON.stringify({
            error: 'mode must be merge|across|vertically|unmerge',
          });
        }
        return JSON.stringify({
          ok: true,
          sheet: sheet.getSheetName(),
          range: range.getA1Notation(),
          mode,
        });
      } catch (error) {
        return JSON.stringify({
          error: error instanceof Error ? error.message : 'merge_range failed',
        });
      }
    }
    if (name === 'set_borders') {
      try {
        const type = mapBorderType(String(args.type || ''));
        if (type == null) {
          return JSON.stringify({
            error:
              'type must be all|outside|inside|top|bottom|left|right|horizontal|vertical|none',
          });
        }
        const styleName = String(args.style || 'thin');
        const style = mapBorderStyle(styleName);
        if (style == null) {
          return JSON.stringify({
            error: 'style must be thin|medium|thick|dashed|dotted|double|none',
          });
        }
        const color = args.color != null ? String(args.color) : '#000000';
        const resolved = resolveSheetRangeA1(sheet, wb, args.range);
        if ('error' in resolved)
          return JSON.stringify({ error: resolved.error });
        const { range } = resolved;
        if (typeof range.setBorder !== 'function') {
          return JSON.stringify({ error: 'setBorder unavailable' });
        }
        range.setBorder(type, style, color);
        return JSON.stringify({
          ok: true,
          sheet: sheet.getSheetName(),
          range: range.getA1Notation(),
          type: String(args.type),
          style: styleName,
          color,
        });
      } catch (error) {
        return JSON.stringify({
          error: error instanceof Error ? error.message : 'set_borders failed',
        });
      }
    }
    if (name === 'set_row_heights') {
      try {
        const applied: Array<{ row: number; index: number; height: number }> =
          [];
        const heights = args.heights as
          Array<{ row: number; height: number }> | undefined;
        if (Array.isArray(heights) && heights.length) {
          const invalid = heights.find((item) => {
            const row1 = Number(item.row);
            const height = Number(item.height);
            return (
              !Number.isFinite(row1) ||
              row1 < 1 ||
              !Number.isFinite(height) ||
              height <= 0
            );
          });
          if (invalid) {
            return JSON.stringify({
              error: `invalid row/height: ${JSON.stringify(invalid)}`,
            });
          }
          const { setRowHeight } = sheet;
          if (typeof setRowHeight !== 'function') {
            return JSON.stringify({ error: 'setRowHeight unavailable' });
          }
          heights.forEach((item) => {
            const row1 = Number(item.row);
            const height = Number(item.height);
            const index = Math.floor(row1) - 1;
            setRowHeight.call(sheet, index, height);
            applied.push({ row: Math.floor(row1), index, height });
          });
        } else if (
          args.startRow != null &&
          args.numRows != null &&
          args.height != null
        ) {
          const start1 = Number(args.startRow);
          const num = Number(args.numRows);
          const height = Number(args.height);
          if (
            !Number.isFinite(start1) ||
            start1 < 1 ||
            !Number.isFinite(num) ||
            num < 1 ||
            !Number.isFinite(height) ||
            height <= 0
          ) {
            return JSON.stringify({ error: 'invalid startRow/numRows/height' });
          }
          const start = Math.floor(start1) - 1;
          const { setRowHeights, setRowHeight } = sheet;
          if (typeof setRowHeights === 'function') {
            setRowHeights.call(sheet, start, num, height);
          } else if (typeof setRowHeight === 'function') {
            Array.from({ length: num }).forEach((unused, i) => {
              void unused;
              setRowHeight.call(sheet, start + i, height);
            });
          } else {
            return JSON.stringify({ error: 'setRowHeight unavailable' });
          }
          Array.from({ length: num }).forEach((unused, i) => {
            void unused;
            applied.push({ row: start + i + 1, index: start + i, height });
          });
        } else {
          return JSON.stringify({
            error: 'provide heights[] or startRow+numRows+height',
          });
        }
        return JSON.stringify({
          ok: true,
          sheet: sheet.getSheetName(),
          rows: applied,
        });
      } catch (error) {
        return JSON.stringify({
          error:
            error instanceof Error ? error.message : 'set_row_heights failed',
        });
      }
    }
  }

  if (ctx.docType === 'DOC') {
    const doc = ctx.api.getActiveDocument();
    if (!doc) return JSON.stringify({ error: 'no document' });
    const stream =
      doc.getBody?.()?.dataStream ??
      (typeof doc.getSnapshot === 'function'
        ? doc.getSnapshot()?.body?.dataStream
        : undefined) ??
      doc.save()?.body?.dataStream ??
      '';

    if (name === 'get_doc_outline') {
      try {
        const maxItems = Math.min(
          200,
          Math.max(1, Number(args.maxItems ?? 80) || 80),
        );
        const paragraphs =
          typeof doc.getParagraphs === 'function' ? doc.getParagraphs() : [];
        const mapped = paragraphs.map((p, i) => {
          const info = p.getInfo?.();
          const named = info?.paragraph?.paragraphStyle?.namedStyleType as
            NamedStyleType | undefined;
          const styleLabel =
            named != null
              ? (HEADING_STYLE_LABEL[named] ?? String(named))
              : 'normal';
          return {
            index: info?.paragraphIndex ?? i,
            id: p.getId?.() ?? '',
            text: p.getText().slice(0, 120),
            namedStyle: styleLabel,
            isHeading: isHeadingStyle(named),
            isList: Boolean(p.isListItem?.()),
            startOffset: info?.startOffset ?? null,
            endOffset: info?.endOffset ?? null,
          };
        });
        const headings = mapped.filter((item) => item.isHeading);
        const items = (headings.length ? headings : mapped).slice(0, maxItems);
        return JSON.stringify({
          count: items.length,
          totalParagraphs: mapped.length,
          headingsOnly: headings.length > 0,
          items,
        });
      } catch (error) {
        return JSON.stringify({
          error:
            error instanceof Error ? error.message : 'get_doc_outline failed',
        });
      }
    }
    if (name === 'read_document') {
      const max = Math.max(1, Number(args.maxChars ?? 4000) || 4000);
      const start = Math.max(0, Number(args.startOffset ?? 0) || 0);
      const text = stream.slice(start, start + max);
      return JSON.stringify({
        length: stream.length,
        startOffset: start,
        text,
        truncated: start + text.length < stream.length,
      });
    }
    if (name === 'search_in_document') {
      try {
        const query = String(args.query ?? '');
        if (!query) return JSON.stringify({ error: 'empty query' });
        const matchCase = Boolean(args.matchCase);
        const maxHits = Math.min(
          100,
          Math.max(
            1,
            Number(args.maxHits ?? SEARCH_MAX_HITS) || SEARCH_MAX_HITS,
          ),
        );
        const paragraphs =
          typeof doc.getParagraphs === 'function' ? doc.getParagraphs() : [];
        const hits: Array<{
          paragraphIndex: number;
          paragraphId: string;
          startOffset: number | null;
          endOffset: number | null;
          preview: string;
        }> = [];
        paragraphs.some((p, i) => {
          if (typeof p.findAllText !== 'function') return false;
          const ranges = p.findAllText(query, { matchCase });
          const info = p.getInfo?.();
          ranges.some((range) => {
            const desc = range.describe?.();
            hits.push({
              paragraphIndex: info?.paragraphIndex ?? i,
              paragraphId: p.getId?.() ?? '',
              startOffset: desc?.startOffset ?? null,
              endOffset: desc?.endOffset ?? null,
              preview: (range.getText?.() ?? p.getText()).slice(0, 80),
            });
            return hits.length >= maxHits;
          });
          return hits.length >= maxHits;
        });
        return JSON.stringify({
          query,
          hits,
          truncated: hits.length >= maxHits,
        });
      } catch (error) {
        return JSON.stringify({
          error:
            error instanceof Error
              ? error.message
              : 'search_in_document failed',
        });
      }
    }
    if (name === 'insert_paragraph') {
      try {
        const text = String(args.text ?? '');
        if (args.index != null && Number.isInteger(Number(args.index))) {
          const index = Number(args.index);
          if (!doc.insertParagraph) {
            return JSON.stringify({ error: 'insertParagraph unavailable' });
          }
          const p = doc.insertParagraph(index, text);
          const info = p.getInfo?.();
          return JSON.stringify({
            ok: true,
            index: info?.paragraphIndex ?? index,
            text: p.getText().slice(0, 120),
          });
        }
        if (typeof doc.appendParagraph !== 'function') {
          return JSON.stringify({ error: 'appendParagraph unavailable' });
        }
        const p = doc.appendParagraph(text);
        const info = p.getInfo?.();
        return JSON.stringify({
          ok: true,
          index: info?.paragraphIndex ?? null,
          text: p.getText().slice(0, 120),
          appended: true,
        });
      } catch (error) {
        return JSON.stringify({
          error:
            error instanceof Error ? error.message : 'insert_paragraph failed',
        });
      }
    }
    if (name === 'set_heading') {
      try {
        const styleKey = String(args.style ?? '').toLowerCase();
        const named = HEADING_STYLE_MAP[styleKey];
        if (named == null) {
          return JSON.stringify({
            error: 'style must be normal|title|subtitle|h1|h2|h3|h4|h5',
          });
        }
        const resolved = resolveDocParagraph(doc, args.paragraphIndex);
        if ('error' in resolved) return JSON.stringify(resolved);
        const info = resolved.paragraph.getInfo?.();
        if (info && doc.setSelection) {
          doc.setSelection(
            info.startOffset,
            Math.max(info.startOffset, info.endOffset),
          );
        }
        let ok = false;
        if (typeof resolved.paragraph.setStyle === 'function') {
          ok = resolved.paragraph.setStyle({ namedStyleType: named });
        }
        if (!ok) {
          ok = await executeDocCommand(SetParagraphNamedStyleCommand.id, {
            value: named,
          });
        }
        if (!ok) return JSON.stringify({ error: 'set_heading failed' });
        return JSON.stringify({
          ok: true,
          style: styleKey,
          paragraphIndex: resolved.index,
        });
      } catch (error) {
        return JSON.stringify({
          error: error instanceof Error ? error.message : 'set_heading failed',
        });
      }
    }
    if (name === 'set_list') {
      try {
        const listType = String(args.type ?? '').toLowerCase();
        if (!['bullet', 'ordered', 'check', 'none'].includes(listType)) {
          return JSON.stringify({
            error: 'type must be bullet|ordered|check|none',
          });
        }
        const resolved = resolveDocParagraph(doc, args.paragraphIndex);
        if ('error' in resolved) return JSON.stringify(resolved);
        const info = resolved.paragraph.getInfo?.();
        if (info && doc.setSelection) {
          doc.setSelection(
            info.startOffset,
            Math.max(info.startOffset, info.endOffset),
          );
        }
        let ok = false;
        if (listType === 'none') {
          if (resolved.paragraph.isListItem?.()) {
            ok = await executeDocCommand(BulletListCommand.id);
          } else {
            ok = true;
          }
        } else if (listType === 'bullet') {
          ok = await executeDocCommand(BulletListCommand.id);
        } else if (listType === 'ordered') {
          ok = await executeDocCommand(OrderListCommand.id);
        } else if (listType === 'check') {
          ok =
            (await executeDocCommand(CheckListCommand.id)) ||
            (await executeDocCommand(ListOperationCommand.id, {
              listType: PresetListType.CHECK_LIST,
            }));
        }
        if (!ok) {
          return JSON.stringify({
            error: 'set_list failed',
            hint: 'Укажите paragraphIndex или поставьте курсор в абзац; type: bullet|ordered|check|none.',
          });
        }
        return JSON.stringify({
          ok: true,
          type: listType,
          paragraphIndex: resolved.index,
          isList: Boolean(resolved.paragraph.isListItem?.()),
        });
      } catch (error) {
        return JSON.stringify({
          error: error instanceof Error ? error.message : 'set_list failed',
          hint: 'Попробуйте insert_paragraph + set_list на новом абзаце.',
        });
      }
    }
    if (name === 'insert_text') {
      const text = String(args.text ?? '');
      if (!text) return JSON.stringify({ error: 'empty text' });
      const at = String(args.at || 'end') === 'selection' ? 'selection' : 'end';
      try {
        if (at === 'selection') {
          const offsets = getDocActiveOffsets();
          if (!offsets) return JSON.stringify({ error: 'no active selection' });
          const { startOffset, endOffset } = offsets;
          if (
            startOffset !== endOffset &&
            typeof doc.deleteRange === 'function'
          ) {
            doc.deleteRange({ startOffset, endOffset });
          }
          if (typeof doc.insertText !== 'function') {
            return JSON.stringify({ error: 'insertText unavailable' });
          }
          const ok = doc.insertText(startOffset, text);
          if (!ok) return JSON.stringify({ error: 'insertText failed' });
          return JSON.stringify({
            ok: true,
            at: 'selection',
            startOffset,
            endOffset: startOffset + text.length,
          });
        }
        const lines = text
          .replace(/\r\n/g, '\n')
          .replace(/\r/g, '\n')
          .split('\n');
        if (typeof doc.appendParagraph === 'function') {
          lines.forEach((line) => {
            doc.appendParagraph(line);
          });
          return JSON.stringify({
            ok: true,
            at: 'end',
            paragraphs: lines.length,
          });
        }
        if (typeof doc.insertText === 'function') {
          let index = stream.length;
          if (stream.endsWith('\n')) index -= 1;
          if (index > 0 && stream[index - 1] === '\r') index -= 1;
          const ok = doc.insertText(Math.max(0, index), text);
          if (!ok) return JSON.stringify({ error: 'insertText failed' });
          return JSON.stringify({ ok: true, at: 'end' });
        }
        return JSON.stringify({ error: 'document write API unavailable' });
      } catch (error) {
        return JSON.stringify({
          error: error instanceof Error ? error.message : 'insert_text failed',
        });
      }
    }
    if (name === 'replace_selection') {
      const text = String(args.text ?? '');
      try {
        const offsets = getDocActiveOffsets();
        if (!offsets) return JSON.stringify({ error: 'no active selection' });
        const { startOffset, endOffset } = offsets;
        if (startOffset === endOffset) {
          return JSON.stringify({ error: 'selection is empty (cursor only)' });
        }
        if (typeof doc.getTextRange === 'function') {
          const ok = doc.getTextRange(startOffset, endOffset).setText(text);
          if (!ok) return JSON.stringify({ error: 'setText failed' });
          return JSON.stringify({
            ok: true,
            startOffset,
            endOffset: startOffset + text.length,
            replacedChars: endOffset - startOffset,
          });
        }
        if (
          typeof doc.deleteRange === 'function' &&
          typeof doc.insertText === 'function'
        ) {
          doc.deleteRange({ startOffset, endOffset });
          const ok = doc.insertText(startOffset, text);
          if (!ok) return JSON.stringify({ error: 'insertText failed' });
          return JSON.stringify({
            ok: true,
            startOffset,
            endOffset: startOffset + text.length,
            replacedChars: endOffset - startOffset,
          });
        }
        return JSON.stringify({
          error: 'document replace selection API unavailable',
        });
      } catch (error) {
        return JSON.stringify({
          error:
            error instanceof Error ? error.message : 'replace_selection failed',
        });
      }
    }
    if (name === 'format_text') {
      try {
        let startOffset: number | null = null;
        let endOffset: number | null = null;
        const offsets = getDocActiveOffsets();
        if (offsets && offsets.startOffset !== offsets.endOffset) {
          startOffset = offsets.startOffset;
          endOffset = offsets.endOffset;
        } else if (args.paragraphIndex != null) {
          const resolved = resolveDocParagraph(doc, args.paragraphIndex);
          if ('error' in resolved) return JSON.stringify(resolved);
          const info = resolved.paragraph.getInfo?.();
          if (!info)
            return JSON.stringify({ error: 'paragraph offsets unavailable' });
          startOffset = info.startOffset;
          endOffset = info.endOffset;
        } else if (offsets) {
          return JSON.stringify({ error: 'selection is empty (cursor only)' });
        } else {
          return JSON.stringify({
            error: 'no selection; pass paragraphIndex or select text',
          });
        }
        if (
          startOffset == null ||
          endOffset == null ||
          startOffset === endOffset
        ) {
          return JSON.stringify({ error: 'empty range' });
        }
        if (typeof doc.getTextRange !== 'function') {
          return JSON.stringify({ error: 'getTextRange unavailable' });
        }
        const style: {
          bl?: number;
          it?: number;
          fs?: number;
          ul?: { s?: number };
          cl?: { rgb?: string };
        } = {};
        const applied: string[] = [];
        if (typeof args.bold === 'boolean') {
          style.bl = args.bold ? 1 : 0;
          applied.push(args.bold ? 'bold' : 'unbold');
        }
        if (typeof args.italic === 'boolean') {
          style.it = args.italic ? 1 : 0;
          applied.push(args.italic ? 'italic' : 'unitalic');
        }
        if (typeof args.underline === 'boolean') {
          style.ul = { s: args.underline ? 1 : 0 };
          applied.push(args.underline ? 'underline' : 'nounderline');
        }
        if (args.fontSize != null) {
          style.fs = Number(args.fontSize);
          applied.push(`fontSize:${style.fs}`);
        }
        if (args.color != null && String(args.color).trim()) {
          style.cl = { rgb: String(args.color).trim() };
          applied.push(`color:${style.cl.rgb}`);
        }
        if (!applied.length) {
          return JSON.stringify({ error: 'no style properties provided' });
        }
        const textRange = doc.getTextRange(startOffset, endOffset);
        if (typeof textRange.setTextStyle !== 'function') {
          return JSON.stringify({ error: 'setTextStyle unavailable' });
        }
        const ok = textRange.setTextStyle(style);
        if (!ok) return JSON.stringify({ error: 'setTextStyle failed' });
        return JSON.stringify({
          ok: true,
          startOffset,
          endOffset,
          applied,
        });
      } catch (error) {
        return JSON.stringify({
          error: error instanceof Error ? error.message : 'format_text failed',
        });
      }
    }
    if (name === 'replace_text') {
      const search = String(args.search ?? '');
      const replacement = String(args.replacement ?? '');
      if (!search) return JSON.stringify({ error: 'empty search' });
      try {
        let replaced = 0;
        if (typeof doc.getParagraphs === 'function') {
          const paragraphs = doc.getParagraphs();
          // reverse order so offsets stay valid while mutating
          [...paragraphs].reverse().forEach((paragraph) => {
            if (typeof paragraph.findAllText !== 'function') return;
            const ranges = paragraph.findAllText(search);
            [...ranges].reverse().forEach((range) => {
              const ok = range.setText(replacement);
              if (ok) replaced += 1;
            });
          });
          return JSON.stringify({ ok: true, replaced });
        }
        if (
          typeof doc.insertText === 'function' &&
          typeof doc.deleteRange === 'function'
        ) {
          let cursor = 0;
          const hits: number[] = [];
          while (cursor < stream.length) {
            const at = stream.indexOf(search, cursor);
            if (at === -1) break;
            hits.push(at);
            cursor = at + search.length;
          }
          // reverse order so offsets stay valid while mutating
          [...hits].reverse().forEach((start) => {
            const end = start + search.length;
            doc.deleteRange({ startOffset: start, endOffset: end });
            doc.insertText(start, replacement);
            replaced += 1;
          });
          return JSON.stringify({ ok: true, replaced });
        }
        return JSON.stringify({ error: 'document replace API unavailable' });
      } catch (error) {
        return JSON.stringify({
          error: error instanceof Error ? error.message : 'replace_text failed',
        });
      }
    }
  }

  return JSON.stringify({ error: `unknown tool ${name}` });
}
