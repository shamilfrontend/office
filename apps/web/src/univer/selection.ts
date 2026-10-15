import { DocSelectionManagerService } from '@univerjs/docs'
import type { UniverApi, UniverDocument } from './useUniver'

export type SheetSelectionInfo = {
  kind: 'SHEET'
  sheet: string
  range: string | null
  cell: string | null
  preview: string[][] | null
  truncated: boolean
}

export type DocSelectionInfo = {
  kind: 'DOC'
  startOffset: number | null
  endOffset: number | null
  text: string | null
  collapsed: boolean | null
}

export type SelectionInfo = SheetSelectionInfo | DocSelectionInfo | { kind: 'none'; reason: string }

export type SheetFilterCriteria = {
  colId?: number
  filters?: { filters?: string[] }
}

export type SheetFilterHandle = {
  getRange?: () => SheetRange | { getA1Notation?: () => string }
  remove?: () => unknown
  setColumnFilterCriteria?: (
    column: number,
    criteria: SheetFilterCriteria | null,
  ) => unknown
  removeColumnFilterCriteria?: (column: number) => unknown
  getColumnFilterCriteria?: (column: number) => SheetFilterCriteria | null
  getFilteredOutRows?: () => number[] | Set<number>
}

export type SheetSortBy =
  | number
  | { column: number; ascending?: boolean }
  | Array<{ column: number; ascending?: boolean }>

export type CfRuleLike = {
  cfId?: string
  ranges?: Array<{
    startRow: number
    startColumn: number
    endRow: number
    endColumn: number
  }>
  rule?: { type?: string; subType?: string; operator?: string; value?: unknown }
}

export type DataValidationRuleLike = {
  rule?: { uid?: string; type?: string; formula1?: string; formula2?: string }
  getCriteriaType?: () => string
  getCriteriaValues?: () => [string | undefined, string | undefined, string | undefined]
  getRanges?: () => Array<{ getA1Notation?: () => string }>
  uid?: string
}

export type CfRuleBuilder = {
  whenCellEmpty: () => CfRuleBuilder
  whenCellNotEmpty: () => CfRuleBuilder
  whenNumberBetween: (start: number, end: number) => CfRuleBuilder
  whenTextContains: (text: string) => CfRuleBuilder
  setDuplicateValues: () => CfRuleBuilder
  setBackground: (color: string) => CfRuleBuilder
  setFontColor: (color: string) => CfRuleBuilder
  setRanges: (ranges: unknown[]) => CfRuleBuilder
  build: () => CfRuleLike
}

export type SheetRange = {
  getA1Notation: (withSheet?: boolean) => string
  getValues: () => unknown[][]
  getFormulas: () => string[][]
  getDisplayValues?: () => string[][]
  getRange?: () => { startRow: number; startColumn: number; endRow: number; endColumn: number }
  getColumn?: () => number
  setValues: (v: unknown[][]) => void
  setFormulas: (f: string[][]) => void
  clear: () => void
  sort?: (by: SheetSortBy) => unknown
  createFilter?: () => SheetFilterHandle | null
  getFilter?: () => SheetFilterHandle | null
  setFontWeight?: (weight: 'bold' | 'normal' | null) => unknown
  setFontSize?: (size: number | null) => unknown
  setFontColor?: (color: string | null) => unknown
  setBackgroundColor?: (color: string) => unknown
  setHorizontalAlignment?: (align: 'left' | 'center' | 'normal') => unknown
  setWrap?: (wrap: boolean) => unknown
  setNumberFormat?: (pattern: string) => unknown
  setBorder?: (type: string, style: number, color?: string) => unknown
  merge?: (options?: { isForceMerge?: boolean }) => unknown
  mergeAcross?: (options?: { isForceMerge?: boolean }) => unknown
  mergeVertically?: (options?: { isForceMerge?: boolean }) => unknown
  breakApart?: () => unknown
  getConditionalFormattingRules?: () => CfRuleLike[]
  createConditionalFormattingRule?: () => CfRuleBuilder
  clearConditionalFormatRules?: () => unknown
  setDataValidation?: (rule: unknown | null) => unknown
  getDataValidation?: () => DataValidationRuleLike | null
  getDataValidations?: () => DataValidationRuleLike[]
}

export type SheetHandle = {
  getSheetName: () => string
  getSheetId?: () => string
  getActiveRange: () => SheetRange | null
  getActiveCell: () => SheetRange | null
  getRange: {
    (a1: string): SheetRange
    (row: number, column: number, numRows: number, numColumns: number): SheetRange
  }
  getMaxRows: () => number
  getMaxColumns: () => number
  getDataRange: () => SheetRange
  setName?: (name: string) => unknown
  activate?: () => unknown
  getFilter?: () => SheetFilterHandle | null
  sort?: (by: SheetSortBy) => unknown
  setColumnWidth?: (columnPosition: number, width: number) => unknown
  setColumnWidths?: (startColumn: number, numColumn: number, width: number) => unknown
  setRowHeight?: (rowPosition: number, height: number) => unknown
  setRowHeights?: (startRow: number, numRows: number, height: number) => unknown
  setFrozenRows?: (rows: number) => unknown
  setFrozenColumns?: (columns: number) => unknown
  cancelFreeze?: () => unknown
  getFrozenRows?: () => number
  getFrozenColumns?: () => number
  newConditionalFormattingRule?: () => CfRuleBuilder
  addConditionalFormattingRule?: (rule: CfRuleLike) => unknown
  getConditionalFormattingRules?: () => CfRuleLike[]
  deleteConditionalFormattingRule?: (cfId: string) => unknown
  clearConditionalFormatRules?: () => unknown
  getDataValidations?: () => DataValidationRuleLike[]
}

export type SheetApi = {
  getActiveSheet: () => SheetHandle
  getActiveRange?: () => SheetRange | null
  getActiveCell?: () => SheetRange | null
  getSheets: () => SheetHandle[]
  getSheetByName?: (name: string) => SheetHandle | null
  insertSheet?: (name?: string, options?: { index?: number }) => SheetHandle | null
  deleteSheet?: (sheet: SheetHandle | string) => boolean
  setActiveSheet?: (sheet: SheetHandle | string) => unknown
}

/** Estimate cell count of an A1 range (e.g. A1:C10 → 30). */
export function estimateA1Area(a1: string): number {
  const clean = String(a1 || '')
    .trim()
    .replace(/^[^!]*!/, '')
  if (!clean) return 0
  const parts = clean.split(':')
  const start = parseA1Cell(parts[0])
  if (!start) return 0
  const end = parts[1] ? parseA1Cell(parts[1]) : start
  if (!end) return 0
  return (Math.abs(end.row - start.row) + 1) * (Math.abs(end.col - start.col) + 1)
}

export function resolveColumnIndex(column: string | number): number | null {
  if (typeof column === 'number' && Number.isFinite(column)) {
    return column >= 0 ? Math.floor(column) : null
  }
  const text = String(column || '').trim()
  if (/^\d+$/.test(text)) {
    const n = Number(text)
    return n >= 0 ? n : null
  }
  const idx = colLettersToIndex(text)
  return idx >= 0 ? idx : null
}

type InjectorLike = {
  get: (token: unknown) => unknown
}

let currentInjector: InjectorLike | null = null

export function setSelectionInjector(injector: InjectorLike | null) {
  currentInjector = injector
}

export function getSelectionInjector(): InjectorLike | null {
  return currentInjector
}

export function parseA1Cell(a1: string): { row: number; col: number } | null {
  const m = String(a1 || '')
    .trim()
    .replace(/^[^!]*!/, '')
    .match(/^\$?([A-Za-z]+)\$?(\d+)$/)
  if (!m) return null
  return { col: colLettersToIndex(m[1]), row: Number(m[2]) - 1 }
}

export function colLettersToIndex(col: string): number {
  let n = 0
  for (const ch of col.toUpperCase()) {
    if (ch < 'A' || ch > 'Z') return -1
    n = n * 26 + (ch.charCodeAt(0) - 64)
  }
  return n - 1
}

export function indexToColLetters(col: number): string {
  let n = col + 1
  let s = ''
  while (n > 0) {
    const r = (n - 1) % 26
    s = String.fromCharCode(65 + r) + s
    n = Math.floor((n - 1) / 26)
  }
  return s
}

export function a1FromRowCol(row: number, col: number): string {
  return `${indexToColLetters(col)}${row + 1}`
}

export function a1Range(startRow: number, startCol: number, numRows: number, numCols: number): string {
  const endRow = startRow + Math.max(1, numRows) - 1
  const endCol = startCol + Math.max(1, numCols) - 1
  const start = a1FromRowCol(startRow, startCol)
  const end = a1FromRowCol(endRow, endCol)
  return start === end ? start : `${start}:${end}`
}

const PREVIEW_MAX_ROWS = 20
const PREVIEW_MAX_COLS = 10
const PREVIEW_MAX_CELLS = 200
const DOC_TEXT_CAP = 500

function capMatrix(values: unknown[][]): { values: unknown[][]; truncated: boolean } {
  const rows = values.length
  const cols = values.reduce((m, row) => Math.max(m, row?.length ?? 0), 0)
  const cells = rows * Math.max(cols, 1)
  if (rows <= PREVIEW_MAX_ROWS && cols <= PREVIEW_MAX_COLS && cells <= PREVIEW_MAX_CELLS) {
    return { values, truncated: false }
  }
  const sliced = values.slice(0, PREVIEW_MAX_ROWS).map((row) =>
    (row ?? []).slice(0, PREVIEW_MAX_COLS),
  )
  return { values: sliced, truncated: true }
}

export function getActiveWorkbook(api: UniverApi | null): SheetApi | null {
  if (!api) return null
  return (api.getActiveWorkbook() as unknown as SheetApi | null) ?? null
}

export function getSheetSelection(api: UniverApi | null): SheetSelectionInfo | null {
  const wb = getActiveWorkbook(api)
  if (!wb) return null
  try {
    const sheet = wb.getActiveSheet()
    const activeRange = sheet.getActiveRange?.() ?? wb.getActiveRange?.() ?? null
    const activeCell = sheet.getActiveCell?.() ?? wb.getActiveCell?.() ?? null
    let preview: string[][] | null = null
    let truncated = false
    if (activeRange) {
      const raw =
        typeof activeRange.getDisplayValues === 'function'
          ? activeRange.getDisplayValues()
          : (activeRange.getValues() as unknown[][]).map((row) =>
              row.map((cell) => (cell == null ? '' : String(cell))),
            )
      const capped = capMatrix(raw)
      preview = capped.values as string[][]
      truncated = capped.truncated
    }
    return {
      kind: 'SHEET',
      sheet: sheet.getSheetName(),
      range: activeRange?.getA1Notation?.() ?? null,
      cell: activeCell?.getA1Notation?.() ?? null,
      preview,
      truncated,
    }
  } catch {
    return null
  }
}

function getDocStream(doc: UniverDocument): string {
  return (
    doc.getBody?.()?.dataStream ??
    (typeof doc.getSnapshot === 'function' ? doc.getSnapshot()?.body?.dataStream : undefined) ??
    doc.save()?.body?.dataStream ??
    ''
  )
}

export function getDocActiveOffsets(): {
  startOffset: number
  endOffset: number
} | null {
  if (!currentInjector) return null
  try {
    const svc = currentInjector.get(DocSelectionManagerService) as {
      getActiveTextRange?: () => { startOffset?: number; endOffset?: number } | null
    }
    const range = svc?.getActiveTextRange?.()
    if (
      !range ||
      typeof range.startOffset !== 'number' ||
      typeof range.endOffset !== 'number'
    ) {
      return null
    }
    return { startOffset: range.startOffset, endOffset: range.endOffset }
  } catch {
    return null
  }
}

export function getDocSelection(api: UniverApi | null): DocSelectionInfo | null {
  const doc = api?.getActiveDocument?.() ?? null
  if (!doc) return null
  try {
    const offsets = getDocActiveOffsets()
    if (!offsets) {
      return {
        kind: 'DOC',
        startOffset: null,
        endOffset: null,
        text: null,
        collapsed: null,
      }
    }
    const { startOffset, endOffset } = offsets
    let text = ''
    try {
      if (typeof doc.getTextRange === 'function') {
        text = doc.getTextRange(startOffset, endOffset).getText?.() ?? ''
      } else {
        text = getDocStream(doc).slice(startOffset, endOffset)
      }
    } catch {
      text = getDocStream(doc).slice(startOffset, endOffset)
    }
    const capped =
      text.length > DOC_TEXT_CAP ? `${text.slice(0, DOC_TEXT_CAP)}…` : text
    return {
      kind: 'DOC',
      startOffset,
      endOffset,
      text: capped,
      collapsed: startOffset === endOffset,
    }
  } catch {
    return null
  }
}

export function getSelectionInfo(
  api: UniverApi | null,
  docType: 'SHEET' | 'DOC',
): SelectionInfo {
  if (!api) return { kind: 'none', reason: 'editor not ready' }
  if (docType === 'SHEET') {
    return getSheetSelection(api) ?? { kind: 'none', reason: 'no sheet selection' }
  }
  return getDocSelection(api) ?? { kind: 'none', reason: 'no doc selection' }
}

export function formatSelectionContext(
  api: UniverApi | null,
  docType: 'SHEET' | 'DOC',
  meta?: { docName?: string },
): string {
  const lines: string[] = []
  if (meta?.docName) lines.push(`Документ: ${meta.docName} (${docType})`)
  const info = getSelectionInfo(api, docType)
  if (info.kind === 'none') {
    lines.push(`Выделение: недоступно (${info.reason}). Используй tools для чтения.`)
    return lines.join('\n')
  }
  if (info.kind === 'SHEET') {
    lines.push(`Активный лист: ${info.sheet}`)
    if (info.range) lines.push(`Выделение: ${info.range}`)
    if (info.cell) lines.push(`Активная ячейка: ${info.cell}`)
    if (info.preview) {
      lines.push(
        `Превью выделения${info.truncated ? ' (обрезано)' : ''}: ${JSON.stringify(info.preview)}`,
      )
    }
    if (!info.range && !info.cell) {
      lines.push('Выделение: нет активного range. Используй get_selection / read_range.')
    }
    return lines.join('\n')
  }
  if (info.startOffset == null || info.endOffset == null) {
    lines.push('Выделение: нет активного текстового range. Используй get_selection.')
    return lines.join('\n')
  }
  lines.push(`Выделение: offsets ${info.startOffset}–${info.endOffset}`)
  if (info.collapsed) lines.push('Курсор (пустое выделение)')
  else if (info.text != null) lines.push(`Текст: ${JSON.stringify(info.text)}`)
  return lines.join('\n')
}
