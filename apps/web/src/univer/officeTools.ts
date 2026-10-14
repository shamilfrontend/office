import { BorderStyleTypes, BorderType } from '@univerjs/core'
import {
  a1FromRowCol,
  a1Range,
  estimateA1Area,
  getActiveWorkbook,
  getDocActiveOffsets,
  getSelectionInfo,
  indexToColLetters,
  parseA1Cell,
  resolveColumnIndex,
  type SheetRange,
} from './selection'
import type { UniverApi } from './useUniver'

export type ChatMode = 'chat' | 'plan' | 'agent'

const READ_MAX_ROWS = 80
const READ_MAX_COLS = 40
const READ_MAX_CELLS = 1200
const SEARCH_MAX_HITS = 50
const FORMAT_APPROVE_CELLS = 50

export const OFFICE_TOOLS = [
  {
    type: 'function' as const,
    function: {
      name: 'get_workbook_outline',
      description: 'Get sheet names and rough size of the active workbook',
      parameters: { type: 'object', properties: {} },
    },
  },
  {
    type: 'function' as const,
    function: {
      name: 'get_selection',
      description:
        'Get current selection: sheet range/cell + preview, or document text offsets/text',
      parameters: { type: 'object', properties: {} },
    },
  },
  {
    type: 'function' as const,
    function: {
      name: 'read_range',
      description:
        'Read cell values from A1 range on active sheet. If range omitted, uses active selection.',
      parameters: {
        type: 'object',
        properties: {
          range: { type: 'string', description: 'A1 notation, e.g. A1:C10. Optional: active selection.' },
        },
      },
    },
  },
  {
    type: 'function' as const,
    function: {
      name: 'search_in_sheet',
      description: 'Search text in used range of the active sheet; returns A1 hits',
      parameters: {
        type: 'object',
        properties: { query: { type: 'string' } },
        required: ['query'],
      },
    },
  },
  {
    type: 'function' as const,
    function: {
      name: 'read_formulas',
      description:
        'Read formula strings and calculated values from A1 range (or active selection). Prefer this over read_range when you need formulas.',
      parameters: {
        type: 'object',
        properties: {
          range: {
            type: 'string',
            description: 'A1 notation, e.g. A1:C10. Optional: active selection.',
          },
        },
      },
    },
  },
  {
    type: 'function' as const,
    function: {
      name: 'set_cells',
      description:
        'Write a 2D values grid starting at start (A1). If start omitted, uses active cell. Writes the full grid size. Use for literals only; for formulas use set_formulas.',
      parameters: {
        type: 'object',
        properties: {
          start: { type: 'string', description: 'A1 start cell; optional → active cell' },
          values: {
            type: 'array',
            items: { type: 'array', items: {} },
          },
        },
        required: ['values'],
      },
    },
  },
  {
    type: 'function' as const,
    function: {
      name: 'set_formulas',
      description:
        'Write a 2D grid of spreadsheet formulas starting at start (A1). English function names, comma args, leading = (e.g. =SUM(A1:A10)).',
      parameters: {
        type: 'object',
        properties: {
          start: { type: 'string', description: 'A1 start cell; optional → active cell' },
          formulas: {
            type: 'array',
            items: { type: 'array', items: { type: 'string' } },
          },
        },
        required: ['formulas'],
      },
    },
  },
  {
    type: 'function' as const,
    function: {
      name: 'clear_range',
      description: 'Clear values in A1 range',
      parameters: {
        type: 'object',
        properties: { range: { type: 'string' } },
        required: ['range'],
      },
    },
  },
  {
    type: 'function' as const,
    function: {
      name: 'format_range',
      description:
        'Apply cell styles to an A1 range (or active selection): bold, font size/color, background, align, wrap',
      parameters: {
        type: 'object',
        properties: {
          range: { type: 'string', description: 'A1 range; optional → active selection' },
          bold: { type: 'boolean' },
          fontSize: { type: 'number' },
          fontColor: { type: 'string' },
          backgroundColor: { type: 'string' },
          horizontalAlign: {
            type: 'string',
            enum: ['left', 'center', 'right'],
          },
          wrap: { type: 'boolean' },
        },
      },
    },
  },
  {
    type: 'function' as const,
    function: {
      name: 'set_column_widths',
      description: 'Set column widths on the active sheet',
      parameters: {
        type: 'object',
        properties: {
          widths: {
            type: 'array',
            items: {
              type: 'object',
              properties: {
                column: {
                  description: 'Column letter (A) or 0-based index',
                },
                width: { type: 'number' },
              },
              required: ['column', 'width'],
            },
          },
          startColumn: {
            description: 'Start column letter or 0-based index (with numColumns+width)',
          },
          numColumns: { type: 'number' },
          width: { type: 'number' },
        },
      },
    },
  },
  {
    type: 'function' as const,
    function: {
      name: 'set_number_format',
      description: 'Set number format pattern on a range (or active selection)',
      parameters: {
        type: 'object',
        properties: {
          range: { type: 'string', description: 'A1 range; optional → active selection' },
          pattern: {
            type: 'string',
            description: 'e.g. #,##0.00 or 0%',
          },
        },
        required: ['pattern'],
      },
    },
  },
  {
    type: 'function' as const,
    function: {
      name: 'merge_range',
      description:
        'Merge or unmerge cells in a range (or active selection). Modes: merge, across, vertically, unmerge',
      parameters: {
        type: 'object',
        properties: {
          range: { type: 'string', description: 'A1 range; optional → active selection' },
          mode: {
            type: 'string',
            enum: ['merge', 'across', 'vertically', 'unmerge'],
            description: 'Default merge',
          },
        },
      },
    },
  },
  {
    type: 'function' as const,
    function: {
      name: 'set_borders',
      description: 'Set borders on a range (or active selection)',
      parameters: {
        type: 'object',
        properties: {
          range: { type: 'string', description: 'A1 range; optional → active selection' },
          type: {
            type: 'string',
            enum: [
              'all',
              'outside',
              'inside',
              'top',
              'bottom',
              'left',
              'right',
              'horizontal',
              'vertical',
              'none',
            ],
          },
          style: {
            type: 'string',
            enum: ['thin', 'medium', 'thick', 'dashed', 'dotted', 'double', 'none'],
            description: 'Default thin',
          },
          color: { type: 'string', description: 'Hex color, default #000000' },
        },
        required: ['type'],
      },
    },
  },
  {
    type: 'function' as const,
    function: {
      name: 'set_row_heights',
      description: 'Set row heights on the active sheet (rows are 1-based)',
      parameters: {
        type: 'object',
        properties: {
          heights: {
            type: 'array',
            items: {
              type: 'object',
              properties: {
                row: { type: 'number', description: '1-based row number' },
                height: { type: 'number' },
              },
              required: ['row', 'height'],
            },
          },
          startRow: { type: 'number', description: '1-based start row' },
          numRows: { type: 'number' },
          height: { type: 'number' },
        },
      },
    },
  },
  {
    type: 'function' as const,
    function: {
      name: 'get_doc_outline',
      description: 'Get short plain-text preview of the active document',
      parameters: { type: 'object', properties: {} },
    },
  },
  {
    type: 'function' as const,
    function: {
      name: 'read_document',
      description: 'Read document text snapshot body stream (truncated)',
      parameters: {
        type: 'object',
        properties: {
          maxChars: { type: 'number' },
        },
      },
    },
  },
  {
    type: 'function' as const,
    function: {
      name: 'replace_text',
      description: 'Replace all occurrences of search with replacement in document body stream',
      parameters: {
        type: 'object',
        properties: {
          search: { type: 'string' },
          replacement: { type: 'string' },
        },
        required: ['search', 'replacement'],
      },
    },
  },
  {
    type: 'function' as const,
    function: {
      name: 'replace_selection',
      description: 'Replace currently selected document text only',
      parameters: {
        type: 'object',
        properties: {
          text: { type: 'string', description: 'Replacement text' },
        },
        required: ['text'],
      },
    },
  },
  {
    type: 'function' as const,
    function: {
      name: 'insert_text',
      description:
        'Insert text into the document. at=end appends; at=selection inserts at cursor or replaces non-empty selection.',
      parameters: {
        type: 'object',
        properties: {
          text: { type: 'string' },
          at: {
            type: 'string',
            enum: ['end', 'selection'],
            description: 'Default end',
          },
        },
        required: ['text'],
      },
    },
  },
  {
    type: 'function' as const,
    function: {
      name: 'format_text',
      description:
        'Apply text styles to the current document selection (bold, italic, fontSize in pt)',
      parameters: {
        type: 'object',
        properties: {
          bold: { type: 'boolean' },
          italic: { type: 'boolean' },
          fontSize: { type: 'number', description: 'Font size in points' },
          at: {
            type: 'string',
            enum: ['selection'],
            description: 'Only selection is supported',
          },
        },
      },
    },
  },
  {
    type: 'function' as const,
    function: {
      name: 'create_plan',
      description: 'Create an execution plan for the user to review',
      parameters: {
        type: 'object',
        properties: {
          name: { type: 'string' },
          overview: { type: 'string' },
          todos: {
            type: 'array',
            items: {
              type: 'object',
              properties: {
                id: { type: 'string' },
                content: { type: 'string' },
              },
              required: ['id', 'content'],
            },
          },
        },
        required: ['name', 'overview', 'todos'],
      },
    },
  },
  {
    type: 'function' as const,
    function: {
      name: 'update_plan',
      description: 'Update plan todos progress',
      parameters: {
        type: 'object',
        properties: {
          todos: {
            type: 'array',
            items: {
              type: 'object',
              properties: {
                id: { type: 'string' },
                content: { type: 'string' },
                status: { type: 'string' },
              },
              required: ['id'],
            },
          },
        },
        required: ['todos'],
      },
    },
  },
]

export function toolsForMode(
  mode: ChatMode,
  buildingPlan = false,
  hasDocument = true,
  docType: 'SHEET' | 'DOC' = 'SHEET',
) {
  const read =
    hasDocument && docType === 'SHEET'
      ? [
          'get_workbook_outline',
          'get_selection',
          'read_range',
          'read_formulas',
          'search_in_sheet',
        ]
      : hasDocument && docType === 'DOC'
        ? ['get_doc_outline', 'get_selection', 'read_document']
        : []
  const write =
    hasDocument && docType === 'SHEET'
      ? [
          'set_cells',
          'set_formulas',
          'clear_range',
          'format_range',
          'set_column_widths',
          'set_number_format',
          'merge_range',
          'set_borders',
          'set_row_heights',
        ]
      : hasDocument && docType === 'DOC'
        ? ['replace_text', 'replace_selection', 'insert_text', 'format_text']
        : []
  const plan = ['create_plan', 'update_plan']
  let names: string[] = []
  if (mode === 'chat') names = [...read]
  if (mode === 'plan') names = [...read, ...plan]
  if (mode === 'agent') {
    names = [...read, ...write]
    if (buildingPlan) names.push('update_plan')
  }
  return OFFICE_TOOLS.filter((t) => names.includes(t.function.name))
}

export function isMassWrite(name: string, args: Record<string, unknown>) {
  if (name === 'clear_range') return true
  if (name === 'merge_range') return true
  if (name === 'set_cells') {
    const values = args.values as unknown[][] | undefined
    const cells = values?.reduce((n, row) => n + (row?.length ?? 0), 0) ?? 0
    return cells > 20
  }
  if (name === 'set_formulas') {
    const formulas = args.formulas as unknown[][] | undefined
    const cells = formulas?.reduce((n, row) => n + (row?.length ?? 0), 0) ?? 0
    return cells > 20
  }
  if (
    name === 'format_range' ||
    name === 'set_number_format' ||
    name === 'set_borders'
  ) {
    const range = args.range != null ? String(args.range).trim() : ''
    if (!range) return false
    return estimateA1Area(range) > FORMAT_APPROVE_CELLS
  }
  if (name === 'set_column_widths') return false
  if (name === 'set_row_heights') return false
  if (name === 'format_text') return false
  if (name === 'replace_text') return true
  if (name === 'replace_selection') return true
  return false
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
  }
  return map[value] ?? null
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
  }
  return map[value] ?? null
}

function resolveSheetRangeA1(
  sheet: ReturnType<NonNullable<ReturnType<typeof getActiveWorkbook>>['getActiveSheet']>,
  wb: NonNullable<ReturnType<typeof getActiveWorkbook>>,
  rangeArg: unknown,
): { a1: string; range: SheetRange } | { error: string } {
  let a1 = rangeArg != null ? String(rangeArg).trim() : ''
  if (!a1) {
    const active = sheet.getActiveRange?.() ?? wb.getActiveRange?.()
    a1 = active?.getA1Notation?.() ?? ''
  }
  if (!a1) return { error: 'no range; pass range or select cells' }
  try {
    return { a1, range: sheet.getRange(a1) }
  } catch (error) {
    return {
      error: error instanceof Error ? error.message : 'invalid range',
    }
  }
}

function mapHorizontalAlign(value: string): 'left' | 'center' | 'normal' | null {
  if (value === 'left' || value === 'center') return value
  if (value === 'right') return 'normal'
  return null
}

function truncateMatrix(values: unknown[][]): {
  values: unknown[][]
  truncated: boolean
  rows: number
  cols: number
} {
  const rows = values.length
  const cols = values.reduce((m, row) => Math.max(m, row?.length ?? 0), 0)
  const cells = rows * Math.max(cols, 1)
  if (rows <= READ_MAX_ROWS && cols <= READ_MAX_COLS && cells <= READ_MAX_CELLS) {
    return { values, truncated: false, rows, cols }
  }
  return {
    values: values.slice(0, READ_MAX_ROWS).map((row) => (row ?? []).slice(0, READ_MAX_COLS)),
    truncated: true,
    rows,
    cols,
  }
}

export async function runOfficeTool(
  name: string,
  args: Record<string, unknown>,
  ctx: {
    api: UniverApi | null
    docType: 'SHEET' | 'DOC'
    plan: { name: string; overview: string; todos: Array<{ id: string; content: string; status?: string }> } | null
    setPlan: (p: typeof ctx.plan) => void
  },
): Promise<string> {
  if (name === 'create_plan') {
    const plan = {
      name: String(args.name || 'План'),
      overview: String(args.overview || ''),
      todos: (args.todos as Array<{ id: string; content: string }>) || [],
    }
    ctx.setPlan(plan)
    return JSON.stringify({ ok: true, plan })
  }
  if (name === 'update_plan') {
    if (!ctx.plan) return JSON.stringify({ error: 'no plan' })
    const todos = args.todos as Array<{ id: string; content?: string; status?: string }>
    const next = {
      ...ctx.plan,
      todos: ctx.plan.todos.map((t) => {
        const u = todos.find((x) => x.id === t.id)
        return u ? { ...t, ...u } : t
      }),
    }
    ctx.setPlan(next)
    return JSON.stringify({ ok: true, plan: next })
  }

  if (!ctx.api) return JSON.stringify({ error: 'editor not ready' })

  if (name === 'get_selection') {
    return JSON.stringify(getSelectionInfo(ctx.api, ctx.docType))
  }

  if (ctx.docType === 'SHEET') {
    const wb = getActiveWorkbook(ctx.api)
    if (!wb) return JSON.stringify({ error: 'no workbook' })
    const sheet = wb.getActiveSheet()

    if (name === 'get_workbook_outline') {
      let usedRange = ''
      try {
        usedRange = sheet.getDataRange().getA1Notation()
      } catch {
        usedRange = ''
      }
      return JSON.stringify({
        sheets: wb.getSheets().map((s) => s.getSheetName()),
        active: sheet.getSheetName(),
        rows: sheet.getMaxRows(),
        cols: sheet.getMaxColumns(),
        usedRange,
      })
    }
    if (name === 'read_range') {
      try {
        let a1 = args.range != null ? String(args.range).trim() : ''
        if (!a1) {
          const active = sheet.getActiveRange?.() ?? wb.getActiveRange?.()
          a1 = active?.getA1Notation?.() ?? ''
        }
        if (!a1) return JSON.stringify({ error: 'no range; pass range or select cells' })
        const range = sheet.getRange(a1)
        const raw = range.getValues()
        const { values, truncated, rows, cols } = truncateMatrix(raw as unknown[][])
        return JSON.stringify({
          sheet: sheet.getSheetName(),
          range: range.getA1Notation(),
          values,
          truncated,
          size: { rows, cols },
        })
      } catch (error) {
        return JSON.stringify({
          error: error instanceof Error ? error.message : 'read_range failed',
        })
      }
    }
    if (name === 'read_formulas') {
      try {
        let a1 = args.range != null ? String(args.range).trim() : ''
        if (!a1) {
          const active = sheet.getActiveRange?.() ?? wb.getActiveRange?.()
          a1 = active?.getA1Notation?.() ?? ''
        }
        if (!a1) return JSON.stringify({ error: 'no range; pass range or select cells' })
        const range = sheet.getRange(a1)
        const rawFormulas = range.getFormulas()
        const rawValues = range.getValues()
        const formulasTrunc = truncateMatrix(rawFormulas as unknown[][])
        const valuesTrunc = truncateMatrix(rawValues as unknown[][])
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
        })
      } catch (error) {
        return JSON.stringify({
          error: error instanceof Error ? error.message : 'read_formulas failed',
        })
      }
    }
    if (name === 'search_in_sheet') {
      try {
        const q = String(args.query || '').toLowerCase()
        if (!q) return JSON.stringify({ error: 'empty query' })
        let searchRange = 'A1:Z200'
        let originRow = 0
        let originCol = 0
        try {
          const data = sheet.getDataRange()
          searchRange = data.getA1Notation()
          const bounds = data.getRange?.()
          if (bounds) {
            originRow = bounds.startRow
            originCol = bounds.startColumn
          } else {
            const startCell = searchRange.split(':')[0]
            const parsed = parseA1Cell(startCell)
            if (parsed) {
              originRow = parsed.row
              originCol = parsed.col
            }
          }
        } catch {
          /* fallback A1:Z200 */
        }
        const values = sheet.getRange(searchRange).getValues() as unknown[][]
        const hits: Array<{ a1: string; v: unknown }> = []
        values.forEach((row, r) =>
          row.forEach((v, c) => {
            if (
              String(v ?? '')
                .toLowerCase()
                .includes(q)
            ) {
              hits.push({ a1: a1FromRowCol(originRow + r, originCol + c), v })
            }
          }),
        )
        return JSON.stringify({
          sheet: sheet.getSheetName(),
          range: searchRange,
          hits: hits.slice(0, SEARCH_MAX_HITS),
          truncated: hits.length > SEARCH_MAX_HITS,
        })
      } catch (error) {
        return JSON.stringify({
          error: error instanceof Error ? error.message : 'search_in_sheet failed',
        })
      }
    }
    if (name === 'set_cells') {
      try {
        const values = args.values as unknown[][]
        if (!Array.isArray(values) || !values.length) {
          return JSON.stringify({ error: 'values must be a non-empty 2D array' })
        }
        const numRows = values.length
        const numCols = Math.max(...values.map((row) => (Array.isArray(row) ? row.length : 0)), 0)
        if (!numCols) return JSON.stringify({ error: 'values rows are empty' })

        let startRow = 0
        let startCol = 0
        if (args.start != null && String(args.start).trim()) {
          const parsed = parseA1Cell(String(args.start))
          if (!parsed || parsed.row < 0 || parsed.col < 0) {
            return JSON.stringify({ error: `invalid start cell: ${String(args.start)}` })
          }
          startRow = parsed.row
          startCol = parsed.col
        } else {
          const cell = sheet.getActiveCell?.() ?? wb.getActiveCell?.()
          const a1 = cell?.getA1Notation?.()
          const parsed = a1 ? parseA1Cell(a1) : null
          if (!parsed) {
            return JSON.stringify({ error: 'no start and no active cell' })
          }
          startRow = parsed.row
          startCol = parsed.col
        }

        const normalized = values.map((row) => {
          const r = Array.isArray(row) ? [...row] : []
          while (r.length < numCols) r.push('')
          return r.slice(0, numCols)
        })
        const written = a1Range(startRow, startCol, numRows, numCols)
        sheet.getRange(startRow, startCol, numRows, numCols).setValues(normalized)
        return JSON.stringify({
          ok: true,
          sheet: sheet.getSheetName(),
          range: written,
          cells: numRows * numCols,
        })
      } catch (error) {
        return JSON.stringify({
          error: error instanceof Error ? error.message : 'set_cells failed',
        })
      }
    }
    if (name === 'set_formulas') {
      try {
        const formulas = args.formulas as unknown[][]
        if (!Array.isArray(formulas) || !formulas.length) {
          return JSON.stringify({ error: 'formulas must be a non-empty 2D array' })
        }
        const numRows = formulas.length
        const numCols = Math.max(
          ...formulas.map((row) => (Array.isArray(row) ? row.length : 0)),
          0,
        )
        if (!numCols) return JSON.stringify({ error: 'formulas rows are empty' })

        let startRow = 0
        let startCol = 0
        if (args.start != null && String(args.start).trim()) {
          const parsed = parseA1Cell(String(args.start))
          if (!parsed || parsed.row < 0 || parsed.col < 0) {
            return JSON.stringify({ error: `invalid start cell: ${String(args.start)}` })
          }
          startRow = parsed.row
          startCol = parsed.col
        } else {
          const cell = sheet.getActiveCell?.() ?? wb.getActiveCell?.()
          const a1 = cell?.getA1Notation?.()
          const parsed = a1 ? parseA1Cell(a1) : null
          if (!parsed) {
            return JSON.stringify({ error: 'no start and no active cell' })
          }
          startRow = parsed.row
          startCol = parsed.col
        }

        const normalized: string[][] = []
        for (let r = 0; r < numRows; r++) {
          const row = Array.isArray(formulas[r]) ? formulas[r] : []
          const out: string[] = []
          for (let c = 0; c < numCols; c++) {
            const raw = row[c] == null ? '' : String(row[c]).trim()
            if (!raw || raw === '=') {
              return JSON.stringify({
                error: `empty or invalid formula at row ${r}, col ${c}`,
              })
            }
            out.push(raw.startsWith('=') ? raw : `=${raw}`)
          }
          normalized.push(out)
        }

        const written = a1Range(startRow, startCol, numRows, numCols)
        sheet.getRange(startRow, startCol, numRows, numCols).setFormulas(normalized)
        return JSON.stringify({
          ok: true,
          sheet: sheet.getSheetName(),
          range: written,
          cells: numRows * numCols,
        })
      } catch (error) {
        return JSON.stringify({
          error: error instanceof Error ? error.message : 'set_formulas failed',
        })
      }
    }
    if (name === 'clear_range') {
      try {
        const a1 = String(args.range || '')
        if (!a1) return JSON.stringify({ error: 'range required' })
        sheet.getRange(a1).clear()
        return JSON.stringify({ ok: true, range: a1, sheet: sheet.getSheetName() })
      } catch (error) {
        return JSON.stringify({
          error: error instanceof Error ? error.message : 'clear_range failed',
        })
      }
    }
    if (name === 'format_range') {
      try {
        const resolved = resolveSheetRangeA1(sheet, wb, args.range)
        if ('error' in resolved) return JSON.stringify({ error: resolved.error })
        const { range } = resolved
        const applied: string[] = []
        if (typeof args.bold === 'boolean') {
          if (typeof range.setFontWeight !== 'function') {
            return JSON.stringify({ error: 'setFontWeight unavailable' })
          }
          range.setFontWeight(args.bold ? 'bold' : 'normal')
          applied.push(args.bold ? 'bold' : 'normal')
        }
        if (args.fontSize != null) {
          if (typeof range.setFontSize !== 'function') {
            return JSON.stringify({ error: 'setFontSize unavailable' })
          }
          range.setFontSize(Number(args.fontSize))
          applied.push(`fontSize:${Number(args.fontSize)}`)
        }
        if (args.fontColor != null) {
          if (typeof range.setFontColor !== 'function') {
            return JSON.stringify({ error: 'setFontColor unavailable' })
          }
          range.setFontColor(String(args.fontColor))
          applied.push(`fontColor:${String(args.fontColor)}`)
        }
        if (args.backgroundColor != null) {
          if (typeof range.setBackgroundColor !== 'function') {
            return JSON.stringify({ error: 'setBackgroundColor unavailable' })
          }
          range.setBackgroundColor(String(args.backgroundColor))
          applied.push(`backgroundColor:${String(args.backgroundColor)}`)
        }
        if (args.horizontalAlign != null) {
          const align = mapHorizontalAlign(String(args.horizontalAlign))
          if (!align) {
            return JSON.stringify({ error: 'horizontalAlign must be left|center|right' })
          }
          if (typeof range.setHorizontalAlignment !== 'function') {
            return JSON.stringify({ error: 'setHorizontalAlignment unavailable' })
          }
          range.setHorizontalAlignment(align)
          applied.push(`align:${String(args.horizontalAlign)}`)
        }
        if (typeof args.wrap === 'boolean') {
          if (typeof range.setWrap !== 'function') {
            return JSON.stringify({ error: 'setWrap unavailable' })
          }
          range.setWrap(args.wrap)
          applied.push(`wrap:${args.wrap}`)
        }
        if (!applied.length) {
          return JSON.stringify({ error: 'no style properties provided' })
        }
        return JSON.stringify({
          ok: true,
          sheet: sheet.getSheetName(),
          range: range.getA1Notation(),
          applied,
        })
      } catch (error) {
        return JSON.stringify({
          error: error instanceof Error ? error.message : 'format_range failed',
        })
      }
    }
    if (name === 'set_column_widths') {
      try {
        const applied: Array<{ column: string; index: number; width: number }> = []
        const widths = args.widths as Array<{ column: string | number; width: number }> | undefined
        if (Array.isArray(widths) && widths.length) {
          for (const item of widths) {
            const index = resolveColumnIndex(item.column)
            const width = Number(item.width)
            if (index == null || !Number.isFinite(width) || width <= 0) {
              return JSON.stringify({
                error: `invalid column/width: ${JSON.stringify(item)}`,
              })
            }
            if (typeof sheet.setColumnWidth !== 'function') {
              return JSON.stringify({ error: 'setColumnWidth unavailable' })
            }
            sheet.setColumnWidth(index, width)
            applied.push({ column: indexToColLetters(index), index, width })
          }
        } else if (args.startColumn != null && args.numColumns != null && args.width != null) {
          const start = resolveColumnIndex(args.startColumn as string | number)
          const num = Number(args.numColumns)
          const width = Number(args.width)
          if (start == null || !Number.isFinite(num) || num < 1 || !Number.isFinite(width) || width <= 0) {
            return JSON.stringify({ error: 'invalid startColumn/numColumns/width' })
          }
          if (typeof sheet.setColumnWidths === 'function') {
            sheet.setColumnWidths(start, num, width)
          } else if (typeof sheet.setColumnWidth === 'function') {
            for (let i = 0; i < num; i++) sheet.setColumnWidth(start + i, width)
          } else {
            return JSON.stringify({ error: 'setColumnWidth unavailable' })
          }
          for (let i = 0; i < num; i++) {
            applied.push({
              column: indexToColLetters(start + i),
              index: start + i,
              width,
            })
          }
        } else {
          return JSON.stringify({
            error: 'provide widths[] or startColumn+numColumns+width',
          })
        }
        return JSON.stringify({
          ok: true,
          sheet: sheet.getSheetName(),
          columns: applied,
        })
      } catch (error) {
        return JSON.stringify({
          error: error instanceof Error ? error.message : 'set_column_widths failed',
        })
      }
    }
    if (name === 'set_number_format') {
      try {
        const pattern = String(args.pattern ?? '')
        if (!pattern) return JSON.stringify({ error: 'pattern required' })
        const resolved = resolveSheetRangeA1(sheet, wb, args.range)
        if ('error' in resolved) return JSON.stringify({ error: resolved.error })
        const { range } = resolved
        if (typeof range.setNumberFormat !== 'function') {
          return JSON.stringify({ error: 'setNumberFormat unavailable' })
        }
        range.setNumberFormat(pattern)
        return JSON.stringify({
          ok: true,
          sheet: sheet.getSheetName(),
          range: range.getA1Notation(),
          pattern,
        })
      } catch (error) {
        return JSON.stringify({
          error: error instanceof Error ? error.message : 'set_number_format failed',
        })
      }
    }
    if (name === 'merge_range') {
      try {
        const mode = String(args.mode || 'merge')
        const resolved = resolveSheetRangeA1(sheet, wb, args.range)
        if ('error' in resolved) return JSON.stringify({ error: resolved.error })
        const { range } = resolved
        if (mode === 'unmerge') {
          if (typeof range.breakApart !== 'function') {
            return JSON.stringify({ error: 'breakApart unavailable' })
          }
          range.breakApart()
        } else if (mode === 'across') {
          if (typeof range.mergeAcross !== 'function') {
            return JSON.stringify({ error: 'mergeAcross unavailable' })
          }
          range.mergeAcross({ isForceMerge: true })
        } else if (mode === 'vertically') {
          if (typeof range.mergeVertically !== 'function') {
            return JSON.stringify({ error: 'mergeVertically unavailable' })
          }
          range.mergeVertically({ isForceMerge: true })
        } else if (mode === 'merge') {
          if (typeof range.merge !== 'function') {
            return JSON.stringify({ error: 'merge unavailable' })
          }
          range.merge({ isForceMerge: true })
        } else {
          return JSON.stringify({
            error: 'mode must be merge|across|vertically|unmerge',
          })
        }
        return JSON.stringify({
          ok: true,
          sheet: sheet.getSheetName(),
          range: range.getA1Notation(),
          mode,
        })
      } catch (error) {
        return JSON.stringify({
          error: error instanceof Error ? error.message : 'merge_range failed',
        })
      }
    }
    if (name === 'set_borders') {
      try {
        const type = mapBorderType(String(args.type || ''))
        if (type == null) {
          return JSON.stringify({
            error:
              'type must be all|outside|inside|top|bottom|left|right|horizontal|vertical|none',
          })
        }
        const styleName = String(args.style || 'thin')
        const style = mapBorderStyle(styleName)
        if (style == null) {
          return JSON.stringify({
            error: 'style must be thin|medium|thick|dashed|dotted|double|none',
          })
        }
        const color = args.color != null ? String(args.color) : '#000000'
        const resolved = resolveSheetRangeA1(sheet, wb, args.range)
        if ('error' in resolved) return JSON.stringify({ error: resolved.error })
        const { range } = resolved
        if (typeof range.setBorder !== 'function') {
          return JSON.stringify({ error: 'setBorder unavailable' })
        }
        range.setBorder(type, style, color)
        return JSON.stringify({
          ok: true,
          sheet: sheet.getSheetName(),
          range: range.getA1Notation(),
          type: String(args.type),
          style: styleName,
          color,
        })
      } catch (error) {
        return JSON.stringify({
          error: error instanceof Error ? error.message : 'set_borders failed',
        })
      }
    }
    if (name === 'set_row_heights') {
      try {
        const applied: Array<{ row: number; index: number; height: number }> = []
        const heights = args.heights as Array<{ row: number; height: number }> | undefined
        if (Array.isArray(heights) && heights.length) {
          for (const item of heights) {
            const row1 = Number(item.row)
            const height = Number(item.height)
            if (!Number.isFinite(row1) || row1 < 1 || !Number.isFinite(height) || height <= 0) {
              return JSON.stringify({
                error: `invalid row/height: ${JSON.stringify(item)}`,
              })
            }
            const index = Math.floor(row1) - 1
            if (typeof sheet.setRowHeight !== 'function') {
              return JSON.stringify({ error: 'setRowHeight unavailable' })
            }
            sheet.setRowHeight(index, height)
            applied.push({ row: Math.floor(row1), index, height })
          }
        } else if (args.startRow != null && args.numRows != null && args.height != null) {
          const start1 = Number(args.startRow)
          const num = Number(args.numRows)
          const height = Number(args.height)
          if (
            !Number.isFinite(start1) ||
            start1 < 1 ||
            !Number.isFinite(num) ||
            num < 1 ||
            !Number.isFinite(height) ||
            height <= 0
          ) {
            return JSON.stringify({ error: 'invalid startRow/numRows/height' })
          }
          const start = Math.floor(start1) - 1
          if (typeof sheet.setRowHeights === 'function') {
            sheet.setRowHeights(start, num, height)
          } else if (typeof sheet.setRowHeight === 'function') {
            for (let i = 0; i < num; i++) sheet.setRowHeight(start + i, height)
          } else {
            return JSON.stringify({ error: 'setRowHeight unavailable' })
          }
          for (let i = 0; i < num; i++) {
            applied.push({ row: start + i + 1, index: start + i, height })
          }
        } else {
          return JSON.stringify({
            error: 'provide heights[] or startRow+numRows+height',
          })
        }
        return JSON.stringify({
          ok: true,
          sheet: sheet.getSheetName(),
          rows: applied,
        })
      } catch (error) {
        return JSON.stringify({
          error: error instanceof Error ? error.message : 'set_row_heights failed',
        })
      }
    }
  }

  if (ctx.docType === 'DOC') {
    const doc = ctx.api.getActiveDocument()
    if (!doc) return JSON.stringify({ error: 'no document' })
    const stream =
      doc.getBody?.()?.dataStream ??
      (typeof doc.getSnapshot === 'function' ? doc.getSnapshot()?.body?.dataStream : undefined) ??
      doc.save()?.body?.dataStream ??
      ''

    if (name === 'get_doc_outline' || name === 'read_document') {
      const max = Number(args.maxChars ?? 4000)
      return JSON.stringify({
        length: stream.length,
        text: stream.slice(0, max),
      })
    }
    if (name === 'insert_text') {
      const text = String(args.text ?? '')
      if (!text) return JSON.stringify({ error: 'empty text' })
      const at = String(args.at || 'end') === 'selection' ? 'selection' : 'end'
      try {
        if (at === 'selection') {
          const offsets = getDocActiveOffsets()
          if (!offsets) return JSON.stringify({ error: 'no active selection' })
          const { startOffset, endOffset } = offsets
          if (startOffset !== endOffset && typeof doc.deleteRange === 'function') {
            doc.deleteRange({ startOffset, endOffset })
          }
          if (typeof doc.insertText !== 'function') {
            return JSON.stringify({ error: 'insertText unavailable' })
          }
          const ok = doc.insertText(startOffset, text)
          if (!ok) return JSON.stringify({ error: 'insertText failed' })
          return JSON.stringify({
            ok: true,
            at: 'selection',
            startOffset,
            endOffset: startOffset + text.length,
          })
        }
        const lines = text.replace(/\r\n/g, '\n').replace(/\r/g, '\n').split('\n')
        if (typeof doc.appendParagraph === 'function') {
          for (const line of lines) {
            doc.appendParagraph(line)
          }
          return JSON.stringify({ ok: true, at: 'end', paragraphs: lines.length })
        }
        if (typeof doc.insertText === 'function') {
          let index = stream.length
          if (stream.endsWith('\n')) index -= 1
          if (index > 0 && stream[index - 1] === '\r') index -= 1
          const ok = doc.insertText(Math.max(0, index), text)
          if (!ok) return JSON.stringify({ error: 'insertText failed' })
          return JSON.stringify({ ok: true, at: 'end' })
        }
        return JSON.stringify({ error: 'document write API unavailable' })
      } catch (error) {
        return JSON.stringify({
          error: error instanceof Error ? error.message : 'insert_text failed',
        })
      }
    }
    if (name === 'replace_selection') {
      const text = String(args.text ?? '')
      try {
        const offsets = getDocActiveOffsets()
        if (!offsets) return JSON.stringify({ error: 'no active selection' })
        const { startOffset, endOffset } = offsets
        if (startOffset === endOffset) {
          return JSON.stringify({ error: 'selection is empty (cursor only)' })
        }
        if (typeof doc.getTextRange === 'function') {
          const ok = doc.getTextRange(startOffset, endOffset).setText(text)
          if (!ok) return JSON.stringify({ error: 'setText failed' })
          return JSON.stringify({
            ok: true,
            startOffset,
            endOffset: startOffset + text.length,
            replacedChars: endOffset - startOffset,
          })
        }
        if (typeof doc.deleteRange === 'function' && typeof doc.insertText === 'function') {
          doc.deleteRange({ startOffset, endOffset })
          const ok = doc.insertText(startOffset, text)
          if (!ok) return JSON.stringify({ error: 'insertText failed' })
          return JSON.stringify({
            ok: true,
            startOffset,
            endOffset: startOffset + text.length,
            replacedChars: endOffset - startOffset,
          })
        }
        return JSON.stringify({ error: 'document replace selection API unavailable' })
      } catch (error) {
        return JSON.stringify({
          error: error instanceof Error ? error.message : 'replace_selection failed',
        })
      }
    }
    if (name === 'format_text') {
      try {
        const offsets = getDocActiveOffsets()
        if (!offsets) return JSON.stringify({ error: 'no active selection' })
        const { startOffset, endOffset } = offsets
        if (startOffset === endOffset) {
          return JSON.stringify({ error: 'selection is empty (cursor only)' })
        }
        if (typeof doc.getTextRange !== 'function') {
          return JSON.stringify({ error: 'getTextRange unavailable' })
        }
        const style: { bl?: number; it?: number; fs?: number } = {}
        const applied: string[] = []
        if (typeof args.bold === 'boolean') {
          style.bl = args.bold ? 1 : 0
          applied.push(args.bold ? 'bold' : 'unbold')
        }
        if (typeof args.italic === 'boolean') {
          style.it = args.italic ? 1 : 0
          applied.push(args.italic ? 'italic' : 'unitalic')
        }
        if (args.fontSize != null) {
          style.fs = Number(args.fontSize)
          applied.push(`fontSize:${style.fs}`)
        }
        if (!applied.length) {
          return JSON.stringify({ error: 'no style properties provided' })
        }
        const textRange = doc.getTextRange(startOffset, endOffset)
        if (typeof textRange.setTextStyle !== 'function') {
          return JSON.stringify({ error: 'setTextStyle unavailable' })
        }
        const ok = textRange.setTextStyle(style)
        if (!ok) return JSON.stringify({ error: 'setTextStyle failed' })
        return JSON.stringify({
          ok: true,
          startOffset,
          endOffset,
          applied,
        })
      } catch (error) {
        return JSON.stringify({
          error: error instanceof Error ? error.message : 'format_text failed',
        })
      }
    }
    if (name === 'replace_text') {
      const search = String(args.search ?? '')
      const replacement = String(args.replacement ?? '')
      if (!search) return JSON.stringify({ error: 'empty search' })
      try {
        let replaced = 0
        if (typeof doc.getParagraphs === 'function') {
          const paragraphs = doc.getParagraphs()
          for (let p = paragraphs.length - 1; p >= 0; p--) {
            const paragraph = paragraphs[p]
            if (typeof paragraph.findAllText !== 'function') continue
            const ranges = paragraph.findAllText(search)
            for (let i = ranges.length - 1; i >= 0; i--) {
              const ok = ranges[i].setText(replacement)
              if (ok) replaced += 1
            }
          }
          return JSON.stringify({ ok: true, replaced })
        }
        if (typeof doc.insertText === 'function' && typeof doc.deleteRange === 'function') {
          let cursor = 0
          const hits: number[] = []
          while (cursor < stream.length) {
            const at = stream.indexOf(search, cursor)
            if (at === -1) break
            hits.push(at)
            cursor = at + search.length
          }
          for (let i = hits.length - 1; i >= 0; i--) {
            const start = hits[i]
            const end = start + search.length
            doc.deleteRange({ startOffset: start, endOffset: end })
            doc.insertText(start, replacement)
            replaced += 1
          }
          return JSON.stringify({ ok: true, replaced })
        }
        return JSON.stringify({ error: 'document replace API unavailable' })
      } catch (error) {
        return JSON.stringify({
          error: error instanceof Error ? error.message : 'replace_text failed',
        })
      }
    }
  }

  return JSON.stringify({ error: `unknown tool ${name}` })
}
