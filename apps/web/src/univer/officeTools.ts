import {
  a1FromRowCol,
  a1Range,
  getActiveWorkbook,
  getDocActiveOffsets,
  getSelectionInfo,
  parseA1Cell,
} from './selection'
import type { UniverApi } from './useUniver'

export type ChatMode = 'chat' | 'plan' | 'agent'

const READ_MAX_ROWS = 80
const READ_MAX_COLS = 40
const READ_MAX_CELLS = 1200
const SEARCH_MAX_HITS = 50

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
      name: 'set_cells',
      description:
        'Write a 2D values grid starting at start (A1). If start omitted, uses active cell. Writes the full grid size.',
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
      ? ['get_workbook_outline', 'get_selection', 'read_range', 'search_in_sheet']
      : hasDocument && docType === 'DOC'
        ? ['get_doc_outline', 'get_selection', 'read_document']
        : []
  const write =
    hasDocument && docType === 'SHEET'
      ? ['set_cells', 'clear_range']
      : hasDocument && docType === 'DOC'
        ? ['replace_text', 'replace_selection', 'insert_text']
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
  if (name === 'set_cells') {
    const values = args.values as unknown[][] | undefined
    const cells = values?.reduce((n, row) => n + (row?.length ?? 0), 0) ?? 0
    return cells > 20
  }
  if (name === 'replace_text') return true
  if (name === 'replace_selection') return true
  return false
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
