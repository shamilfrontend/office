import type { UniverApi } from './useUniver'

export type ChatMode = 'chat' | 'plan' | 'agent'

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
      name: 'read_range',
      description: 'Read cell values from A1 range on active sheet',
      parameters: {
        type: 'object',
        properties: {
          range: { type: 'string', description: 'A1 notation, e.g. A1:C10' },
        },
        required: ['range'],
      },
    },
  },
  {
    type: 'function' as const,
    function: {
      name: 'search_in_sheet',
      description: 'Search text in active workbook cells',
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
      description: 'Write values into cells. values is 2D array starting at start cell',
      parameters: {
        type: 'object',
        properties: {
          start: { type: 'string', description: 'A1 start cell' },
          values: {
            type: 'array',
            items: { type: 'array', items: {} },
          },
        },
        required: ['start', 'values'],
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
      name: 'insert_text',
      description: 'Append text to the end of the document body stream',
      parameters: {
        type: 'object',
        properties: { text: { type: 'string' } },
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
      ? ['get_workbook_outline', 'read_range', 'search_in_sheet']
      : hasDocument && docType === 'DOC'
        ? ['get_doc_outline', 'read_document']
        : []
  const write =
    hasDocument && docType === 'SHEET'
      ? ['set_cells', 'clear_range']
      : hasDocument && docType === 'DOC'
        ? ['replace_text', 'insert_text']
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
    const cells = values?.reduce((n, row) => n + row.length, 0) ?? 0
    return cells > 20
  }
  if (name === 'replace_text') return true
  return false
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

  if (ctx.docType === 'SHEET') {
    const wb = ctx.api.getActiveWorkbook() as unknown as {
      getActiveSheet: () => {
        getSheetName: () => string
        getRange: (a1: string) => {
          getA1Notation: () => string
          getValues: () => unknown[][]
          setValues: (v: unknown[][]) => void
          clear: () => void
        }
        getMaxRows: () => number
        getMaxColumns: () => number
        getDataRange: () => { getA1Notation: () => string }
      }
      getSheets: () => Array<{ getSheetName: () => string }>
    } | null
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
      const range = sheet.getRange(String(args.range))
      return JSON.stringify({ range: range.getA1Notation(), values: range.getValues() })
    }
    if (name === 'search_in_sheet') {
      const q = String(args.query || '').toLowerCase()
      const values = sheet.getRange('A1:Z200').getValues()
      const hits: Array<{ r: number; c: number; v: unknown }> = []
      values.forEach((row, r) =>
        row.forEach((v, c) => {
          if (String(v ?? '')
            .toLowerCase()
            .includes(q))
            hits.push({ r, c, v })
        }),
      )
      return JSON.stringify({ hits: hits.slice(0, 50) })
    }
    if (name === 'set_cells') {
      const start = String(args.start || 'A1')
      const values = args.values as unknown[][]
      sheet.getRange(start).setValues(values)
      return JSON.stringify({ ok: true })
    }
    if (name === 'clear_range') {
      sheet.getRange(String(args.range)).clear()
      return JSON.stringify({ ok: true })
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
      const lines = text.replace(/\r\n/g, '\n').replace(/\r/g, '\n').split('\n')
      try {
        if (typeof doc.appendParagraph === 'function') {
          for (const line of lines) {
            doc.appendParagraph(line)
          }
          return JSON.stringify({ ok: true, paragraphs: lines.length })
        }
        if (typeof doc.insertText === 'function') {
          let index = stream.length
          if (stream.endsWith('\n')) index -= 1
          if (index > 0 && stream[index - 1] === '\r') index -= 1
          const ok = doc.insertText(Math.max(0, index), text)
          if (!ok) return JSON.stringify({ error: 'insertText failed' })
          return JSON.stringify({ ok: true })
        }
        return JSON.stringify({ error: 'document write API unavailable' })
      } catch (error) {
        return JSON.stringify({
          error: error instanceof Error ? error.message : 'insert_text failed',
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
