import type OpenAI from 'openai';

export type ChatMode = 'chat' | 'plan' | 'agent';
export type DocType = 'SHEET' | 'DOC';

/** Keep in sync with apps/web/src/univer/officeTools.ts OFFICE_TOOLS. */
export const OFFICE_TOOLS: OpenAI.Chat.Completions.ChatCompletionTool[] = [
  {
    type: 'function',
    function: {
      name: 'get_workbook_outline',
      description: 'Get sheet names and rough size of the active workbook',
      parameters: { type: 'object', properties: {} },
    },
  },
  {
    type: 'function',
    function: {
      name: 'get_selection',
      description:
        'Get current selection: sheet range/cell + preview, or document text offsets/text',
      parameters: { type: 'object', properties: {} },
    },
  },
  {
    type: 'function',
    function: {
      name: 'read_range',
      description:
        'Read cell values from A1 range on active sheet. If range omitted, uses active selection.',
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
    type: 'function',
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
    type: 'function',
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
    type: 'function',
    function: {
      name: 'set_cells',
      description:
        'Write a 2D values grid starting at start (A1). If start omitted, uses active cell. Writes the full grid size. Use for literals only; for formulas use set_formulas.',
      parameters: {
        type: 'object',
        properties: {
          start: {
            type: 'string',
            description: 'A1 start cell; optional → active cell',
          },
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
    type: 'function',
    function: {
      name: 'set_formulas',
      description:
        'Write a 2D grid of spreadsheet formulas starting at start (A1). English function names, comma args, leading = (e.g. =SUM(A1:A10)).',
      parameters: {
        type: 'object',
        properties: {
          start: {
            type: 'string',
            description: 'A1 start cell; optional → active cell',
          },
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
    type: 'function',
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
    type: 'function',
    function: {
      name: 'format_range',
      description:
        'Apply cell styles to an A1 range (or active selection): bold, font size/color, background, align, wrap',
      parameters: {
        type: 'object',
        properties: {
          range: {
            type: 'string',
            description: 'A1 range; optional → active selection',
          },
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
    type: 'function',
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
            description:
              'Start column letter or 0-based index (with numColumns+width)',
          },
          numColumns: { type: 'number' },
          width: { type: 'number' },
        },
      },
    },
  },
  {
    type: 'function',
    function: {
      name: 'set_number_format',
      description: 'Set number format pattern on a range (or active selection)',
      parameters: {
        type: 'object',
        properties: {
          range: {
            type: 'string',
            description: 'A1 range; optional → active selection',
          },
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
    type: 'function',
    function: {
      name: 'merge_range',
      description:
        'Merge or unmerge cells in a range (or active selection). Modes: merge, across, vertically, unmerge',
      parameters: {
        type: 'object',
        properties: {
          range: {
            type: 'string',
            description: 'A1 range; optional → active selection',
          },
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
    type: 'function',
    function: {
      name: 'set_borders',
      description: 'Set borders on a range (or active selection)',
      parameters: {
        type: 'object',
        properties: {
          range: {
            type: 'string',
            description: 'A1 range; optional → active selection',
          },
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
            enum: [
              'thin',
              'medium',
              'thick',
              'dashed',
              'dotted',
              'double',
              'none',
            ],
            description: 'Default thin',
          },
          color: { type: 'string', description: 'Hex color, default #000000' },
        },
        required: ['type'],
      },
    },
  },
  {
    type: 'function',
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
    type: 'function',
    function: {
      name: 'get_doc_outline',
      description: 'Get short plain-text preview of the active document',
      parameters: { type: 'object', properties: {} },
    },
  },
  {
    type: 'function',
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
    type: 'function',
    function: {
      name: 'replace_text',
      description:
        'Replace all occurrences of search with replacement in document body stream',
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
    type: 'function',
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
    type: 'function',
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
    type: 'function',
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
    type: 'function',
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
    type: 'function',
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
];

export const ALL_TOOL_NAMES = new Set(
  OFFICE_TOOLS.map((t) =>
    t.type === 'function' ? t.function.name : '',
  ).filter(Boolean),
);

export function toolsForMode(
  mode: ChatMode,
  buildingPlan = false,
  hasDocument = true,
  docType: DocType = 'SHEET',
): OpenAI.Chat.Completions.ChatCompletionTool[] {
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
        : [];
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
        : [];
  const plan = ['create_plan', 'update_plan'];
  let names: string[] = [];
  if (mode === 'chat') names = [...read];
  if (mode === 'plan') names = [...read, ...plan];
  if (mode === 'agent') {
    names = [...read, ...write];
    if (buildingPlan) names.push('update_plan');
  }
  return OFFICE_TOOLS.filter(
    (t) => t.type === 'function' && names.includes(t.function.name),
  );
}
