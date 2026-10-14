import { randomUUID } from 'crypto';
import ExcelJS from 'exceljs';
import mammoth from 'mammoth';
import {
  Document,
  Packer,
  Paragraph,
  TextRun,
} from 'docx';

type CellValue = string | number | boolean;

type SheetCell = {
  v?: CellValue;
  f?: string;
  t?: number;
};

type SheetSnapshot = {
  id: string;
  name: string;
  sheetOrder: string[];
  sheets: Record<
    string,
    {
      id: string;
      name: string;
      rowCount: number;
      columnCount: number;
      cellData: Record<string, Record<string, SheetCell>>;
    }
  >;
};

type DocSnapshot = {
  id: string;
  title: string;
  body: {
    dataStream: string;
    paragraphs: Array<{ startIndex: number; paragraphId: string }>;
    textRuns?: Array<{ st: number; ed: number; ts?: { bl?: number; fs?: number } }>;
  };
  documentStyle: {
    pageSize: { width: number; height: number };
    marginTop: number;
    marginBottom: number;
    marginLeft: number;
    marginRight: number;
  };
};

function cellFromExcel(value: ExcelJS.CellValue): SheetCell | null {
  if (value == null || value === '') return null;

  if (typeof value === 'object' && 'formula' in value) {
    const formula = String(value.formula ?? '').replace(/^=/, '');
    const result = (value as { result?: unknown }).result;
    const cell: SheetCell = { f: formula };
    if (typeof result === 'number' || typeof result === 'string' || typeof result === 'boolean') {
      cell.v = result;
    }
    return cell;
  }

  if (typeof value === 'object' && 'richText' in value) {
    const text = (value.richText ?? []).map((p) => p.text).join('');
    return text ? { v: text } : null;
  }

  if (typeof value === 'object' && 'text' in value) {
    const text = String((value as { text?: string }).text ?? '');
    return text ? { v: text } : null;
  }

  if (typeof value === 'object' && 'result' in value) {
    const result = (value as { result?: unknown }).result;
    if (typeof result === 'number' || typeof result === 'string' || typeof result === 'boolean') {
      return { v: result };
    }
    return null;
  }

  if (value instanceof Date) {
    return { v: value.toISOString() };
  }

  if (typeof value === 'number' || typeof value === 'string' || typeof value === 'boolean') {
    return { v: value };
  }

  return { v: String(value) };
}

export async function xlsxToSheet(
  name: string,
  buffer: Buffer,
): Promise<SheetSnapshot> {
  const workbook = new ExcelJS.Workbook();
  // exceljs typings expect NodeJS.Buffer; runtime accepts Buffer
  await workbook.xlsx.load(buffer as unknown as ExcelJS.Buffer);

  const sheetOrder: string[] = [];
  const sheets: SheetSnapshot['sheets'] = {};

  workbook.eachSheet((ws) => {
    const id = `sheet-${sheetOrder.length + 1}`;
    sheetOrder.push(id);
    const cellData: Record<string, Record<string, SheetCell>> = {};
    let maxRow = 0;
    let maxCol = 0;

    ws.eachRow({ includeEmpty: false }, (row, rowNumber) => {
      const r = rowNumber - 1;
      maxRow = Math.max(maxRow, r);
      row.eachCell({ includeEmpty: false }, (cell, colNumber) => {
        const c = colNumber - 1;
        maxCol = Math.max(maxCol, c);
        const mapped = cellFromExcel(cell.value);
        if (!mapped) return;
        if (!cellData[r]) cellData[r] = {};
        cellData[r][c] = mapped;
      });
    });

    sheets[id] = {
      id,
      name: ws.name || `Sheet${sheetOrder.length}`,
      rowCount: Math.max(100, maxRow + 21),
      columnCount: Math.max(20, maxCol + 5),
      cellData,
    };
  });

  if (!sheetOrder.length) {
    const id = 'sheet-1';
    sheetOrder.push(id);
    sheets[id] = {
      id,
      name: 'Sheet1',
      rowCount: 100,
      columnCount: 20,
      cellData: {},
    };
  }

  return {
    id: randomUUID(),
    name,
    sheetOrder,
    sheets,
  };
}

export async function sheetToXlsx(
  snapshot: Record<string, unknown>,
): Promise<Buffer> {
  const workbook = new ExcelJS.Workbook();
  const sheets = (snapshot.sheets ?? {}) as SheetSnapshot['sheets'];
  const order =
    (snapshot.sheetOrder as string[]) ?? Object.keys(sheets);

  for (const sheetId of order) {
    const src = sheets[sheetId];
    if (!src) continue;
    const ws = workbook.addWorksheet(src.name || 'Sheet1');
    const cellData = src.cellData ?? {};
    for (const [rKey, row] of Object.entries(cellData)) {
      const r = Number(rKey) + 1;
      for (const [cKey, cell] of Object.entries(row)) {
        const c = Number(cKey) + 1;
        const target = ws.getCell(r, c);
        if (cell.f) {
          target.value = {
            formula: cell.f.startsWith('=') ? cell.f.slice(1) : cell.f,
            result: cell.v as string | number | boolean | undefined,
          };
        } else if (cell.v != null) {
          target.value = cell.v;
        }
      }
    }
  }

  if (!workbook.worksheets.length) {
    workbook.addWorksheet('Sheet1');
  }

  const arrayBuffer = await workbook.xlsx.writeBuffer();
  return Buffer.from(arrayBuffer);
}

function buildDocBody(paragraphsText: string[]): DocSnapshot['body'] {
  const paragraphs: Array<{ startIndex: number; paragraphId: string }> = [];
  let dataStream = '';

  for (let i = 0; i < paragraphsText.length; i++) {
    const text = paragraphsText[i] ?? '';
    dataStream += text;
    const breakIndex = dataStream.length;
    dataStream += '\r';
    paragraphs.push({
      startIndex: breakIndex,
      paragraphId: `p-${i + 1}`,
    });
  }

  // Univer docs expect a section break `\n` at the end of the stream.
  dataStream += '\n';

  return { dataStream, paragraphs };
}

export async function docxToDoc(
  title: string,
  buffer: Buffer,
): Promise<DocSnapshot> {
  const result = await mammoth.extractRawText({ buffer });
  // mammoth separates paragraphs with blank lines (`\n\n`)
  const paragraphsText = result.value
    .replace(/\r\n/g, '\n')
    .split(/\n{2,}/)
    .map((block) => block.replace(/\r/g, '').replace(/\n/g, ' ').trimEnd())
    .filter((block, index, arr) => block !== '' || arr.length === 1);

  while (
    paragraphsText.length > 1 &&
    paragraphsText[paragraphsText.length - 1] === ''
  ) {
    paragraphsText.pop();
  }

  // Keep at least one empty paragraph so Univer can open the doc.
  const lines = paragraphsText.length ? paragraphsText : [''];
  const body = buildDocBody(lines);

  return {
    id: randomUUID(),
    title,
    body,
    documentStyle: {
      pageSize: { width: 595.3, height: 841.9 },
      marginTop: 72,
      marginBottom: 72,
      marginLeft: 90,
      marginRight: 90,
    },
  };
}

export async function docToDocx(
  snapshot: Record<string, unknown>,
): Promise<Buffer> {
  const body = snapshot.body as
    | { dataStream?: string }
    | undefined;
  const raw = body?.dataStream ?? '';
  // Split on Univer paragraph markers; drop trailing section `\n`.
  const parts = raw
    .replace(/\n$/, '')
    .split('\r')
    .map((p) => p.replace(/\n/g, ''));

  const children =
    parts.length > 0
      ? parts.map(
          (text) =>
            new Paragraph({
              children: [new TextRun(text)],
            }),
        )
      : [new Paragraph({ children: [new TextRun('')] })];

  const doc = new Document({
    sections: [
      {
        properties: {},
        children,
      },
    ],
  });

  return Packer.toBuffer(doc);
}
