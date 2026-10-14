import ExcelJS from 'exceljs';
import mammoth from 'mammoth';
import JSZip from 'jszip';

const MAX_CHARS = 20_000;

export async function extractAttachmentText(
  filename: string,
  buffer: Buffer,
): Promise<string> {
  const ext = filename.split('.').pop()?.toLowerCase() ?? '';
  let text = '';
  if (ext === 'xlsx' || ext === 'xls') text = await xlsxToText(buffer);
  else if (ext === 'docx') text = await docxToText(buffer);
  else if (ext === 'pptx') text = await pptxToText(buffer);
  else if (ext === 'pdf') text = await pdfToText(buffer);
  else text = buffer.toString('utf8');
  return text.replace(/\u0000/g, '').slice(0, MAX_CHARS);
}

async function xlsxToText(buffer: Buffer) {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(buffer as unknown as ExcelJS.Buffer);
  const parts: string[] = [];
  workbook.eachSheet((sheet) => {
    parts.push(`# ${sheet.name}`);
    sheet.eachRow((row) => {
      const cells = (row.values as unknown[])
        .slice(1)
        .map((value) => cellText(value))
        .join('\t');
      if (cells.trim()) parts.push(cells);
    });
  });
  return parts.join('\n');
}

function cellText(value: unknown): string {
  if (value == null) return '';
  if (typeof value === 'object' && value && 'richText' in value) {
    return ((value as { richText?: Array<{ text?: string }> }).richText ?? [])
      .map((part) => part.text ?? '')
      .join('');
  }
  if (typeof value === 'object' && value && 'text' in value) {
    return String((value as { text?: string }).text ?? '');
  }
  if (typeof value === 'object' && value && 'result' in value) {
    return String((value as { result?: unknown }).result ?? '');
  }
  return String(value);
}

async function docxToText(buffer: Buffer) {
  const result = await mammoth.extractRawText({ buffer });
  return result.value;
}

async function pptxToText(buffer: Buffer) {
  const zip = await JSZip.loadAsync(buffer);
  const names = Object.keys(zip.files)
    .filter((name) => /ppt\/slides\/slide\d+\.xml$/.test(name))
    .sort();
  const slides: string[] = [];
  for (const name of names) {
    const xml = await zip.files[name].async('string');
    const texts = [...xml.matchAll(/<a:t[^>]*>([^<]*)<\/a:t>/g)].map(
      (match) => match[1],
    );
    if (texts.length) slides.push(texts.join(' '));
  }
  return slides.join('\n\n');
}

async function pdfToText(buffer: Buffer) {
  // CJS require keeps module.parent set, so pdf-parse does not read its debug fixture.
  // eslint-disable-next-line @typescript-eslint/no-require-imports
  const pdfParse = require('pdf-parse') as (
    data: Buffer,
  ) => Promise<{ text?: string }>;
  const parsed = await pdfParse(buffer);
  return parsed.text ?? '';
}
