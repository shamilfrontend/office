import {
  Injectable,
  NotFoundException,
  BadRequestException,
  PayloadTooLargeException,
} from '@nestjs/common';
import { ConfigService } from '@nestjs/config';
import type { User } from '@prisma/client';
import { randomUUID } from 'crypto';
import * as path from 'path';
import type { DocType } from '@office/shared';
import { PrismaService } from '../prisma/prisma.service';
import { UserStorageService } from '../storage/user-storage.service';
import type { MulterFile } from '../types/uploaded-file';
import { decodeUploadFilename } from './decode-filename';
import {
  docToDocx,
  docxToDoc,
  sheetToXlsx,
  xlsxToSheet,
} from './office-convert';

export type { DocType };
export type DownloadFormat = 'json' | 'csv' | 'xlsx' | 'docx';

@Injectable()
export class FilesService {
  constructor(
    private readonly prisma: PrismaService,
    private readonly storage: UserStorageService,
    private readonly config: ConfigService,
  ) {}

  list(userId: string) {
    return this.prisma.document.findMany({
      where: { userId },
      orderBy: { updatedAt: 'desc' },
      select: {
        id: true,
        name: true,
        type: true,
        size: true,
        updatedAt: true,
        createdAt: true,
      },
    });
  }

  async get(userId: string, id: string) {
    const doc = await this.prisma.document.findFirst({
      where: { id, userId },
    });
    if (!doc) throw new NotFoundException('Document not found');
    const snapshot = await this.storage.readDocument(userId, doc.relativePath);
    return { ...doc, snapshot: JSON.parse(snapshot) };
  }

  async create(user: User, type: DocType, name?: string) {
    const title = name?.trim() || 'Без названия';
    const fileName = `${randomUUID()}.json`;
    const relativePath = path.join('documents', fileName);
    // Empty object: Univer createWorkbook/createDocument fills defaults on the client.
    const snapshot = { id: randomUUID(), name: title };
    const content = JSON.stringify(snapshot);
    await this.ensureQuota(user.id, Buffer.byteLength(content));
    const { size } = await this.storage.writeDocument(
      user.id,
      relativePath,
      content,
    );
    return this.prisma.document.create({
      data: {
        userId: user.id,
        name: title,
        type,
        relativePath,
        size,
      },
    });
  }

  async save(userId: string, id: string, snapshot: unknown, name?: string) {
    const doc = await this.prisma.document.findFirst({
      where: { id, userId },
    });
    if (!doc) throw new NotFoundException('Document not found');
    const content = JSON.stringify(snapshot);
    await this.ensureQuota(userId, Buffer.byteLength(content) - doc.size);
    const { size } = await this.storage.writeDocument(
      userId,
      doc.relativePath,
      content,
    );
    return this.prisma.document.update({
      where: { id },
      data: {
        size,
        ...(name?.trim() ? { name: name.trim() } : {}),
      },
    });
  }

  async rename(userId: string, id: string, name: string) {
    const doc = await this.prisma.document.findFirst({
      where: { id, userId },
    });
    if (!doc) throw new NotFoundException('Document not found');
    if (!name.trim()) throw new BadRequestException('Name required');
    return this.prisma.document.update({
      where: { id },
      data: { name: name.trim() },
    });
  }

  async remove(userId: string, id: string) {
    const doc = await this.prisma.document.findFirst({
      where: { id, userId },
    });
    if (!doc) throw new NotFoundException('Document not found');
    await this.storage.deleteDocument(userId, doc.relativePath);
    await this.prisma.document.delete({ where: { id } });
    return { ok: true };
  }

  async uploadFromComputer(user: User, file: MulterFile, typeHint?: string) {
    if (!file) throw new BadRequestException('File required');
    const original = decodeUploadFilename(file.originalname || 'import.json');
    const lower = original.toLowerCase();
    let type: DocType = 'SHEET';
    let snapshot: unknown;
    const baseName = original.replace(/\.(json|csv|xlsx|docx)$/i, '');

    try {
      if (lower.endsWith('.xlsx')) {
        type = 'SHEET';
        snapshot = await xlsxToSheet(baseName, file.buffer);
      } else if (lower.endsWith('.docx')) {
        type = 'DOC';
        snapshot = await docxToDoc(baseName, file.buffer);
      } else if (lower.endsWith('.csv')) {
        type = 'SHEET';
        snapshot = this.csvToSheet(baseName, file.buffer.toString('utf8'));
      } else if (lower.endsWith('.json')) {
        const parsed = JSON.parse(file.buffer.toString('utf8')) as Record<
          string,
          unknown
        >;
        if (typeHint === 'DOC' || parsed.body || parsed.documentStyle) {
          type = 'DOC';
        } else {
          type = 'SHEET';
        }
        snapshot = parsed;
      } else {
        throw new BadRequestException(
          'Supported formats: .xlsx, .docx, .csv, .json',
        );
      }
    } catch (err) {
      if (err instanceof BadRequestException) throw err;
      throw new BadRequestException(
        err instanceof Error ? err.message : 'Failed to parse office file',
      );
    }

    const fileName = `${randomUUID()}.json`;
    const relativePath = path.join('documents', fileName);
    const content = JSON.stringify(snapshot);
    await this.ensureQuota(user.id, Buffer.byteLength(content));
    const { size } = await this.storage.writeDocument(
      user.id,
      relativePath,
      content,
    );
    return this.prisma.document.create({
      data: {
        userId: user.id,
        name: baseName,
        type,
        relativePath,
        size,
      },
    });
  }

  async download(userId: string, id: string, format: DownloadFormat = 'json') {
    const doc = await this.get(userId, id);
    const snapshot = doc.snapshot as Record<string, unknown>;

    if (format === 'xlsx') {
      if (doc.type !== 'SHEET') {
        throw new BadRequestException('XLSX only for sheets');
      }
      return {
        filename: `${doc.name}.xlsx`,
        contentType:
          'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        body: await sheetToXlsx(snapshot),
      };
    }

    if (format === 'docx') {
      if (doc.type !== 'DOC') {
        throw new BadRequestException('DOCX only for documents');
      }
      return {
        filename: `${doc.name}.docx`,
        contentType:
          'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
        body: await docToDocx(snapshot),
      };
    }

    if (format === 'csv') {
      if (doc.type !== 'SHEET') {
        throw new BadRequestException('CSV only for sheets');
      }
      return {
        filename: `${doc.name}.csv`,
        contentType: 'text/csv; charset=utf-8',
        body: this.sheetToCsv(snapshot),
      };
    }

    return {
      filename: `${doc.name}.json`,
      contentType: 'application/json; charset=utf-8',
      body: JSON.stringify(snapshot, null, 2),
    };
  }

  private csvToSheet(name: string, csv: string) {
    const rows = csv
      .split(/\r?\n/)
      .filter((line) => line.length > 0)
      .map((line) => this.parseCsvLine(line));
    const cellData: Record<string, Record<string, { v: string | number }>> = {};
    rows.forEach((row, r) => {
      cellData[r] = {};
      row.forEach((value, c) => {
        const num = Number(value);
        cellData[r][c] = {
          v:
            value !== '' && !Number.isNaN(num) && value.trim() !== ''
              ? num
              : value,
        };
      });
    });
    return {
      id: randomUUID(),
      name,
      sheetOrder: ['sheet-1'],
      sheets: {
        'sheet-1': {
          id: 'sheet-1',
          name: 'Sheet1',
          rowCount: Math.max(100, rows.length + 20),
          columnCount: Math.max(20, ...(rows.map((r) => r.length) as number[])),
          cellData,
        },
      },
    };
  }

  private parseCsvLine(line: string): string[] {
    const result: string[] = [];
    let current = '';
    let inQuotes = false;
    for (let i = 0; i < line.length; i += 1) {
      const ch = line[i];
      if (ch === '"') {
        if (inQuotes && line[i + 1] === '"') {
          current += '"';
          i += 1;
        } else {
          inQuotes = !inQuotes;
        }
      } else if (ch === ',' && !inQuotes) {
        result.push(current);
        current = '';
      } else {
        current += ch;
      }
    }
    result.push(current);
    return result;
  }

  private sheetToCsv(snapshot: Record<string, unknown>) {
    const sheets = snapshot.sheets as Record<
      string,
      {
        cellData?: Record<string, Record<string, { v?: unknown }>>;
      }
    >;
    const order =
      (snapshot.sheetOrder as string[]) ?? Object.keys(sheets ?? {});
    const sheet = sheets?.[order[0]];
    const cellData = sheet?.cellData ?? {};
    const rowKeys = Object.keys(cellData)
      .map(Number)
      .sort((a, b) => a - b);
    const lines = rowKeys.map((r) => {
      const row = cellData[String(r)] ?? {};
      const colKeys = Object.keys(row)
        .map(Number)
        .sort((a, b) => a - b);
      const maxCol = colKeys.length ? Math.max(...colKeys) : -1;
      const values: string[] = [];
      for (let c = 0; c <= maxCol; c += 1) {
        const v = row[String(c)]?.v;
        const raw = v == null ? '' : String(v);
        values.push(/[",\n]/.test(raw) ? `"${raw.replace(/"/g, '""')}"` : raw);
      }
      return values.join(',');
    });
    return lines.join('\n');
  }

  private async ensureQuota(userId: string, delta: number) {
    const quota = Number(
      this.config.get<string>('USER_QUOTA_BYTES') ?? 209715200,
    );
    const used = await this.storage.usageBytes(userId);
    if (used + Math.max(0, delta) > quota) {
      throw new PayloadTooLargeException('User storage quota exceeded');
    }
  }
}
