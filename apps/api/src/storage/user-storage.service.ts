import {
  Injectable,
  ForbiddenException,
  OnModuleInit,
} from '@nestjs/common';
import { ConfigService } from '@nestjs/config';
import * as fs from 'fs/promises';
import * as path from 'path';

@Injectable()
export class UserStorageService implements OnModuleInit {
  private dataRoot!: string;

  constructor(private readonly config: ConfigService) {}

  async onModuleInit() {
    const configured = this.config.get<string>('DATA_ROOT') ?? '../../data';
    this.dataRoot = path.resolve(process.cwd(), configured);
    await fs.mkdir(this.dataRoot, { recursive: true });
  }

  getDataRoot() {
    return this.dataRoot;
  }

  userRoot(userId: string) {
    return path.join(this.dataRoot, 'users', userId);
  }

  documentsDir(userId: string) {
    return path.join(this.userRoot(userId), 'documents');
  }

  async ensureUserFolder(userId: string) {
    const docs = this.documentsDir(userId);
    await fs.mkdir(docs, { recursive: true });
    return this.userRoot(userId);
  }

  resolveSafe(userId: string, relativePath: string) {
    const root = this.userRoot(userId);
    const resolved = path.resolve(root, relativePath);
    const normalizedRoot = path.resolve(root);
    if (
      resolved !== normalizedRoot &&
      !resolved.startsWith(normalizedRoot + path.sep)
    ) {
      throw new ForbiddenException('Invalid path');
    }
    return resolved;
  }

  async writeDocument(
    userId: string,
    relativePath: string,
    content: string | Buffer,
  ) {
    const full = this.resolveSafe(userId, relativePath);
    await fs.mkdir(path.dirname(full), { recursive: true });
    await fs.writeFile(full, content);
    const stat = await fs.stat(full);
    return { full, size: stat.size };
  }

  async readDocument(userId: string, relativePath: string) {
    const full = this.resolveSafe(userId, relativePath);
    return fs.readFile(full, 'utf8');
  }

  async deleteDocument(userId: string, relativePath: string) {
    const full = this.resolveSafe(userId, relativePath);
    await fs.rm(full, { force: true });
  }

  async usageBytes(userId: string) {
    const root = this.userRoot(userId);
    return this.dirSize(root);
  }

  private async dirSize(dir: string): Promise<number> {
    try {
      const entries = await fs.readdir(dir, { withFileTypes: true });
      let total = 0;
      for (const entry of entries) {
        const full = path.join(dir, entry.name);
        if (entry.isDirectory()) {
          total += await this.dirSize(full);
        } else if (entry.isFile()) {
          total += (await fs.stat(full)).size;
        }
      }
      return total;
    } catch {
      return 0;
    }
  }
}
