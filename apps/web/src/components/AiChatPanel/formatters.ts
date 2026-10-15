import type { ChatMode } from '@/univer/officeTools';
import type { LoopStopReason } from './types';

export function retryLabel(reason: LoopStopReason): string | null {
  if (reason === 'rejected') return 'Отклонено';
  if (reason === 'aborted') return 'Остановлено';
  if (reason === 'editor') return 'Редактор не готов';
  if (reason === 'tools') return 'Стоп по ошибкам tools';
  if (reason === 'stream') return 'Ошибка сети';
  if (reason === 'loop') return 'Зацикливание tools';
  return null;
}

export function formatHistoryDate(value: string) {
  try {
    return new Date(value).toLocaleString('ru-RU', {
      day: '2-digit',
      month: 'short',
      hour: '2-digit',
      minute: '2-digit',
    });
  } catch {
    return value;
  }
}

export function visibleUser(text: string) {
  const idx = text.indexOf('\n\n---\n');
  return idx === -1 ? text : text.slice(0, idx);
}

export function asMode(value: string): ChatMode {
  if (value === 'chat' || value === 'plan' || value === 'agent') return value;
  return 'agent';
}

export function modeLabel(mode: ChatMode) {
  if (mode === 'plan') return 'План';
  if (mode === 'chat') return 'Чат';
  return 'Агент';
}

export function compact(n: number) {
  if (n >= 1000) return `${Math.max(1, Math.round(n / 1000))}K`;
  return String(Math.max(0, Math.round(n)));
}

export function titleFromMarkdown(markdown: string, fallback: string) {
  const heading = markdown.match(/^#\s+(.+)$/m);
  return (heading?.[1] || fallback).trim() || 'План';
}

export function shortTitle(text: string) {
  const clean = text.replace(/\s+/g, ' ').trim();
  if (!clean) return 'Новый чат';
  return clean.length > 22 ? `${clean.slice(0, 22)}…` : clean;
}

export function toolResultStatus(result: string): 'ok' | 'error' {
  try {
    const parsed = JSON.parse(result) as { error?: unknown };
    if (parsed && typeof parsed === 'object' && parsed.error != null)
      return 'error';
  } catch {
    /* plain text result */
  }
  return 'ok';
}
