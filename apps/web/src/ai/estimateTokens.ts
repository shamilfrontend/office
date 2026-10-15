/** Накладные токены на одно сообщение чата (роль, разделители). */
export const MESSAGE_FRAMING_TOKENS = 4;

/**
 * Оценка токенов с учётом Unicode без токенизатора модели.
 * ASCII ~4 симв./токен; кириллица/latin-ext плотнее; CJK ~1 симв./токен.
 */
export function estimateTextTokens(text: string): number {
  if (!text) return 0;
  let tokens = 0;
  for (let i = 0; i < text.length; i += 1) {
    const c = text.charCodeAt(i);
    if (c <= 0x7f) {
      tokens += 0.25;
    } else if (c <= 0x7ff) {
      tokens += 0.6;
    } else {
      tokens += 1;
    }
  }
  return Math.max(1, Math.ceil(tokens));
}

export function estimateMessageTokens(content: string): number {
  const textTokens = content ? estimateTextTokens(content) : 0;
  return textTokens + MESSAGE_FRAMING_TOKENS;
}

export function formatTokenCount(n: number): string {
  if (n >= 1_000_000) {
    const v = n / 1_000_000;
    return `${v >= 10 ? Math.round(v) : v.toFixed(1).replace(/\.0$/, '')}M`;
  }
  if (n >= 1_000) {
    const v = n / 1_000;
    return `${v >= 10 ? Math.round(v) : v.toFixed(1).replace(/\.0$/, '')}K`;
  }
  return String(Math.max(0, Math.round(n)));
}

/**
 * EMA-калибратор: подтягивает локальные оценки к `usage.prompt_tokens` шлюза.
 */
export class TokenEstimateCalibrator {
  private factor = 1;

  private readonly alpha: number;

  private readonly minFactor: number;

  private readonly maxFactor: number;

  constructor(opts?: {
    alpha?: number;
    minFactor?: number;
    maxFactor?: number;
  }) {
    this.alpha = opts?.alpha ?? 0.2;
    this.minFactor = opts?.minFactor ?? 0.5;
    this.maxFactor = opts?.maxFactor ?? 2.0;
  }

  getFactor(): number {
    return this.factor;
  }

  observe(estimatedPromptTokens: number, actualPromptTokens: number): void {
    if (estimatedPromptTokens < 50 || actualPromptTokens < 50) return;
    const ratio = actualPromptTokens / estimatedPromptTokens;
    if (!Number.isFinite(ratio) || ratio <= 0) return;
    const next = this.alpha * ratio + (1 - this.alpha) * this.factor;
    this.factor = Math.min(this.maxFactor, Math.max(this.minFactor, next));
  }

  apply(estimated: number): number {
    return Math.max(0, Math.round(estimated * this.factor));
  }
}

export const chatTokenCalibrator = new TokenEstimateCalibrator();
