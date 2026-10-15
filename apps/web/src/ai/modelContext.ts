type ModelContextInfo = {
  contextWindow: number;
  reservedOutputTokenSpace: number;
};

const DEFAULT: ModelContextInfo = {
  contextWindow: 128_000,
  reservedOutputTokenSpace: 4_096,
};

/** Правила по подстроке имени модели (длинные ключи проверяются первыми). */
const RULES: Array<{ match: RegExp; info: ModelContextInfo }> = [
  {
    match: /gemini-2\.5|gemini-2\.0|gemini-1\.5/i,
    info: { contextWindow: 1_048_576, reservedOutputTokenSpace: 8_192 },
  },
  {
    match: /claude-opus-4|claude-sonnet-4|claude-3-7|claude-3\.5|claude-3-5/i,
    info: { contextWindow: 200_000, reservedOutputTokenSpace: 8_192 },
  },
  {
    match: /claude-3-opus|claude-3-sonnet|claude-3-haiku/i,
    info: { contextWindow: 200_000, reservedOutputTokenSpace: 4_096 },
  },
  {
    match: /gpt-5|o3|o4-mini/i,
    info: { contextWindow: 200_000, reservedOutputTokenSpace: 16_384 },
  },
  {
    match: /gpt-4\.1|gpt-4o/i,
    info: { contextWindow: 128_000, reservedOutputTokenSpace: 16_384 },
  },
  {
    match: /gpt-4-turbo|gpt-4/i,
    info: { contextWindow: 128_000, reservedOutputTokenSpace: 4_096 },
  },
  {
    match: /deepseek/i,
    info: { contextWindow: 128_000, reservedOutputTokenSpace: 8_192 },
  },
  {
    match: /glm-4|glm-5|glm-/i,
    info: { contextWindow: 128_000, reservedOutputTokenSpace: 16_384 },
  },
  {
    match: /qwen.*235|qwen3/i,
    info: { contextWindow: 131_072, reservedOutputTokenSpace: 8_192 },
  },
  {
    match: /qwen/i,
    info: { contextWindow: 128_000, reservedOutputTokenSpace: 8_192 },
  },
  {
    match: /mistral|mixtral|devstral/i,
    info: { contextWindow: 128_000, reservedOutputTokenSpace: 8_192 },
  },
  {
    match: /grok/i,
    info: { contextWindow: 131_072, reservedOutputTokenSpace: 8_192 },
  },
  {
    match: /llama.*70|llama-3\.1|llama3\.1/i,
    info: { contextWindow: 128_000, reservedOutputTokenSpace: 8_192 },
  },
];

export function getModelContextInfo(
  model: string | null | undefined,
): ModelContextInfo {
  const name = (model || '').trim();
  if (!name) return DEFAULT;
  const matched = RULES.find((rule) => rule.match.test(name));
  return matched ? matched.info : DEFAULT;
}

export function effectiveReservedTokens(
  contextWindow: number,
  reserved: number,
): number {
  return Math.min(contextWindow, Math.max(0, reserved));
}
