export type ProviderKind =
  'DEEPSEEK' | 'OLLAMA' | 'LMSTUDIO' | 'OPENAI_COMPATIBLE';

export const PROVIDER_DEFAULTS: Record<
  ProviderKind,
  { baseUrl: string; models: string[]; needsKey: boolean }
> = {
  DEEPSEEK: {
    baseUrl: 'https://api.deepseek.com/v1',
    models: [],
    needsKey: true,
  },
  OLLAMA: {
    baseUrl: 'http://127.0.0.1:11434/v1',
    models: [],
    needsKey: false,
  },
  LMSTUDIO: {
    baseUrl: 'http://127.0.0.1:1234/v1',
    models: [],
    needsKey: false,
  },
  OPENAI_COMPATIBLE: {
    baseUrl: 'http://127.0.0.1:8080/v1',
    models: [],
    needsKey: false,
  },
};

export const LLM_ERROR_HINTS: Record<number, string> = {
  401: 'Ключ отклонён. Обновите API-ключ в Настройках → Провайдеры.',
  408: 'Провайдер не успел ответить. Повторите запрос позже.',
  429: 'Лимит запросов. Подождите и попробуйте снова.',
};
