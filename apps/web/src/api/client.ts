const JSON_HEADERS = { 'Content-Type': 'application/json' }

async function request<T>(path: string, init?: RequestInit): Promise<T> {
  const res = await fetch(path, {
    credentials: 'include',
    ...init,
    headers: {
      ...(init?.body instanceof FormData ? {} : JSON_HEADERS),
      ...init?.headers,
    },
  })
  if (!res.ok) {
    let message = res.statusText
    try {
      const data = (await res.json()) as { message?: string | string[] }
      message = Array.isArray(data.message)
        ? data.message.join(', ')
        : data.message || message
    } catch {
      /* ignore */
    }
    throw new Error(message)
  }
  if (res.status === 204) return undefined as T
  return res.json() as Promise<T>
}

export type Me = {
  id: string
  email: string | null
  displayName: string | null
  avatarUrl: string | null
  onboardingSkipped: boolean
  hasAiProvider: boolean
  providers: Array<{
    id: string
    provider: string
    baseUrl: string
    model: string | null
    enabled: boolean
  }>
}

export type DocItem = {
  id: string
  name: string
  type: 'SHEET' | 'DOC'
  size: number
  updatedAt: string
  createdAt: string
}

export const api = {
  authConfig: () => request<{ yandexEnabled: boolean }>('/auth/config'),
  me: () => request<Me>('/auth/me'),
  logout: () => request<{ ok: boolean }>('/auth/logout', { method: 'POST' }),
  skipOnboarding: () =>
    request<Me>('/users/onboarding', {
      method: 'PATCH',
      body: JSON.stringify({ skipped: true }),
    }),
  listFiles: () => request<DocItem[]>('/files'),
  getFile: (id: string) =>
    request<DocItem & { snapshot: unknown }>(`/files/${id}`),
  createFile: (type: 'SHEET' | 'DOC', name?: string) =>
    request<DocItem>('/files', {
      method: 'POST',
      body: JSON.stringify({ type, name }),
    }),
  saveFile: (id: string, snapshot: unknown, name?: string) =>
    request<DocItem>(`/files/${id}`, {
      method: 'PUT',
      body: JSON.stringify({ snapshot, name }),
    }),
  renameFile: (id: string, name: string) =>
    request<DocItem>(`/files/${id}`, {
      method: 'PATCH',
      body: JSON.stringify({ name }),
    }),
  deleteFile: (id: string) =>
    request<{ ok: boolean }>(`/files/${id}`, { method: 'DELETE' }),
  uploadFile: async (file: File, type?: string) => {
    const form = new FormData()
    form.append('file', file)
    if (type) form.append('type', type)
    return request<DocItem>('/files/upload', { method: 'POST', body: form })
  },
  downloadUrl: (
    id: string,
    format: 'json' | 'csv' | 'xlsx' | 'docx' = 'json',
  ) => `/files/${id}/download?format=${format}`,
  aiDefaults: () => request<Record<string, unknown>>('/ai/defaults'),
  listProviders: () =>
    request<
      Array<{
        id: string
        provider: string
        baseUrl: string
        model: string | null
        enabled: boolean
        hasApiKey: boolean
      }>
    >('/ai/providers'),
  upsertProvider: (body: Record<string, unknown>) =>
    request('/ai/providers', { method: 'POST', body: JSON.stringify(body) }),
  deleteProvider: (id: string) =>
    request(`/ai/providers/${id}`, { method: 'DELETE' }),
  listProviderModels: (id: string) =>
    request<{ models: string[] }>(`/ai/providers/${id}/models`),
  testProvider: (body: Record<string, unknown>) =>
    request<{ ok: boolean; models: string[]; error?: string }>(
      '/ai/providers/test',
      { method: 'POST', body: JSON.stringify(body) },
    ),
  listOpenChats: (documentId: string) =>
    request<ChatRecord[]>(
      `/ai/chats?open=1&documentId=${encodeURIComponent(documentId)}`,
    ),
  listClosedChats: (documentId: string) =>
    request<ChatSummary[]>(
      `/ai/chats?open=0&documentId=${encodeURIComponent(documentId)}`,
    ),
  createChat: (body: {
    mode?: string
    providerId?: string
    model?: string
    documentId?: string
  }) =>
    request<ChatRecord>('/ai/chats', {
      method: 'POST',
      body: JSON.stringify(body),
    }),
  updateChat: (id: string, body: Record<string, unknown>) =>
    request<ChatRecord>(`/ai/chats/${id}`, {
      method: 'PATCH',
      body: JSON.stringify(body),
    }),
  closeChat: (id: string) =>
    request<{ ok: boolean; chat: ChatRecord | null }>(`/ai/chats/${id}/close`, {
      method: 'POST',
    }),
  deleteChat: (id: string) =>
    request<{ ok: boolean; chat: ChatRecord | null }>(`/ai/chats/${id}`, {
      method: 'DELETE',
    }),
  addChatMessage: (id: string, body: Record<string, unknown>) =>
    request<ChatMessageRecord>(`/ai/chats/${id}/messages`, {
      method: 'POST',
      body: JSON.stringify(body),
    }),
  extractAttachment: async (file: File) => {
    const form = new FormData()
    form.append('file', file)
    return request<{ name: string; text: string }>('/ai/attachments/extract', {
      method: 'POST',
      body: form,
    })
  },
}

export type ChatMessageRecord = {
  id: string
  role: string
  content: string
  toolName: string | null
  toolCallId: string | null
  toolStatus: string | null
  toolCallsJson: string | null
  createdAt: string
}

export type ChatSummary = {
  id: string
  title: string
  mode: string
  providerId: string | null
  model: string | null
  documentId: string | null
  planTitle: string | null
  planMarkdown: string | null
  planExecuted: boolean | null
  pinnedOpen: boolean
  sortOrder: number
  createdAt: string
  updatedAt: string
}

export type ChatRecord = ChatSummary & {
  messages: ChatMessageRecord[]
}
