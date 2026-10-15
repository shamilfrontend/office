function parseSsePart(part: string): Record<string, unknown> | null {
  const line = part.trim();
  if (!line.startsWith('data:')) return null;
  try {
    return JSON.parse(line.slice(5).trim()) as Record<string, unknown>;
  } catch {
    return null;
  }
}

/* eslint-disable no-await-in-loop, no-restricted-syntax -- ordered SSE stream processing */
export async function streamChat(
  body: Record<string, unknown>,
  onEvent: (ev: Record<string, unknown>) => void,
  signal: AbortSignal,
) {
  const res = await fetch('/ai/chat', {
    method: 'POST',
    credentials: 'include',
    signal,
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(body),
  });
  if (!res.ok || !res.body)
    throw new Error('Не удалось связаться с AI gateway');
  const reader = res.body.getReader();
  const decoder = new TextDecoder();
  let buffer = '';
  try {
    while (true) {
      const { done, value } = await reader.read();
      if (done) break;
      buffer += decoder.decode(value, { stream: true });
      const parts = buffer.split('\n\n');
      buffer = parts.pop() || '';
      for (const part of parts) {
        const event = parseSsePart(part);
        if (event) onEvent(event);
      }
    }
    buffer += decoder.decode();
    if (buffer.trim()) {
      const event = parseSsePart(buffer);
      if (event) onEvent(event);
    }
  } finally {
    try {
      await reader.cancel();
    } catch {
      /* ignore */
    }
  }
}
/* eslint-enable no-await-in-loop, no-restricted-syntax */
