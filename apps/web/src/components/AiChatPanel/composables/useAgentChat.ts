import { nextTick } from 'vue';
import { api } from '@/api/client';
import { describeImpact } from '@/univer/officeTools';
import { retryLabel, shortTitle } from '../formatters';
import type { Tab } from '../types';
import { runAgentLoop } from './runAgentLoop';
import type { UseAgentChatOptions } from './agentTypes';

export type {
  AgentChatEmit,
  AgentChatProps,
  UseAgentChatOptions,
} from './agentTypes';

export function useAgentChat(options: UseAgentChatOptions) {
  const {
    props,
    emit,
    active,
    hasProvider,
    providers,
    attachError,
    approveExpanded,
    findTab,
    patchTab,
    pushSessionLog,
    toggleModelMenu,
    resizeComposer,
    refreshContextUsage,
    clearStreamHtml,
    flushStreamHtml,
    scheduleStreamHtml,
  } = options;

  let lastTokenRateAt = 0;

  function askApprove(tab: Tab, name: string, args: Record<string, unknown>) {
    return new Promise<boolean>((resolve) => {
      const target = findTab(tab.id) ?? tab;
      if (target.abort?.signal.aborted) {
        resolve(false);
        return;
      }
      let settled = false;
      function onAbort() {
        finish(false);
      }
      function finish(ok: boolean) {
        if (settled) return;
        settled = true;
        target.abort?.signal.removeEventListener('abort', onAbort);
        if (target.pendingApprove) target.pendingApprove = null;
        approveExpanded.value = false;
        resolve(ok);
      }
      target.abort?.signal.addEventListener('abort', onAbort, { once: true });
      const summary = describeImpact(name, args, props.univerApi ?? null);
      target.pendingApprove = { name, args, summary, resolve: finish };
    });
  }

  function resolveApprove(ok: boolean, always = false) {
    const tab = active.value;
    if (tab?.pendingApprove) {
      if (always && ok) tab.allowedTools.add(tab.pendingApprove.name);
      let decision = 'rejected';
      if (always && ok) decision = 'always';
      else if (ok) decision = 'accepted';
      pushSessionLog(tab, {
        kind: 'system',
        name: 'approve',
        content: JSON.stringify({
          tool: tab.pendingApprove.name,
          args: tab.pendingApprove.args,
          decision,
        }),
        status: ok ? 'ok' : 'error',
      });
    }
    tab?.pendingApprove?.resolve(ok);
    if (tab) tab.pendingApprove = null;
    approveExpanded.value = false;
  }

  function agentLoop(tabId: string) {
    return runAgentLoop({
      tabId,
      props,
      emit,
      providers: providers.value,
      findTab,
      patchTab,
      pushSessionLog,
      askApprove,
      flushStreamHtml,
      scheduleStreamHtml,
      refreshContextUsage,
      getLastTokenRateAt: () => lastTokenRateAt,
      setLastTokenRateAt: (value) => {
        lastTokenRateAt = value;
      },
    });
  }

  async function send(preset?: { display: string; wire: string }) {
    const tab = active.value;
    if (!tab || tab.streaming || !hasProvider.value) return;
    const display = (preset?.display ?? tab.draft).trim();
    if (!display) return;
    if (!tab.model) {
      attachError.value = 'Выберите модель в чате ассистента';
      void toggleModelMenu();
      return;
    }
    const contextParts: string[] = [];
    if (tab.includeDoc && props.docName) {
      contextParts.push(
        `Текущий документ: ${props.docName} (${props.docType})`,
      );
      const sel = props.getSelectionContext?.();
      if (sel) contextParts.push(sel);
    }
    if (!preset) {
      tab.attachments.forEach((file) => {
        contextParts.push(`Файл ${file.name}:\n${file.text}`);
      });
    }
    const contextBlock = contextParts.length
      ? `\n\n---\nКонтекст:\n${contextParts.join('\n\n')}`
      : '';
    let wire = display;
    if (preset?.wire) {
      wire = `${preset.wire}${tab.includeDoc && contextParts.length ? contextBlock : ''}`;
    } else if (contextParts.length) {
      wire = `${display}${contextBlock}`;
    }

    tab.messages.push({
      key: crypto.randomUUID(),
      role: 'user',
      content: display,
      wire,
    });
    if (tab.title === 'Новый чат') {
      tab.title = shortTitle(display);
      void patchTab(tab, { title: tab.title });
    }
    pushSessionLog(tab, { kind: 'user', content: wire });
    tab.draft = '';
    tab.attachments = [];
    void nextTick(resizeComposer);
    tab.streaming = true;
    tab.tokenRate = null;
    tab.tokenMeta = null;
    tab.showCheckpoint = false;
    tab.wroteSinceCheckpoint = false;
    tab.canRetry = false;
    tab.retryReason = null;
    clearStreamHtml();
    lastTokenRateAt = 0;
    const controller = new AbortController();
    tab.abort = controller;
    if (props.univerApi) emit('checkpoint');
    void api.addChatMessage(tab.id, { role: 'user', content: wire });

    try {
      const reason = await agentLoop(tab.id);
      const label = retryLabel(reason);
      if (label) {
        tab.canRetry = true;
        tab.retryReason = label;
      }
    } catch (e) {
      const aborted = e instanceof DOMException && e.name === 'AbortError';
      if (!aborted) {
        const content = e instanceof Error ? e.message : 'Ошибка';
        tab.messages.push({
          key: crypto.randomUUID(),
          role: 'assistant',
          content,
        });
        pushSessionLog(tab, {
          kind: 'system',
          name: 'error',
          content,
          status: 'error',
        });
        tab.canRetry = true;
        tab.retryReason = 'Ошибка сети';
      } else {
        pushSessionLog(tab, {
          kind: 'system',
          name: 'abort',
          content: 'Остановлено пользователем',
          status: 'error',
        });
        tab.canRetry = true;
        tab.retryReason = 'Остановлено';
      }
    } finally {
      tab.showCheckpoint = tab.wroteSinceCheckpoint && Boolean(props.univerApi);
      tab.streaming = false;
      tab.abort = null;
      clearStreamHtml();
    }
  }

  async function retryTurn() {
    const tab = active.value;
    if (!tab || tab.streaming || !tab.canRetry) return;
    tab.canRetry = false;
    tab.retryReason = null;
    tab.streaming = true;
    tab.tokenRate = null;
    tab.tokenMeta = null;
    tab.showCheckpoint = false;
    clearStreamHtml();
    lastTokenRateAt = 0;
    const controller = new AbortController();
    tab.abort = controller;
    if (props.univerApi) emit('checkpoint');
    try {
      const reason = await agentLoop(tab.id);
      const label = retryLabel(reason);
      if (label) {
        tab.canRetry = true;
        tab.retryReason = label;
      }
    } catch (e) {
      const aborted = e instanceof DOMException && e.name === 'AbortError';
      if (!aborted) {
        const content = e instanceof Error ? e.message : 'Ошибка';
        tab.messages.push({
          key: crypto.randomUUID(),
          role: 'assistant',
          content,
        });
        pushSessionLog(tab, {
          kind: 'system',
          name: 'error',
          content,
          status: 'error',
        });
        tab.canRetry = true;
        tab.retryReason = 'Ошибка сети';
      } else {
        tab.canRetry = true;
        tab.retryReason = 'Остановлено';
      }
    } finally {
      tab.showCheckpoint = tab.wroteSinceCheckpoint && Boolean(props.univerApi);
      tab.streaming = false;
      tab.abort = null;
      clearStreamHtml();
    }
  }

  function stop() {
    const tab = active.value;
    if (!tab) return;
    tab.abort?.abort();
    if (tab.pendingApprove) {
      tab.pendingApprove.resolve(false);
      tab.pendingApprove = null;
      approveExpanded.value = false;
    }
  }

  return {
    send,
    retryTurn,
    stop,
    askApprove,
    resolveApprove,
    clearLastTokenRate() {
      lastTokenRateAt = 0;
    },
  };
}
