import { computed, ref, type ComputedRef } from 'vue';
import { parsePlanMarkdown, planAllTodosDone } from '@office/shared';
import { titleFromMarkdown } from '../formatters';
import type { Tab } from '../types';

export function useChatPlan(options: {
  active: ComputedRef<Tab | null>;
  patchTab: (tab: Tab, body: Record<string, unknown>) => Promise<void>;
  send: (preset?: { display: string; wire: string }) => Promise<void>;
}) {
  const { active, patchTab, send } = options;

  const planOpen = ref(false);
  const planDraft = ref('');

  const planOverview = computed(() => {
    const tab = active.value;
    if (!tab?.planMarkdown) return '';
    if (tab.plan?.overview) return tab.plan.overview.slice(0, 160);
    const body = tab.planMarkdown.replace(/^#.*\n+/, '');
    return body.split('\n').filter(Boolean).slice(0, 2).join(' ').slice(0, 160);
  });

  function openPlan() {
    if (!active.value?.planMarkdown) return;
    planDraft.value = active.value.planMarkdown;
    planOpen.value = true;
  }

  async function savePlan() {
    const tab = active.value;
    if (!tab) return;
    tab.planMarkdown = planDraft.value;
    tab.planTitle = titleFromMarkdown(planDraft.value, tab.planTitle || 'План');
    tab.plan = parsePlanMarkdown(planDraft.value);
    tab.planEdited = true;
    tab.planRunState = 'idle';
    planOpen.value = false;
    await patchTab(tab, {
      planTitle: tab.planTitle,
      planMarkdown: tab.planMarkdown,
      planExecuted: false,
    });
  }

  async function runPlan() {
    const tab = active.value;
    if (!tab?.planMarkdown || tab.streaming) return;
    if (!tab.plan) tab.plan = parsePlanMarkdown(tab.planMarkdown);
    const title = tab.planTitle || tab.plan.name || 'План';
    tab.mode = 'agent';
    tab.executingPlan = true;
    tab.planRunState = 'running';
    await patchTab(tab, { mode: 'agent' });
    try {
      await send({
        display: `Сделать по плану «${title}»`,
        wire: `Сделать по плану «${title}»\n\n---\n${tab.planMarkdown}`,
      });
      const stopped =
        tab.retryReason === 'Остановлено' || tab.retryReason === 'Отклонено';
      if (stopped) {
        tab.planRunState = planAllTodosDone(tab.plan) ? 'done' : 'idle';
      } else {
        tab.planRunState = 'done';
        void patchTab(tab, { planExecuted: true });
      }
    } finally {
      tab.executingPlan = false;
      if (tab.planRunState === 'running') {
        tab.planRunState = planAllTodosDone(tab.plan) ? 'done' : 'idle';
      }
    }
  }

  return {
    planOpen,
    planDraft,
    planOverview,
    openPlan,
    savePlan,
    runPlan,
  };
}
