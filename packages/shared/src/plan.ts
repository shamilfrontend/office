export type PlanTodo = {
  id: string;
  content: string;
  status?: string;
};

export type Plan = {
  name: string;
  overview: string;
  todos: PlanTodo[];
};

export function planToMarkdown(plan: Plan): string {
  const todos = plan.todos
    .map((todo) => {
      const done = todo.status === 'done' || todo.status === 'completed';
      return `- [${done ? 'x' : ' '}] ${todo.content}`;
    })
    .join('\n');
  return `# ${plan.name}\n\n${plan.overview}\n\n${todos}`.trim();
}

export function parsePlanMarkdown(markdown: string): Plan {
  const lines = markdown.replace(/\r\n/g, '\n').split('\n');
  let name = 'План';
  let i = 0;
  if (lines[0]?.startsWith('# ')) {
    name = lines[0].slice(2).trim() || 'План';
    i = 1;
    while (i < lines.length && !lines[i].trim()) i += 1;
  }
  const overviewLines: string[] = [];
  while (i < lines.length && !/^-\s+\[[ xX]\]\s+/.test(lines[i])) {
    overviewLines.push(lines[i]);
    i += 1;
  }
  const todos: PlanTodo[] = [];
  while (i < lines.length) {
    const match = lines[i].match(/^-\s+\[([ xX])\]\s+(.+)$/);
    if (match) {
      const done = match[1].toLowerCase() === 'x';
      todos.push({
        id: String(todos.length + 1),
        content: match[2].trim(),
        ...(done ? { status: 'completed' } : {}),
      });
    }
    i += 1;
  }
  return {
    name,
    overview: overviewLines.join('\n').trim(),
    todos,
  };
}

export function planAllTodosDone(plan: Plan | null | undefined): boolean {
  if (!plan?.todos.length) return false;
  return plan.todos.every(
    (t) => t.status === 'done' || t.status === 'completed',
  );
}

export function normalizePlanStatus(status: string): string {
  const s = status.trim().toLowerCase();
  if (s === 'done' || s === 'complete' || s === 'completed' || s === 'x') {
    return 'completed';
  }
  if (s === 'pending' || s === 'todo' || s === 'open' || s === '')
    return 'pending';
  if (s === 'in_progress' || s === 'doing' || s === 'active')
    return 'in_progress';
  if (s === 'cancelled' || s === 'canceled' || s === 'skipped')
    return 'cancelled';
  return s;
}

/** Pure merge used by update_plan (id match, else index). */
export function mergePlanTodos(
  plan: Plan,
  todos: Array<{ id?: string; content?: string; status?: string }>,
): Plan {
  return {
    ...plan,
    todos: plan.todos.map((t, i) => {
      const byId = todos.find(
        (x) => x.id != null && String(x.id) === String(t.id),
      );
      const byIndex = byId ? undefined : todos[i];
      const u = byId ?? byIndex;
      if (!u) return t;
      return {
        ...t,
        ...(u.content != null ? { content: String(u.content) } : {}),
        ...(u.status != null
          ? { status: normalizePlanStatus(String(u.status)) }
          : {}),
      };
    }),
  };
}
