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
