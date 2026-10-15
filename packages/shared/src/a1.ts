export function colLettersToIndex(col: string): number {
  let n = 0;
  const upper = col.toUpperCase();
  for (let i = 0; i < upper.length; i += 1) {
    const ch = upper[i];
    if (ch < 'A' || ch > 'Z') return -1;
    n = n * 26 + (ch.charCodeAt(0) - 64);
  }
  return n - 1;
}

export function indexToColLetters(col: number): string {
  let n = col + 1;
  let s = '';
  while (n > 0) {
    const r = (n - 1) % 26;
    s = String.fromCharCode(65 + r) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

/** Shift relative A1 refs in a formula by dRow/dCol (absolute $ refs stay put). */
export function adjustFormulaRefs(
  formula: string,
  dRow: number,
  dCol: number,
): string {
  if (!dRow && !dCol) return formula;
  return formula.replace(
    /(^|[^A-Za-z0-9_])(\$?)([A-Za-z]+)(\$?)(\d+)\b/g,
    (
      full,
      prefix: string,
      colAbs: string,
      col: string,
      rowAbs: string,
      row: string,
    ) => {
      let c = colLettersToIndex(col);
      let r = Number(row) - 1;
      if (!Number.isFinite(c) || c < 0 || !Number.isFinite(r) || r < 0)
        return full;
      if (!colAbs) c += dCol;
      if (!rowAbs) r += dRow;
      if (c < 0 || r < 0) return full;
      return `${prefix}${colAbs}${indexToColLetters(c)}${rowAbs}${r + 1}`;
    },
  );
}

/** Collect non-empty formulas from a sparse matrix (for set_formulas). */
export function collectFormulaCells(
  formulas: unknown[][],
  startRow: number,
  startCol: number,
): Array<{ row: number; col: number; formula: string }> {
  const cells: Array<{ row: number; col: number; formula: string }> = [];
  const numRows = formulas.length;
  const numCols = Math.max(
    ...formulas.map((row) => (Array.isArray(row) ? row.length : 0)),
    0,
  );
  for (let r = 0; r < numRows; r += 1) {
    const row = Array.isArray(formulas[r]) ? formulas[r] : [];
    for (let c = 0; c < numCols; c += 1) {
      const raw = row[c] == null ? '' : String(row[c]).trim();
      if (raw && raw !== '=') {
        cells.push({
          row: startRow + r,
          col: startCol + c,
          formula: raw.startsWith('=') ? raw : `=${raw}`,
        });
      }
    }
  }
  return cells;
}

export type PlanToolTodo = {
  id: string;
  content: string;
  status?: string;
};

export type PlanState = {
  name: string;
  overview: string;
  todos: PlanToolTodo[];
};

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
  plan: PlanState,
  todos: Array<{ id?: string; content?: string; status?: string }>,
): PlanState {
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
