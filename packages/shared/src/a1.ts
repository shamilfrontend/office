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

export function parseA1Cell(a1: string): { row: number; col: number } | null {
  const matched = String(a1 || '')
    .trim()
    .replace(/^[^!]*!/, '')
    .match(/^\$?([A-Za-z]+)\$?(\d+)$/);
  if (!matched) return null;
  return { col: colLettersToIndex(matched[1]), row: Number(matched[2]) - 1 };
}

/** Estimate cell count of an A1 range (e.g. A1:C10 → 30). */
export function estimateA1Area(a1: string): number {
  const clean = String(a1 || '')
    .trim()
    .replace(/^[^!]*!/, '');
  if (!clean) return 0;
  const parts = clean.split(':');
  const start = parseA1Cell(parts[0]);
  if (!start) return 0;
  const end = parts[1] ? parseA1Cell(parts[1]) : start;
  if (!end) return 0;
  return (
    (Math.abs(end.row - start.row) + 1) * (Math.abs(end.col - start.col) + 1)
  );
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
