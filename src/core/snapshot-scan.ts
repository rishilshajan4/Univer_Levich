// Pure, React-free scans over a raw Univer `IWorkbookData` snapshot. Kept separate
// from LevichSheet so the open-at-# (feature #2) and empty-formula (feature #12
// companion to NO_CALCULATION) logic can be unit-tested without a DOM/engine.

type SnapCell = { v?: unknown; f?: unknown };
type SnapSheet = { cellData?: Record<string, Record<string, SnapCell>> };
type Snap = { sheetOrder?: string[]; sheets?: Record<string, SnapSheet> };

/** The active (first) sheet's cellData grid from a raw Univer snapshot (best-effort). */
export function activeCellData(snapshot: unknown): Record<string, Record<string, SnapCell>> | undefined {
  const s = snapshot as Snap | undefined;
  const order = s?.sheetOrder;
  if (!order?.length) return undefined;
  // Under the host's lazy SHELL-workbook model only the ACTIVE sheet carries real
  // cellData; every other sheet is an empty shell. `sheetOrder[0]` is NOT necessarily
  // the active sheet, so return the first sheet that actually has cell data (the
  // hydrated/active one). Falls back to sheetOrder[0] for a fully-populated snapshot.
  // This fixes open-at-# (findHashCell) and the empty-formula recompute for multi-sheet
  // files whose active sheet isn't index 0.
  for (const id of order) {
    const cd = s?.sheets?.[id]?.cellData;
    if (cd && Object.keys(cd).length > 0) return cd;
  }
  return s?.sheets?.[order[0]]?.cellData;
}

/** Matches Excel error tokens (`#REF!`, `#NAME?`, …) so the "#" anchor scan skips them. */
export const EXCEL_ERROR = /^#(REF|NAME|VALUE|DIV\/0|N\/A|NULL|NUM|SPILL|CALC|GETTING_DATA)!?/i;

/**
 * Feature #2 — find the first deliberate "#" anchor cell: a cell whose text value
 * starts with `#` but is NOT an Excel error token. Scans the in-memory snapshot
 * (cheap) rather than probing the grid via the Facade.
 */
export function findHashCell(snapshot: unknown): { row: number; column: number } | null {
  const cellData = activeCellData(snapshot);
  if (!cellData) return null;
  for (const [r, cols] of Object.entries(cellData)) {
    for (const [c, cell] of Object.entries(cols)) {
      const v = cell?.v;
      if (typeof v === "string") {
        const t = v.trim();
        // A deliberate review marker: "#" alone, or "#" followed by a space/letter
        // (e.g. "#review", "# note"). Excludes coded IDs like "#123"/"#4A" and Excel
        // error tokens (#REF!/#NAME? — caught by EXCEL_ERROR), which shouldn't be anchors.
        if (/^#($|\s|[a-zA-Z])/.test(t) && !EXCEL_ERROR.test(t)) return { row: Number(r), column: Number(c) };
      }
    }
  }
  return null;
}

/**
 * Feature #12 (companion to NO_CALCULATION) — formula cells that carry NO cached value
 * (Excel never computed them, e.g. a freshly exported sheet whose `=SUM` totals are
 * blank). Under NO_CALCULATION these would render empty, so we recompute ONLY these
 * post-load by re-setting their formula — never the cells that already have a cached
 * value (those are preserved verbatim, which is what keeps genuine zero totals intact).
 */
export function emptyFormulaCells(snapshot: unknown): Array<{ row: number; column: number; formula: string }> {
  const cellData = activeCellData(snapshot);
  if (!cellData) return [];
  const out: Array<{ row: number; column: number; formula: string }> = [];
  for (const [r, cols] of Object.entries(cellData)) {
    for (const [c, cell] of Object.entries(cols)) {
      const f = cell && typeof cell.f === "string" ? cell.f : "";
      const hasFormula = f.length > 0;
      const noCachedValue = cell?.v === undefined || cell?.v === null || cell?.v === "";
      // Do NOT recompute a cross-sheet formula on load. Under the lazy shell-workbook
      // model only the active sheet is hydrated; every other sheet is an empty shell,
      // so recomputing `='P&L'!B12` / `=SUM(Detail!A:A)` here resolves the reference
      // against nothing and renders #NAME?/#REF! (the "formula broken" regression).
      // A "!" in an Excel formula reliably marks a sheet-qualified reference; same-sheet
      // formulas never contain one, so those still get their blank totals filled in.
      const referencesOtherSheet = f.includes("!");
      if (hasFormula && noCachedValue && !referencesOtherSheet) out.push({ row: Number(r), column: Number(c), formula: f });
    }
  }
  return out;
}
