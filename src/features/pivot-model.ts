/**
 * Interactive pivot engine (Excel / Google-Sheets grade), free-tier — no Univer Pro.
 *
 * `computePivotModel` buckets a `PivotSource` into a nested ROW tree × flat COLUMN
 * leaves × multiple value fields, computing per-group aggregates, per-level SUBTOTALS
 * and GRAND totals. Crucially, subtotals/totals aggregate the UNDERLYING values (not a
 * sum of child aggregates) so `average`/`min`/`max` match Excel exactly.
 *
 * `renderPivotModel` walks that tree into styled `cellData`, honouring collapse,
 * compact-vs-tabular layout, per-value number formats, and indent (via the same `pd.l`
 * left-padding used for imported pivots).
 */
import { ALIGN_RIGHT, NUMBER_PATTERN } from "./formatting";
import type { Cell, CellStyle, PivotAggregate, PivotModel, PivotNode, PivotSource, PivotSpec, PivotValueField } from "../core/types";

const SEP = "␟"; // ␟ — a path separator that won't collide with real field values.
// Dedicated colPath for the row-Total column. A NUL byte can't appear in a stringified
// cell value, so this never collides with a real (or blank) column-field path.
export const ROW_TOTAL = "\u0000TOTAL";

function aggregate(values: number[], agg: PivotAggregate): number {
  const nums = values;
  switch (agg) {
    case "count":
      return nums.length;
    case "countNumbers":
      return nums.filter((n) => Number.isFinite(n)).length;
    case "average": {
      const f = nums.filter((n) => Number.isFinite(n));
      return f.length ? f.reduce((s, x) => s + x, 0) / f.length : 0;
    }
    case "min": {
      // Exclude non-numbers (text / blanks) like Excel's MIN — else one stray "N/A"
      // poisons the whole group to NaN.
      const f = nums.filter(Number.isFinite);
      return f.length ? Math.min(...f) : 0;
    }
    case "max": {
      const f = nums.filter(Number.isFinite);
      return f.length ? Math.max(...f) : 0;
    }
    case "sum":
    default:
      return nums.reduce((s, x) => s + (Number.isFinite(x) ? x : 0), 0);
  }
}

/* ─── Mergeable accumulators ──────────────────────────────────────────────────
   To make deep pivots fast we compute a per-(col,value) accumulator ONCE at each
   LEAF and then ROLL UP bottom-up: a parent's accumulator is the O(children)
   merge of its children's accumulators — never a re-scan of all descendant leaves.
   Each aggregate keeps just enough SUFFICIENT STATISTICS to be exact after
   merging (sum keeps a running sum; count keeps n; average keeps {sum,n} over
   finite values; min/max keep the running extreme; countNumbers keeps the
   finite-count). This yields the SAME result as scanning the raw union of values
   (so average totals = avg of ALL underlying values, min/max ignore non-numbers),
   at ~O(rows × depth) instead of ~O(rows × depth × cols × values). */
interface Acc {
  sum: number; // Σ of finite values (for "sum" / "average").
  n: number; // total observations (for "count").
  fn: number; // finite-value count (for "countNumbers" / "average").
  min: number; // running min over finite values (Infinity if none seen).
  max: number; // running max over finite values (-Infinity if none seen).
}
const newAcc = (): Acc => ({ sum: 0, n: 0, fn: 0, min: Infinity, max: -Infinity });

/** Fold one raw value into an accumulator. */
function pushAcc(a: Acc, x: number): void {
  a.n += 1;
  if (Number.isFinite(x)) {
    a.sum += x;
    a.fn += 1;
    if (x < a.min) a.min = x;
    if (x > a.max) a.max = x;
  }
}

/** Merge `src` INTO `dst` in O(1) — associative + commutative, so roll-up order
 *  doesn't matter and a parent = merge of its children = merge of all its leaves. */
function mergeAcc(dst: Acc, src: Acc): void {
  dst.sum += src.sum;
  dst.n += src.n;
  dst.fn += src.fn;
  if (src.min < dst.min) dst.min = src.min;
  if (src.max > dst.max) dst.max = src.max;
}

/** Read the final aggregate out of an accumulator (matches `aggregate()` exactly). */
function readAcc(a: Acc, agg: PivotAggregate): number {
  switch (agg) {
    case "count":
      return a.n;
    case "countNumbers":
      return a.fn;
    case "average":
      return a.fn ? a.sum / a.fn : 0;
    case "min":
      return a.fn ? a.min : 0;
    case "max":
      return a.fn ? a.max : 0;
    case "sum":
    default:
      return a.sum;
  }
}

/** Default header label for a value field, e.g. "Sum of Amount". */
export function valueLabel(v: PivotValueField): string {
  if (v.label) return v.label;
  const verb: Record<PivotAggregate, string> = {
    sum: "Sum",
    count: "Count",
    countNumbers: "Count",
    average: "Average",
    min: "Min",
    max: "Max",
  };
  return `${verb[v.aggregate]} of ${v.field}`;
}

const num = (x: unknown): number => {
  const n = typeof x === "number" ? x : Number(x);
  return Number.isFinite(n) ? n : NaN;
};

/** Compute the full pivot tree from a source + spec. */
export function computePivotModel(source: PivotSource, spec: PivotSpec): PivotModel {
  const values = spec.values.length ? spec.values : [{ field: source.fields[0] ?? "value", aggregate: "count" as PivotAggregate }];

  // 1. Filter rows.
  const filters = spec.filters ?? [];
  const rows = source.rows.filter((r) => filters.every((f) => !f.include || f.include.includes(String(r[f.field] ?? ""))));

  const nValues = values.length;
  // Cell key: `${colPath}${SEP}${vi}`. The row Total column uses a DEDICATED sentinel
  // colPath (ROW_TOTAL) that can never equal a real column path — including the ""
  // path produced by a column field whose VALUE is blank — so the two never collide.
  const cellKey = (colPath: string, vi: number) => `${colPath}${SEP}${vi}`;

  // 2. Per-LEAF accumulators, computed ONCE per row: rowLeafPath → colLeafPath →
  //    valueIndex → Acc (sufficient statistics). This is the single scan of the
  //    underlying data; everything above rolls these UP without re-scanning.
  //    An `AccGroup` is the array of accumulators for one (rowLeaf,colLeaf) cell.
  type AccGroup = Acc[]; // length nValues
  const rawAcc = new Map<string, Map<string, AccGroup>>();
  const rowLeafOrder: string[] = [];
  const colLeafSet = new Map<string, string[]>(); // colPath → the ordered key parts (for header tree)
  const seenRowLeaf = new Set<string>();

  const pathOf = (r: Record<string, unknown>, fields: string[]): { parts: string[]; path: string } => {
    const parts = fields.map((f) => String(r[f] ?? ""));
    return { parts, path: parts.join(SEP) };
  };

  for (const r of rows) {
    const rl = pathOf(r, spec.rows);
    const cl = pathOf(r, spec.columns);
    if (!seenRowLeaf.has(rl.path)) {
      seenRowLeaf.add(rl.path);
      rowLeafOrder.push(rl.path);
    }
    if (!colLeafSet.has(cl.path)) colLeafSet.set(cl.path, cl.parts);
    let byCol = rawAcc.get(rl.path);
    if (!byCol) rawAcc.set(rl.path, (byCol = new Map()));
    let grp = byCol.get(cl.path);
    if (!grp) {
      grp = new Array(nValues);
      for (let vi = 0; vi < nValues; vi++) grp[vi] = newAcc();
      byCol.set(cl.path, grp);
    }
    for (let vi = 0; vi < nValues; vi++) pushAcc(grp[vi], num(r[values[vi].field]));
  }

  const colLeaves = [...colLeafSet.keys()];

  // Roll a rowLeaf's per-column groups into a single "row Total" (all columns
  // unioned) group — merging accumulators, so an average Total is the average of
  // ALL underlying values, NOT an average of per-column averages (Excel-exact).
  // Stored under the dedicated ROW_TOTAL colPath so it NEVER collides with a real
  // column path — including the "" path a blank column-field value produces (that
  // blank data must still be counted in the Total, which the old ""-keyed total
  // wrongly dropped). Always built; render decides whether to show the Total column.
  const rowTotalAcc = new Map<string, AccGroup>();
  for (const [rl, byCol] of rawAcc) {
    const tot: AccGroup = new Array(nValues);
    for (let vi = 0; vi < nValues; vi++) tot[vi] = newAcc();
    for (const grp of byCol.values()) for (let vi = 0; vi < nValues; vi++) mergeAcc(tot[vi], grp[vi]);
    rowTotalAcc.set(rl, tot);
  }

  // 3. Build the nested row tree from the ordered leaf paths.
  interface Build {
    key: string;
    path: string;
    level: number;
    children: Map<string, Build>;
    childOrder: string[];
    leaves: string[]; // rowLeafPaths directly at this node (only populated on true leaves)
    /** Rolled-up accumulators: colLeafPath → per-value Acc, plus "" = row Total. */
    acc: Map<string, AccGroup>;
  }
  const makeBuild = (key: string, path: string, level: number): Build => ({
    key,
    path,
    level,
    children: new Map(),
    childOrder: [],
    leaves: [],
    acc: new Map(),
  });
  const rootChildren = new Map<string, Build>();
  const rootOrder: string[] = [];
  for (const leaf of rowLeafOrder) {
    const parts = leaf.split(SEP);
    let map = rootChildren;
    let order = rootOrder;
    let prefix = "";
    let node: Build | undefined;
    for (let lvl = 0; lvl < parts.length; lvl++) {
      const key = parts[lvl];
      prefix = lvl === 0 ? key : `${prefix}${SEP}${key}`;
      node = map.get(key);
      if (!node) {
        node = makeBuild(key, prefix, lvl);
        map.set(key, node);
        order.push(key);
      }
      map = node.children;
      order = node.childOrder;
    }
    // `node` is now the true-leaf Build for this rowLeaf path.
    if (node) node.leaves.push(leaf);
  }

  // Merge one leaf's accumulators (per column + row Total) into a node's acc map.
  const ensureGroup = (m: Map<string, AccGroup>, colPath: string): AccGroup => {
    let g = m.get(colPath);
    if (!g) {
      g = new Array(nValues);
      for (let vi = 0; vi < nValues; vi++) g[vi] = newAcc();
      m.set(colPath, g);
    }
    return g;
  };
  const mergeGroupInto = (dst: Map<string, AccGroup>, colPath: string, src: AccGroup): void => {
    const g = ensureGroup(dst, colPath);
    for (let vi = 0; vi < nValues; vi++) mergeAcc(g[vi], src[vi]);
  };

  // Bottom-up finalize: a parent's accumulators are the O(children) merge of its
  // children's — descendants are NEVER re-scanned.
  const finalize = (b: Build): PivotNode => {
    const children = b.childOrder.map((k) => finalize(b.children.get(k)!));
    // Seed this node's accumulators from its own direct leaves (true leaves only).
    for (const leaf of b.leaves) {
      const byCol = rawAcc.get(leaf);
      if (byCol) for (const [col, grp] of byCol) mergeGroupInto(b.acc, col, grp);
      const tot = rowTotalAcc.get(leaf);      if (tot) mergeGroupInto(b.acc, ROW_TOTAL, tot);
    }
    // Merge each child's rolled-up accumulators upward.
    for (let i = 0; i < children.length; i++) {
      const cb = b.children.get(b.childOrder[i])!;
      for (const [col, grp] of cb.acc) mergeGroupInto(b.acc, col, grp);
    }
    const node: PivotNode = { key: b.key, path: b.path, level: b.level, children, values: new Map() };
    for (const [col, grp] of b.acc) for (let vi = 0; vi < nValues; vi++) node.values.set(cellKey(col, vi), readAcc(grp[vi], values[vi].aggregate));
    // Guarantee zero-filled cells for every (col,value) even if this node had no
    // data for that column (preserves the previous behaviour where aggFor→0).
    for (const col of colLeaves) for (let vi = 0; vi < nValues; vi++) { const k = cellKey(col, vi); if (!node.values.has(k)) node.values.set(k, aggregate([], values[vi].aggregate)); }
    for (let vi = 0; vi < nValues; vi++) { const k = cellKey(ROW_TOTAL, vi); if (!node.values.has(k)) node.values.set(k, aggregate([], values[vi].aggregate)); }
    return node;
  };
  const rootBuilds = rootOrder.map((k) => rootChildren.get(k)!);
  const rowTree = rootBuilds.map((b) => finalize(b));

  // 4. Grand totals (over ALL leaves) — merge every top-level node's accumulators.
  const grandAcc = new Map<string, AccGroup>();
  for (const b of rootBuilds) for (const [col, grp] of b.acc) mergeGroupInto(grandAcc, col, grp);
  const grand = new Map<string, number>();
  for (const [col, grp] of grandAcc) for (let vi = 0; vi < nValues; vi++) grand.set(cellKey(col, vi), readAcc(grp[vi], values[vi].aggregate));
  for (const col of colLeaves) for (let vi = 0; vi < nValues; vi++) { const k = cellKey(col, vi); if (!grand.has(k)) grand.set(k, aggregate([], values[vi].aggregate)); }
  for (let vi = 0; vi < nValues; vi++) { const k = cellKey(ROW_TOTAL, vi); if (!grand.has(k)) grand.set(k, aggregate([], values[vi].aggregate)); }

  // 6. Column header tree (levels of the column fields).
  const colTree = buildColTree(colLeaves);

  return { spec, rowTree, colLeaves, colTree, grand, values };
}

function buildColTree(colLeaves: string[]): PivotNode[] {
  const roots: PivotNode[] = [];
  const byKey = new Map<string, PivotNode>();
  for (const leaf of colLeaves) {
    if (leaf === "") continue;
    const parts = leaf.split(SEP);
    let siblings = roots;
    let prefix = "";
    for (let lvl = 0; lvl < parts.length; lvl++) {
      prefix = lvl === 0 ? parts[lvl] : `${prefix}${SEP}${parts[lvl]}`;
      let node = byKey.get(prefix);
      if (!node) {
        node = { key: parts[lvl], path: prefix, level: lvl, children: [], values: new Map() };
        byKey.set(prefix, node);
        siblings.push(node);
      }
      siblings = node.children;
    }
  }
  return roots;
}

/* ─── Render ────────────────────────────────────────────────────────────────── */

const HEADER_STYLE: CellStyle = { bl: 1, bg: { rgb: "#F9FAFB" }, cl: { rgb: "#475467" } };
const TOTAL_LABEL_STYLE: CellStyle = { bl: 1, bg: { rgb: "#F9FAFB" } };
const numStyle = (pattern: string, total = false): CellStyle =>
  total ? { n: { pattern }, ht: ALIGN_RIGHT, bl: 1, bg: { rgb: "#F9FAFB" } } : { n: { pattern }, ht: ALIGN_RIGHT };
const indentStyle = (level: number, extra?: CellStyle): CellStyle => (level > 0 ? { ...extra, pd: { l: level * 12 } } : { ...extra });

export interface RenderedPivot {
  cells: Record<number, Record<number, Cell>>;
  rowCount: number;
  columnCount: number;
}

/** Render a computed pivot model into a styled cell region. */
export function renderPivotModel(model: PivotModel): RenderedPivot {
  const { spec, colLeaves, values } = model;
  const collapsed = new Set(spec.collapsed ?? []);
  const showRowSubtotals = spec.showRowSubtotals ?? spec.rows.length > 1;
  const showGrand = spec.showGrandTotals ?? { row: true, column: true };
  const realCols = colLeaves.filter((c) => c !== "");

  const cells: Record<number, Record<number, Cell>> = {};
  const set = (r: number, c: number, cell: Cell) => {
    (cells[r] ??= {})[c] = cell;
  };

  // Column geometry: col 0 = row labels; then (realCols × values); then value Totals (if showGrand.column).
  const nValues = values.length;
  const dataStart = 1;
  const totalStart = dataStart + realCols.length * nValues;
  const columnCount = totalStart + (showGrand.column ? nValues : 0);

  // Header rows: a column-key header line per column level (if any), then the value labels.
  const colDepth = spec.columns.length;
  let headerRows = 0;
  const cellKey = (colPath: string, vi: number) => `${colPath}${SEP}${vi}`;

  // Column-group header (single flattened line for simplicity of the compact view).
  if (colDepth > 0) {
    const hr = headerRows;
    set(hr, 0, { v: "", s: HEADER_STYLE });
    realCols.forEach((col, ci) => {
      const label = col.split(SEP).join(" / ");
      // Blank-fill the group's span, then write the column label on the FIRST sub-column of
      // the group so a multi-value pivot (e.g. Sum | Count under each column) unambiguously
      // shows which column each value pair belongs to (Excel puts the group label above the span).
      values.forEach((_, vi) => set(hr, dataStart + ci * nValues + vi, { v: "", s: HEADER_STYLE }));
      set(hr, dataStart + ci * nValues, { v: label, s: HEADER_STYLE });
    });
    headerRows++;
  }
  // Value-label header line.
  {
    const hr = headerRows;
    set(hr, 0, { v: spec.rows.join(" / ") || "", s: HEADER_STYLE });
    realCols.forEach((col, ci) => {
      values.forEach((v, vi) => {
        const label = nValues > 1 || colDepth === 0 ? valueLabel(v) : col.split(SEP).join(" / ");
        set(hr, dataStart + ci * nValues + vi, { v: label, s: HEADER_STYLE });
      });
    });
    if (showGrand.column) values.forEach((v, vi) => set(hr, totalStart + vi, { v: nValues > 1 ? `Total ${valueLabel(v)}` : "Grand Total", s: HEADER_STYLE }));
    headerRows++;
  }

  // Body rows: walk the row tree depth-first, emitting a row per node (+ subtotal when it has children).
  let r = headerRows;
  const emitValueCells = (row: number, node: PivotNode | null, total: boolean) => {
    realCols.forEach((col, ci) => {
      values.forEach((v, vi) => {
        const val = node ? node.values.get(cellKey(col, vi)) : model.grand.get(cellKey(col, vi));
        set(row, dataStart + ci * nValues + vi, { v: val ?? 0, s: numStyle(v.numFmt ?? NUMBER_PATTERN, total) });
      });
    });
    if (showGrand.column) {
      values.forEach((v, vi) => {
        const val = node ? node.values.get(cellKey(ROW_TOTAL, vi)) : model.grand.get(cellKey(ROW_TOTAL, vi));
        set(row, totalStart + vi, { v: val ?? 0, s: numStyle(v.numFmt ?? NUMBER_PATTERN, true) });
      });
    }
  };

  const walk = (nodes: PivotNode[]) => {
    for (const node of nodes) {
      const hasChildren = node.children.length > 0;
      const isCollapsed = collapsed.has(node.path);
      const chevron = hasChildren ? (isCollapsed ? "▸ " : "▾ ") : "";
      // Group header / leaf row.
      set(r, 0, { v: `${chevron}${node.key}`, s: indentStyle(node.level, hasChildren ? { bl: 1 } : undefined) });
      emitValueCells(r, node, false);
      r++;
      if (hasChildren && !isCollapsed) {
        walk(node.children);
        if (showRowSubtotals) {
          set(r, 0, { v: `${node.key} Total`, s: indentStyle(node.level, TOTAL_LABEL_STYLE) });
          emitValueCells(r, node, true);
          r++;
        }
      }
    }
  };
  walk(model.rowTree);

  // Grand-total row.
  if (showGrand.row) {
    set(r, 0, { v: "Grand Total", s: TOTAL_LABEL_STYLE });
    emitValueCells(r, null, true);
    r++;
  }

  return { cells, rowCount: r, columnCount };
}
