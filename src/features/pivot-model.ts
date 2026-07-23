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
  const f = nums.filter(Number.isFinite);
  switch (agg) {
    case "count":
      return nums.length;
    case "countNumbers":
      return f.length;
    case "countunique":
      return new Set(f).size;
    case "average":
      return f.length ? f.reduce((s, x) => s + x, 0) / f.length : 0;
    case "min":
      // Exclude non-numbers (text / blanks) like Excel's MIN — else one stray "N/A"
      // poisons the whole group to NaN.
      return f.length ? Math.min(...f) : 0;
    case "max":
      return f.length ? Math.max(...f) : 0;
    case "median": {
      if (!f.length) return 0;
      const s = [...f].sort((x, y) => x - y);
      const m = s.length >> 1;
      return s.length % 2 ? s[m] : (s[m - 1] + s[m]) / 2;
    }
    case "product":
      return f.length ? f.reduce((p, x) => p * x, 1) : 0;
    case "var":
    case "stdev": {
      if (f.length < 2) return 0;
      const mean = f.reduce((s, x) => s + x, 0) / f.length;
      const v = f.reduce((s, x) => s + (x - mean) ** 2, 0) / (f.length - 1);
      return agg === "var" ? v : Math.sqrt(v);
    }
    case "varp":
    case "stdevp": {
      if (!f.length) return 0;
      const mean = f.reduce((s, x) => s + x, 0) / f.length;
      const v = f.reduce((s, x) => s + (x - mean) ** 2, 0) / f.length;
      return agg === "varp" ? v : Math.sqrt(v);
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
  n: number; // total observations (for "count"/COUNTA).
  fn: number; // finite-value count (for "countNumbers"/COUNT / "average").
  min: number; // running min over finite values (Infinity if none seen).
  max: number; // running max over finite values (-Infinity if none seen).
  sq: number; // Σx² over finite values (for STDEV/STDEVP/VAR/VARP — mergeable).
  prod: number; // Πx over finite values (for PRODUCT — mergeable; identity 1).
  vals?: number[]; // finite values, tracked ONLY when a value field uses MEDIAN.
  uniq?: Set<string>; // distinct non-empty raw values, ONLY when a field uses COUNTUNIQUE.
}
const newAcc = (needVals = false, needUniq = false): Acc => ({
  sum: 0,
  n: 0,
  fn: 0,
  min: Infinity,
  max: -Infinity,
  sq: 0,
  prod: 1,
  vals: needVals ? [] : undefined,
  uniq: needUniq ? new Set<string>() : undefined,
});

/** Fold one RAW value into an accumulator (keeps distinctness for COUNTUNIQUE + the
 *  value list for MEDIAN; everything else is O(1) sufficient statistics). */
/**
 * Coerce a cell value to a number for numeric aggregation. Imported .xlsx cells often
 * carry the DISPLAY string ("$196,282.09", "(1,234.50)", "45%") rather than a raw number,
 * and plain Number() returns NaN for those — which silently dropped them from SUM/AVG/etc.
 * (the "pivot shows 0 / no data" bug). Strip currency symbols + thousands separators, read
 * accounting-style "(n)" as negative, and honor a trailing "%".
 */
function toNumber(raw: unknown): number {
  if (typeof raw === "number") return raw;
  if (typeof raw !== "string") return Number(raw);
  let s = raw.trim();
  if (s === "") return NaN;
  const paren = /^\((.*)\)$/.exec(s);
  if (paren) s = "-" + paren[1];
  const pct = s.endsWith("%");
  if (pct) s = s.slice(0, -1);
  s = s.replace(/[,$£€¥\s ]/g, "");
  const n = Number(s);
  return Number.isFinite(n) ? (pct ? n / 100 : n) : NaN;
}

function pushAcc(a: Acc, raw: unknown): void {
  a.n += 1;
  if (a.uniq && raw != null && String(raw).trim() !== "") a.uniq.add(String(raw));
  const x = toNumber(raw);
  if (Number.isFinite(x)) {
    a.sum += x;
    a.fn += 1;
    a.sq += x * x;
    a.prod *= x;
    if (x < a.min) a.min = x;
    if (x > a.max) a.max = x;
    a.vals?.push(x);
  }
}

/** Merge `src` INTO `dst` in O(1) (O(k) when tracking median/uniq) — associative +
 *  commutative, so roll-up order doesn't matter and a parent = merge of its children. */
function mergeAcc(dst: Acc, src: Acc): void {
  dst.sum += src.sum;
  dst.n += src.n;
  dst.fn += src.fn;
  dst.sq += src.sq;
  dst.prod *= src.prod;
  if (src.min < dst.min) dst.min = src.min;
  if (src.max > dst.max) dst.max = src.max;
  if (dst.vals && src.vals) for (const v of src.vals) dst.vals.push(v);
  if (dst.uniq && src.uniq) for (const u of src.uniq) dst.uniq.add(u);
}

/** Read the final aggregate out of an accumulator (matches `aggregate()` exactly). */
function readAcc(a: Acc, agg: PivotAggregate): number {
  switch (agg) {
    case "count":
      return a.n;
    case "countNumbers":
      return a.fn;
    case "countunique":
      return a.uniq ? a.uniq.size : 0;
    case "average":
      return a.fn ? a.sum / a.fn : 0;
    case "min":
      return a.fn ? a.min : 0;
    case "max":
      return a.fn ? a.max : 0;
    case "median": {
      if (!a.vals || !a.vals.length) return 0;
      const s = [...a.vals].sort((x, y) => x - y);
      const m = s.length >> 1;
      return s.length % 2 ? s[m] : (s[m - 1] + s[m]) / 2;
    }
    case "product":
      return a.fn ? a.prod : 0;
    case "var":
    case "stdev": {
      if (a.fn < 2) return 0;
      const v = (a.sq - (a.sum * a.sum) / a.fn) / (a.fn - 1); // sample variance
      return agg === "var" ? v : Math.sqrt(Math.max(0, v));
    }
    case "varp":
    case "stdevp": {
      if (a.fn < 1) return 0;
      const v = (a.sq - (a.sum * a.sum) / a.fn) / a.fn; // population variance
      return agg === "varp" ? v : Math.sqrt(Math.max(0, v));
    }
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
    count: "COUNTA",
    countNumbers: "COUNT",
    countunique: "COUNTUNIQUE",
    average: "Average",
    min: "Min",
    max: "Max",
    median: "Median",
    product: "Product",
    stdev: "STDEV",
    stdevp: "STDEVP",
    var: "VAR",
    varp: "VARP",
  };
  return `${verb[v.aggregate]} of ${v.field}`;
}

/** Compute the full pivot tree from a source + spec. */
export function computePivotModel(source: PivotSource, spec: PivotSpec): PivotModel {
  // Use ONLY the value fields the user configured. Previously an empty Values list
  // silently invented `count(fields[0])`, which manufactured a constant "Grand Total =
  // row-count" (e.g. 1000) that never reflected the layout — the source of the "pivot
  // shows data I didn't ask for / won't clear" bug. With no values the pivot shows just
  // the row/column labels (Google-Sheets behavior).
  const values = spec.values;

  // 1. Filter rows.
  const filters = spec.filters ?? [];
  const rows = source.rows.filter((r) => filters.every((f) => !f.include || f.include.includes(String(r[f.field] ?? ""))));

  const nValues = values.length;
  // MEDIAN needs the value multiset + COUNTUNIQUE the distinct set — track them per value
  // field ONLY when used, so the fast O(1) roll-up is unaffected for the common aggregates.
  const needVals = values.map((v) => v.aggregate === "median");
  const needUniq = values.map((v) => v.aggregate === "countunique");
  const mkGroup = (): Acc[] => {
    const g = new Array<Acc>(nValues);
    for (let vi = 0; vi < nValues; vi++) g[vi] = newAcc(needVals[vi], needUniq[vi]);
    return g;
  };
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
      grp = mkGroup();
      byCol.set(cl.path, grp);
    }
    for (let vi = 0; vi < nValues; vi++) pushAcc(grp[vi], r[values[vi].field]);
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
    const tot: AccGroup = mkGroup();
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
      g = mkGroup();
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
  // "Order" (asc/desc) per dimension field — sort a node's children by their label.
  const cmp = new Intl.Collator(undefined, { numeric: true, sensitivity: "base" });
  const sortOrderFor = (field: string | undefined): "asc" | "desc" | undefined => (field ? spec.dimSettings?.[field]?.order : undefined);
  const sortKeys = (keys: string[], field: string | undefined) => {
    // Google Sheets sorts pivot row/column groups ascending by default; "desc" flips it.
    // (Matches the panel's Order select, which shows "Ascending" when unset.)
    const ord = sortOrderFor(field);
    keys.sort((a, c) => (ord === "desc" ? -cmp.compare(a, c) : cmp.compare(a, c)));
  };
  // "Sort by": when a dimension's `sortBy` names a VALUE field (not its own label), the
  // sibling groups at that level are ordered by that value's aggregated total instead of by
  // label. The child PivotNodes must already be finalized (their ROW_TOTAL values populated).
  const valueSortIndex = (field: string | undefined): number => {
    if (!field) return -1;
    const sb = spec.dimSettings?.[field]?.sortBy;
    if (!sb || sb === field) return -1; // default: sort by label (handled by sortKeys)
    return values.findIndex((v) => v.field === sb);
  };
  const applyValueSort = (children: PivotNode[], field: string | undefined): PivotNode[] => {
    const vi = valueSortIndex(field);
    if (vi < 0) return children;
    const ord = sortOrderFor(field) === "desc" ? -1 : 1;
    const totOf = (n: PivotNode) => (n.values.get(cellKey(ROW_TOTAL, vi)) ?? 0) as number;
    // Stable numeric sort by the chosen value's grand total for each group.
    return children
      .map((n, i) => ({ n, i }))
      .sort((a, b) => ord * (totOf(a.n) - totOf(b.n)) || a.i - b.i)
      .map((x) => x.n);
  };
  const finalize = (b: Build): PivotNode => {
    sortKeys(b.childOrder, spec.rows[b.level + 1]); // children are the NEXT row field
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
    // "Sort by": reorder these children by a value's total if this level's dim asks for it
    // (children are finalized here, so their ROW_TOTAL values exist). Falls back to the
    // label order established by sortKeys() above when sortBy is the field's own label.
    const orderedChildren = applyValueSort(children, spec.rows[b.level + 1]);
    const node: PivotNode = { key: b.key, path: b.path, level: b.level, children: orderedChildren, values: new Map() };
    for (const [col, grp] of b.acc) for (let vi = 0; vi < nValues; vi++) node.values.set(cellKey(col, vi), readAcc(grp[vi], values[vi].aggregate));
    // Guarantee zero-filled cells for every (col,value) even if this node had no
    // data for that column (preserves the previous behaviour where aggFor→0).
    for (const col of colLeaves) for (let vi = 0; vi < nValues; vi++) { const k = cellKey(col, vi); if (!node.values.has(k)) node.values.set(k, aggregate([], values[vi].aggregate)); }
    for (let vi = 0; vi < nValues; vi++) { const k = cellKey(ROW_TOTAL, vi); if (!node.values.has(k)) node.values.set(k, aggregate([], values[vi].aggregate)); }
    return node;
  };
  sortKeys(rootOrder, spec.rows[0]); // top-level row groups (label order)
  const rootBuilds = rootOrder.map((k) => rootChildren.get(k)!);
  const rowTree = applyValueSort(rootBuilds.map((b) => finalize(b)), spec.rows[0]); // then Sort-by-value

  // 4. Grand totals (over ALL leaves) — merge every top-level node's accumulators.
  const grandAcc = new Map<string, AccGroup>();
  for (const b of rootBuilds) for (const [col, grp] of b.acc) mergeGroupInto(grandAcc, col, grp);
  const grand = new Map<string, number>();
  for (const [col, grp] of grandAcc) for (let vi = 0; vi < nValues; vi++) grand.set(cellKey(col, vi), readAcc(grp[vi], values[vi].aggregate));
  for (const col of colLeaves) for (let vi = 0; vi < nValues; vi++) { const k = cellKey(col, vi); if (!grand.has(k)) grand.set(k, aggregate([], values[vi].aggregate)); }
  for (let vi = 0; vi < nValues; vi++) { const k = cellKey(ROW_TOTAL, vi); if (!grand.has(k)) grand.set(k, aggregate([], values[vi].aggregate)); }

  // 6. Column "Sort by" (value-based): order the column leaves within each parent group by
  // the chosen value's grand total. Parent groups keep their established (label) order so
  // nested column headers stay contiguous; only siblings under a shared prefix are reordered.
  const colDimField = spec.columns[spec.columns.length - 1];
  const colSortVi = valueSortIndex(colDimField);
  let orderedColLeaves = colLeaves;
  if (colSortVi >= 0 && colLeaves.length > 1) {
    const ord = sortOrderFor(colDimField) === "desc" ? -1 : 1;
    const parentOf = (leaf: string) => { const i = leaf.lastIndexOf(SEP); return i < 0 ? "" : leaf.slice(0, i); };
    const totOf = (leaf: string) => (grand.get(cellKey(leaf, colSortVi)) ?? 0) as number;
    const groups: string[][] = [];
    const groupIdx = new Map<string, number>();
    for (const leaf of colLeaves) {
      const p = parentOf(leaf);
      let gi = groupIdx.get(p);
      if (gi === undefined) { gi = groups.length; groupIdx.set(p, gi); groups.push([]); }
      groups[gi].push(leaf);
    }
    for (const g of groups) g.sort((a, b) => ord * (totOf(a) - totOf(b)));
    orderedColLeaves = groups.flat();
  }

  // 7. Column header tree (levels of the column fields).
  const colTree = buildColTree(orderedColLeaves);

  return { spec, rowTree, colLeaves: orderedColLeaves, colTree, grand, values };
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
  // A fully-empty pivot (no rows, columns, or values) renders NOTHING — the in-place
  // apply then clears any stale rectangle, so an unconfigured/cleared pivot is blank
  // (matches Google Sheets) instead of leaving a phantom "Grand Total" behind.
  if (spec.rows.length === 0 && spec.columns.length === 0 && values.length === 0) {
    return { cells: {}, rowCount: 0, columnCount: 0 };
  }
  const collapsed = new Set(spec.collapsed ?? []);
  const showRowSubtotals = spec.showRowSubtotals ?? spec.rows.length > 1;
  // Grand totals need at least one value to total. With rows/columns but no values the
  // pivot lists the distinct labels only (no numeric grand total), like Google Sheets.
  const showGrand = values.length === 0 ? { row: false, column: false } : (spec.showGrandTotals ?? { row: true, column: true });
  const realCols = colLeaves.filter((c) => c !== "");

  const cells: Record<number, Record<number, Cell>> = {};
  const set = (r: number, c: number, cell: Cell) => {
    (cells[r] ??= {})[c] = cell;
  };

  // Column geometry: col 0 = row labels; then (realCols × values); then value Totals (if showGrand.column).
  // `perCol` = column-slots per distinct COLUMN value. With no values we still give each column
  // ONE slot so a Columns field lays out its labels (Google Sheets shows the distinct column
  // values as headers even before a Value is added) instead of collapsing to nothing.
  const nValues = values.length;
  const perCol = nValues || 1;
  const dataStart = 1;
  const totalStart = dataStart + realCols.length * perCol;
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
      // Blank-fill the group's span (perCol slots — ≥1 even with no values), then write the
      // column label on the FIRST sub-column so a multi-value pivot (e.g. Sum | Count under each
      // column) unambiguously shows which column each value belongs to.
      for (let vi = 0; vi < perCol; vi++) set(hr, dataStart + ci * perCol + vi, { v: "", s: HEADER_STYLE });
      set(hr, dataStart + ci * perCol, { v: label, s: HEADER_STYLE });
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
        set(hr, dataStart + ci * perCol + vi, { v: label, s: HEADER_STYLE });
      });
    });
    if (showGrand.column) values.forEach((v, vi) => set(hr, totalStart + vi, { v: nValues > 1 ? `Total ${valueLabel(v)}` : "Grand Total", s: HEADER_STYLE }));
    headerRows++;
  }

  // Body rows: walk the row tree depth-first, emitting a row per node (+ subtotal when it has children).
  let r = headerRows;
  // "Show as": re-express a raw cell as a % of its row total / column total / grand total.
  const PCT_PATTERN = "0.0%";
  const showAsCell = (raw: number | undefined, colPath: string, vi: number, node: PivotNode | null): { v: number; pct: boolean } => {
    const mode = values[vi].showAs ?? "default";
    if (mode === "default" || raw == null) return { v: raw ?? 0, pct: false };
    const rowTot = (node ? node.values.get(cellKey(ROW_TOTAL, vi)) : model.grand.get(cellKey(ROW_TOTAL, vi))) ?? 0;
    const colTot = model.grand.get(cellKey(colPath, vi)) ?? 0;
    const grandTot = model.grand.get(cellKey(ROW_TOTAL, vi)) ?? 0;
    const den = mode === "pctOfRow" ? rowTot : mode === "pctOfCol" ? colTot : grandTot;
    return { v: den ? raw / den : 0, pct: true };
  };
  const emitValueCells = (row: number, node: PivotNode | null, total: boolean) => {
    realCols.forEach((col, ci) => {
      values.forEach((v, vi) => {
        const raw = node ? node.values.get(cellKey(col, vi)) : model.grand.get(cellKey(col, vi));
        const { v: out, pct } = showAsCell(raw, col, vi, node);
        set(row, dataStart + ci * perCol + vi, { v: out, s: numStyle(pct ? PCT_PATTERN : (v.numFmt ?? NUMBER_PATTERN), total) });
      });
    });
    if (showGrand.column) {
      values.forEach((v, vi) => {
        const raw = node ? node.values.get(cellKey(ROW_TOTAL, vi)) : model.grand.get(cellKey(ROW_TOTAL, vi));
        const { v: out, pct } = showAsCell(raw, ROW_TOTAL, vi, node);
        set(row, totalStart + vi, { v: out, s: numStyle(pct ? PCT_PATTERN : (v.numFmt ?? NUMBER_PATTERN), true) });
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
        // Per-level "Show totals" (dimSettings), falling back to the global default.
        const showThisTotal = spec.dimSettings?.[spec.rows[node.level]]?.showTotals ?? showRowSubtotals;
        if (showThisTotal) {
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
