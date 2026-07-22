/**
 * Rich XLSX → Univer snapshot converter (free-tier, client-only).
 *
 * Univer's own `importXLSXToSnapshotAsync` requires the Univer (Pro) server.
 * This module does the same job in the browser for free: it reads a workbook
 * with ExcelJS and emits a full `IWorkbookData` snapshot that Univer can render
 * directly — preserving cell VALUES, STYLES (fill / font / colour / bold /
 * italic / underline / strike / alignment / wrap / rotation / borders), NUMBER
 * FORMATS, DATES, MERGED CELLS, COLUMN WIDTHS, ROW HEIGHTS, FROZEN PANES and
 * ALL WORKSHEETS.
 *
 * Contrast with `parseFileToGrid` (import-data.ts), which flattens everything to
 * plain values and (because ExcelJS returns the master value for every slave
 * cell of a merge) duplicates merged content. This path fixes all of that.
 *
 * The heavy `exceljs` dependency is dynamically imported so it is only pulled in
 * when an actual import happens.
 */
import type { WorkbookData } from "./types";
import type { ImportedPivot } from "./pivot-import";

/**
 * The snapshot returned by {@link parseXlsxToSnapshot}. It is a Univer
 * `IWorkbookData` (Univer ignores the extra non-schema keys) PLUS two escape-hatch
 * fields carrying data Univer's snapshot can't hold natively:
 *   - `drawingsImport` — floating images for the Facade insert path.
 *   - `pivotsImport` — reconstructed INTERACTIVE pivots (source + spec + location)
 *     so a host can open an imported pivot in `LevichSheet`'s `pivotInteractive`
 *     mode (seeding `initialSpec`) instead of the static cell render.
 */
export type ParsedSnapshot = WorkbookData & {
  drawingsImport?: ImportImage[];
  pivotsImport?: ImportedPivot[];
};

/* -------------------------------------------------------------------------- */
/* Loose ExcelJS shapes (avoid tight coupling to exceljs' generics)           */
/* -------------------------------------------------------------------------- */
type Argb = { argb?: string; theme?: number; tint?: number };
interface XFont {
  name?: string;
  size?: number;
  bold?: boolean;
  italic?: boolean;
  underline?: boolean | string;
  strike?: boolean;
  color?: Argb;
}
interface XBorderEdge {
  style?: string;
  color?: Argb;
}
interface XCell {
  row: number;
  col: number;
  value: unknown;
  /** ExcelJS resolved-formula getter — populated for shared-formula slaves too. */
  formula?: string;
  /** ExcelJS cell model — carries `result` (cached value) for formula cells whose
   *  `value.result` ExcelJS leaves undefined (notably cross-sheet-reference formulas). */
  model?: { result?: unknown };
  /** Cell hyperlink target (e.g. "mailto:x@y.com" / "https://…"). */
  hyperlink?: string | { hyperlink?: string };
  type?: number;
  numFmt?: string;
  font?: XFont;
  fill?: { type?: string; pattern?: string; fgColor?: Argb; bgColor?: Argb };
  alignment?: { horizontal?: string; vertical?: string; wrapText?: boolean; textRotation?: number | string; indent?: number };
  border?: { top?: XBorderEdge; bottom?: XBorderEdge; left?: XBorderEdge; right?: XBorderEdge };
}
interface XRow {
  number: number;
  height?: number;
  hidden?: boolean;
  outlineLevel?: number;
  eachCell(opts: { includeEmpty?: boolean }, cb: (cell: XCell, colNumber: number) => void): void;
}
interface XColumn {
  number?: number;
  width?: number;
  hidden?: boolean;
  outlineLevel?: number;
}
interface XAnchor { nativeCol?: number; nativeColOff?: number; nativeRow?: number; nativeRowOff?: number; col?: number; row?: number }
interface XImageAnchor { imageId: string | number; range?: { tl?: XAnchor; br?: XAnchor; ext?: { width?: number; height?: number } } }
interface XWorksheet {
  name: string;
  rowCount: number;
  columnCount: number;
  actualColumnCount?: number;
  columns?: XColumn[];
  state?: string; // "visible" | "hidden" | "veryHidden"
  views?: Array<{ state?: string; xSplit?: number; ySplit?: number; showGridLines?: boolean }>;
  properties?: { tabColor?: Argb; defaultRowHeight?: number; defaultColWidth?: number };
  autoFilter?: XAutoFilter;
  conditionalFormattings?: XCfRuleSet[];
  model?: { merges?: string[] };
  eachRow(opts: { includeEmpty?: boolean }, cb: (row: XRow, rowNumber: number) => void): void;
  getImages?(): XImageAnchor[];
}
type XAddr = string | { row: number; column: number };
type XAutoFilter = string | { from?: XAddr; to?: XAddr } | null;
interface URange { startRow: number; startColumn: number; endRow: number; endColumn: number }
interface XCfValue { type?: string; value?: number | string }
interface XCfRule {
  type?: string;
  operator?: string;
  formulae?: Array<string | number>;
  text?: string;
  style?: { font?: { bold?: boolean; italic?: boolean; color?: Argb }; fill?: { type?: string; pattern?: string; fgColor?: Argb; bgColor?: Argb } };
  cfvo?: XCfValue[];
  color?: Argb | Argb[];
}
interface XCfRuleSet { ref?: string; rules?: XCfRule[] }
interface XMedia { type?: string; extension?: string; base64?: string; buffer?: Uint8Array | ArrayBuffer | { data?: number[] } }
interface XWorkbook { worksheets: XWorksheet[]; media?: XMedia[]; model?: { media?: XMedia[] } }

/** A floating image extracted from the xlsx, to be placed via the Facade after
 *  the sheet loads (Univer's snapshot drawing resource is not hand-authored). */
export interface ImportImage {
  sheetIndex: number;
  /** `data:image/…;base64,…` source. */
  base64: string;
  /** 0-based anchor cell + pixel offset within it. */
  col: number;
  row: number;
  colOffset: number;
  rowOffset: number;
  /** Rendered size in pixels. */
  width: number;
  height: number;
}

/* -------------------------------------------------------------------------- */
/* Univer style shape (short keys, matching build-workbook.ts / IStyleData)   */
/* -------------------------------------------------------------------------- */
type BorderSeg = { s: number; cl: { rgb: string } };
interface UStyle {
  bl?: number;
  it?: number;
  ul?: { s: number };
  st?: { s: number };
  fs?: number;
  ff?: string;
  cl?: { rgb: string };
  bg?: { rgb: string };
  ht?: number;
  vt?: number;
  tb?: number;
  tr?: { a: number; v: number };
  bd?: { t?: BorderSeg; b?: BorderSeg; l?: BorderSeg; r?: BorderSeg };
  n?: { pattern: string };
  pd?: { t?: number; r?: number; b?: number; l?: number };
}
type UCell = { v?: string | number | boolean; f?: string; s?: string } | null;

/* ---- Colour ---------------------------------------------------------------
   The standard Office ("Office 2013+") theme palette, indexed the way Excel's
   cell `theme` attribute references it. NOTE the dk1/lt1 and dk2/lt2 SWAP:
   Excel's colour map swaps indices 0↔1 and 2↔3 relative to the raw scheme
   order, so `theme 0` = lt1 (white), `theme 1` = dk1 (BLACK), `theme 2` = lt2,
   `theme 3` = dk2. Getting this wrong turns black theme-text white (and vice
   versa). Explicit ARGB colours (the common case) don't use this table. */
const THEME_COLORS = [
  "#FFFFFF", "#000000", "#E7E6E6", "#44546A", "#4472C4", "#ED7D31",
  "#A5A5A5", "#FFC000", "#5B9BD5", "#70AD47", "#0563C1", "#954F72",
];

function hexToRgb(hex: string): [number, number, number] {
  const n = parseInt(hex.replace(/^#/, ""), 16);
  return [(n >> 16) & 255, (n >> 8) & 255, n & 255];
}
const to2 = (x: number) => Math.max(0, Math.min(255, Math.round(x))).toString(16).padStart(2, "0");
const rgbToHex = (r: number, g: number, b: number) => `#${to2(r)}${to2(g)}${to2(b)}`.toUpperCase();

/** Excel tint: negative darkens, positive lightens (approximated per-channel). */
function applyTint(hex: string, tint: number): string {
  if (!tint) return hex;
  const [r, g, b] = hexToRgb(hex);
  const f = (c: number) => (tint < 0 ? c * (1 + tint) : c * (1 - tint) + 255 * tint);
  return rgbToHex(f(r), f(g), f(b));
}

function argbToHex(argb: string): string | null {
  let s = argb.replace(/^#/, "");
  if (s.length === 8) s = s.slice(2); // drop leading alpha
  if (s.length === 6) return `#${s.toUpperCase()}`;
  return null;
}

/** The legacy Excel 56-colour indexed palette (indices 8–63 are the usable
 *  colours; 64 = automatic/window-text, 65 = window-background). */
const INDEXED_PALETTE: Record<number, string> = {
  0: "#000000", 1: "#FFFFFF", 2: "#FF0000", 3: "#00FF00", 4: "#0000FF", 5: "#FFFF00", 6: "#FF00FF", 7: "#00FFFF",
  8: "#000000", 9: "#FFFFFF", 10: "#FF0000", 11: "#00FF00", 12: "#0000FF", 13: "#FFFF00", 14: "#FF00FF", 15: "#00FFFF",
  16: "#800000", 17: "#008000", 18: "#000080", 19: "#808000", 20: "#800080", 21: "#008080", 22: "#C0C0C0", 23: "#808080",
  24: "#9999FF", 25: "#993366", 26: "#FFFFCC", 27: "#CCFFFF", 28: "#660066", 29: "#FF8080", 30: "#0066CC", 31: "#CCCCFF",
  32: "#000080", 33: "#FF00FF", 34: "#FFFF00", 35: "#00FFFF", 36: "#800080", 37: "#800000", 38: "#008080", 39: "#0000FF",
  40: "#00CCFF", 41: "#CCFFFF", 42: "#CCFFCC", 43: "#FFFF99", 44: "#99CCFF", 45: "#FF99CC", 46: "#CC99FF", 47: "#FFCC99",
  48: "#3366FF", 49: "#33CCCC", 50: "#99CC00", 51: "#FFCC00", 52: "#FF9900", 53: "#FF6600", 54: "#666699", 55: "#969696",
  56: "#003366", 57: "#339966", 58: "#003300", 59: "#333300", 60: "#993300", 61: "#993366", 62: "#333399", 63: "#333333",
};

/** Resolve an ExcelJS colour object to a `#RRGGBB` hex, or null if unknown. */
function colorToHex(color?: Argb): string | null {
  if (!color) return null;
  if (typeof color.argb === "string") return argbToHex(color.argb);
  if (typeof color.theme === "number") {
    const base = THEME_COLORS[color.theme];
    return base ? applyTint(base, color.tint ?? 0) : null;
  }
  const indexed = (color as { indexed?: number }).indexed;
  if (typeof indexed === "number") {
    if (indexed === 64 || indexed === 65) return null; // automatic → let it default
    return INDEXED_PALETTE[indexed] ?? null;
  }
  return null;
}

/* ---- Borders -------------------------------------------------------------- */
const BORDER_STYLE_MAP: Record<string, number> = {
  thin: 1, hair: 2, dotted: 3, dashed: 4, dashDot: 5, dashDotDot: 6,
  double: 7, medium: 8, mediumDashed: 9, mediumDashDot: 10,
  mediumDashDotDot: 11, slantDashDot: 12, thick: 13,
};
function edge(b?: XBorderEdge): BorderSeg | null {
  if (!b || !b.style) return null;
  return { s: BORDER_STYLE_MAP[b.style] ?? 1, cl: { rgb: colorToHex(b.color) ?? "#000000" } };
}

/* ---- Alignment ------------------------------------------------------------ */
const H_ALIGN: Record<string, number> = { left: 1, center: 2, right: 3, justify: 4, fill: 1, centerContinuous: 2, distributed: 6 };
const V_ALIGN: Record<string, number> = { top: 1, middle: 2, bottom: 3, distributed: 2, justify: 2 };

/* ---- Dates ---------------------------------------------------------------- */
/** JS Date → Excel serial number (days since 1899-12-30, UTC-based). */
function dateToSerial(d: Date): number {
  const serial = d.getTime() / 86_400_000 + 25_569;
  return Math.round(serial * 1e6) / 1e6; // trim float noise, keep intraday time
}

/* -------------------------------------------------------------------------- */
/* Cell → { value, style }                                                    */
/* -------------------------------------------------------------------------- */
function cellValue(cell: XCell): { v?: string | number | boolean; f?: string; isDate?: boolean } {
  const v = cell.value;
  if (v == null) return {};
  if (typeof v === "number" || typeof v === "string" || typeof v === "boolean") return { v };
  if (v instanceof Date) return { v: dateToSerial(v), isDate: true };
  if (typeof v === "object") {
    const o = v as Record<string, unknown>;
    // Formula cell. Prefer the RESOLVED `cell.formula` getter — for a shared-
    // formula slave the raw value only holds `{ sharedFormula: <master> }`, but
    // the getter returns this cell's own translated formula (e.g. Text(B16,…)).
    // Without it the slave came out as an empty but date-formatted cell → "NaN".
    if ((typeof cell.formula === "string" && cell.formula) || "formula" in o || "sharedFormula" in o) {
      const f = (typeof cell.formula === "string" && cell.formula) ? cell.formula : ((o.formula as string | undefined) || undefined);
      // ExcelJS often omits `value.result` for cross-sheet-reference formulas but still
      // carries the cached value on `cell.model.result`. Fall back to it so the cell keeps
      // Excel's number instead of arriving empty — an empty formula cell is recomputed by
      // Univer (WHEN_EMPTY), turning unsupported functions / broken refs into #NAME?/#VALUE!
      // and slowing the initial open with tens of thousands of needless recalcs.
      const res = o.result ?? cell.model?.result;
      const out: { v?: string | number | boolean; f?: string; isDate?: boolean } = {};
      if (f) out.f = `=${f}`;
      // Always carry the cached result as the value so the cell renders
      // immediately (Univer recalcs live formulas on top of it).
      if (res instanceof Date) { out.v = dateToSerial(res); out.isDate = true; }
      else if (typeof res === "number" || typeof res === "string" || typeof res === "boolean") out.v = res;
      else if (res && typeof res === "object" && "error" in (res as object)) out.v = String((res as { error?: unknown }).error ?? "");
      return out;
    }
    // Rich text: concat runs (per-run inline styling is flattened).
    if ("richText" in o && Array.isArray(o.richText)) return { v: (o.richText as Array<{ text?: string }>).map((t) => t.text ?? "").join("") };
    // Hyperlink cell: keep the display text.
    if ("text" in o) return { v: String(o.text ?? "") };
    if ("error" in o) return { v: String(o.error ?? "") };
  }
  return { v: String(v) };
}

function buildStyle(cell: XCell, isDate: boolean): UStyle {
  const s: UStyle = {};
  const font = cell.font;
  if (font) {
    if (font.bold) s.bl = 1;
    if (font.italic) s.it = 1;
    if (font.underline) s.ul = { s: 1 };
    if (font.strike) s.st = { s: 1 };
    if (font.size) s.fs = font.size;
    if (font.name) s.ff = font.name;
    const fc = colorToHex(font.color);
    if (fc) s.cl = { rgb: fc };
  }
  const fill = cell.fill;
  if (fill && fill.type === "pattern" && fill.pattern && fill.pattern !== "none") {
    const bg = colorToHex(fill.fgColor);
    if (bg) s.bg = { rgb: bg };
  }
  const al = cell.alignment;
  if (al) {
    const h = al.horizontal ? H_ALIGN[al.horizontal] : undefined;
    if (h) s.ht = h;
    const vv = al.vertical ? V_ALIGN[al.vertical] : undefined;
    if (vv) s.vt = vv;
    if (al.wrapText) s.tb = 3; // WrapStrategy.WRAP
    if (typeof al.textRotation === "number" && al.textRotation) s.tr = { a: al.textRotation, v: 0 };
    // Preserve cell indentation (Excel `alignment.indent`, one unit ≈ 3 chars). This is
    // how PIVOT TABLES in the default "Compact" layout render their nested row-label
    // hierarchy (feature #5) — without it the grouped labels collapse to a flat column
    // and the pivot looks nothing like the original. Univer has no outline-grouping in
    // the free preset, but cell left-padding (`pd.l`) faithfully reproduces the indent.
    if (typeof al.indent === "number" && al.indent > 0) s.pd = { ...s.pd, l: Math.min(al.indent, 15) * 10 };
  }
  const b = cell.border;
  if (b) {
    const bd: NonNullable<UStyle["bd"]> = {};
    const t = edge(b.top), bo = edge(b.bottom), l = edge(b.left), r = edge(b.right);
    if (t) bd.t = t;
    if (bo) bd.b = bo;
    if (l) bd.l = l;
    if (r) bd.r = r;
    if (Object.keys(bd).length) s.bd = bd;
  }
  if (cell.numFmt) s.n = { pattern: cell.numFmt };
  else if (isDate) s.n = { pattern: "yyyy-mm-dd" }; // date with no explicit format

  // Preserve source colours EXACTLY for Excel fidelity: white text on no fill is
  // invisible in Excel (a deliberate hidden label) and must stay invisible here.
  // (An earlier guard dropped near-white fonts to "un-hide" them, but that made
  // intentionally-hidden white labels appear — non-Excel behaviour. The theme
  // colour swap now resolves genuine dark text correctly, so the guard is gone.)
  return s;
}

/* -------------------------------------------------------------------------- */
/* Merges                                                                     */
/* -------------------------------------------------------------------------- */
interface Merge { startRow: number; startColumn: number; endRow: number; endColumn: number }

/** Parse a cell address ("AB12") to 0-based { row, col }. */
function parseAddr(addr: string): { row: number; col: number } | null {
  const m = /^([A-Z]+)(\d+)$/.exec(addr.trim().toUpperCase().replace(/\$/g, ""));
  if (!m) return null;
  let col = 0;
  for (const ch of m[1]) col = col * 26 + (ch.charCodeAt(0) - 64);
  return { row: Number(m[2]) - 1, col: col - 1 };
}

/** Auto-detect a plain-text email / URL cell and return a link target, so it
 *  renders clickable even when the file stored no hyperlink (Excel linkifies as
 *  you type; exports often lose it). Only fires when the ENTIRE cell value is a
 *  single address — NOT multi-address lists ("a@x.com, b@y.com") or prose. */
function autoLink(v: string | number | boolean | undefined): string | null {
  if (typeof v !== "string") return null;
  const s = v.trim();
  if (/^[^\s@,;]+@[^\s@,;]+\.[^\s@,;]{2,}$/.test(s)) return `mailto:${s}`;      // single email
  if (/^https?:\/\/\S+$/i.test(s)) return s;                                     // http(s) URL
  if (/^www\.[^\s]+\.[^\s]{2,}$/i.test(s)) return `https://${s}`;                // bare www URL
  return null;
}

/** ExcelJS `autoFilter` (string "A1:N100" or {from,to}) → 0-based Univer range. */
function parseAutoFilterRef(af: XAutoFilter | undefined): URange | null {
  if (!af) return null;
  const toRC = (v: XAddr): { row: number; col: number } | null =>
    typeof v === "string" ? parseAddr(v) : v && typeof v.row === "number" ? { row: v.row - 1, col: v.column - 1 } : null;
  let from: { row: number; col: number } | null = null;
  let to: { row: number; col: number } | null = null;
  if (typeof af === "string") {
    const [a, b] = af.split(":");
    from = parseAddr(a);
    to = parseAddr(b ?? a);
  } else {
    from = af.from ? toRC(af.from) : null;
    to = af.to ? toRC(af.to) : null;
  }
  if (!from || !to) return null;
  return { startRow: from.row, startColumn: from.col, endRow: to.row, endColumn: to.col };
}

/* ---- Conditional formatting (ExcelJS rule-sets → Univer CF resource) ------- */
const CF_NUM_OPS = new Set(["greaterThan", "greaterThanOrEqual", "lessThan", "lessThanOrEqual", "equal", "notEqual", "between", "notBetween"]);

/** ExcelJS differential style (dxf) → Univer style short-keys (the format the
 *  rule paints when it matches). */
function cfStyle(dxf?: XCfRule["style"]): Record<string, unknown> {
  const s: Record<string, unknown> = {};
  const font = dxf?.font;
  if (font) {
    if (font.bold) s.bl = 1;
    if (font.italic) s.it = 1;
    const fc = colorToHex(font.color);
    if (fc) s.cl = { rgb: fc };
  }
  const fill = dxf?.fill;
  if (fill && (fill.type === "pattern" || fill.pattern)) {
    const bg = colorToHex(fill.bgColor) ?? colorToHex(fill.fgColor); // CF fills carry the colour in bgColor
    if (bg) s.bg = { rgb: bg };
  }
  return s;
}

/** A CF `ref` ("A1:A10 C1:C10", with $) → 0-based Univer ranges. */
function parseRanges(ref?: string): URange[] {
  if (!ref) return [];
  const out: URange[] = [];
  for (const part of ref.trim().split(/\s+/)) {
    const [a, b] = part.split(":");
    const from = parseAddr(a);
    const to = parseAddr(b ?? a);
    if (from && to) out.push({ startRow: from.row, startColumn: from.col, endRow: to.row, endColumn: to.col });
  }
  return out;
}
const mapCfValue = (v?: XCfValue) => ({ type: v?.type ?? "num", value: v?.value });

/** Map a worksheet's ExcelJS conditional-formatting rule-sets to Univer CF
 *  rules (highlightCell / colorScale / dataBar). */
function readConditionalFormats(ws: XWorksheet, sheetId: string, nextId: () => number): Array<Record<string, unknown>> {
  const out: Array<Record<string, unknown>> = [];
  const colorsOf = (r: XCfRule): (Argb | undefined)[] => (Array.isArray(r.color) ? r.color : r.color ? [r.color] : []);
  for (const set of ws.conditionalFormattings ?? []) {
    const ranges = parseRanges(set.ref);
    if (!ranges.length) continue;
    for (const r of set.rules ?? []) {
      const style = cfStyle(r.style);
      let rule: Record<string, unknown> | null = null;
      if (r.type === "cellIs" && r.operator && CF_NUM_OPS.has(r.operator)) {
        const f = (r.formulae ?? []).map(Number);
        if (r.operator === "between" || r.operator === "notBetween") {
          if (!Number.isNaN(f[0]) && !Number.isNaN(f[1])) rule = { type: "highlightCell", subType: "number", operator: r.operator, value: [f[0], f[1]], style };
        } else if (!Number.isNaN(f[0])) {
          rule = { type: "highlightCell", subType: "number", operator: r.operator, value: f[0], style };
        }
      } else if (r.type === "expression") {
        const f = (r.formulae ?? [])[0];
        if (f != null) { const fs = String(f); rule = { type: "highlightCell", subType: "formula", value: fs.startsWith("=") ? fs : `=${fs}`, style }; }
      } else if (r.type === "containsText") {
        const opMap: Record<string, string> = { containsText: "containsText", notContains: "notContainsText", beginsWith: "beginsWith", endsWith: "endsWith" };
        rule = { type: "highlightCell", subType: "text", operator: opMap[r.operator ?? "containsText"] ?? "containsText", value: r.text ?? String((r.formulae ?? [])[0] ?? ""), style };
      } else if (r.type === "colorScale" && r.cfvo) {
        const cs = colorsOf(r);
        rule = { type: "colorScale", config: r.cfvo.map((v, i) => ({ index: i, color: colorToHex(cs[i]) ?? "#FFFFFF", value: mapCfValue(v) })) };
      } else if (r.type === "dataBar" && r.cfvo) {
        const bar = colorToHex(colorsOf(r)[0]) ?? "#638EC6";
        rule = { type: "dataBar", isShowValue: true, config: { min: mapCfValue(r.cfvo[0]), max: mapCfValue(r.cfvo[1]), isGradient: true, positiveColor: bar, nativeColor: bar } };
      }
      if (rule) out.push({ ranges, cfId: `${sheetId}_cf_${nextId()}`, stopIfTrue: false, rule });
    }
  }
  return out;
}

/** Read a worksheet's merges as 0-based ranges + a Set of slave "r:c" keys. */
function readMerges(ws: XWorksheet): { merges: Merge[]; slaves: Set<string> } {
  const merges: Merge[] = [];
  const slaves = new Set<string>();
  const raw = ws.model?.merges ?? [];
  for (const range of raw) {
    const [a, b] = range.split(":");
    const s = parseAddr(a);
    const e = parseAddr(b ?? a);
    if (!s || !e) continue;
    const startRow = Math.min(s.row, e.row), endRow = Math.max(s.row, e.row);
    const startColumn = Math.min(s.col, e.col), endColumn = Math.max(s.col, e.col);
    merges.push({ startRow, startColumn, endRow, endColumn });
    for (let r = startRow; r <= endRow; r++)
      for (let c = startColumn; c <= endColumn; c++)
        if (!(r === startRow && c === startColumn)) slaves.add(`${r}:${c}`);
  }
  return { merges, slaves };
}

/* ---- Dimensions ----------------------------------------------------------- */
const ROW_HEADROOM = 30;
const COL_HEADROOM = 4;
const MIN_ROWS = 100;
const MIN_COLS = 26;
const DEFAULT_COL_WIDTH = 88; // Excel-ish default (~8.43 chars)
const DEFAULT_ROW_HEIGHT_PX = 24; // Univer default row height
const EMU_PER_PX = 9525; // Excel drawing units (English Metric Units) per pixel
/** Excel column width (char units) → pixels. */
const charWidthToPx = (w: number) => Math.round(w * 7 + 5);
/** Excel row height (points) → pixels. */
const ptToPx = (pt: number) => Math.round(pt * (96 / 72));

/** Browser-safe bytes → base64 (chunked to avoid arg-count limits). */
function bytesToBase64(bytes: Uint8Array): string {
  let bin = "";
  const chunk = 0x8000;
  for (let i = 0; i < bytes.length; i += chunk) bin += String.fromCharCode(...bytes.subarray(i, i + chunk));
  return btoa(bin);
}
function toU8(buf: Uint8Array | ArrayBuffer | { data?: number[] } | undefined): Uint8Array | null {
  if (!buf) return null;
  if (buf instanceof Uint8Array) return buf;
  if (buf instanceof ArrayBuffer) return new Uint8Array(buf);
  if (Array.isArray((buf as { data?: number[] }).data)) return new Uint8Array((buf as { data: number[] }).data);
  return null;
}

/* -------------------------------------------------------------------------- */
/* Helpers for inserting imported sheets into an EXISTING workbook            */
/* -------------------------------------------------------------------------- */

/** Map original worksheet index → sheet name, from a converted snapshot. */
export function sheetIndexToName(snapshot: WorkbookData): Record<number, string> {
  const order = (snapshot.sheetOrder as string[] | undefined) ?? [];
  const sheets = (snapshot.sheets as Record<string, { name?: string }> | undefined) ?? {};
  const map: Record<number, string> = {};
  order.forEach((id, i) => { map[i] = sheets[id]?.name ?? `Sheet${i + 1}`; });
  return map;
}

/**
 * Resolve a sheet's registry style ids to INLINE style objects. Needed when a
 * sheet from an imported snapshot is inserted into a different workbook (via the
 * Facade `create(..., { sheet })`) whose style registry doesn't hold those ids —
 * otherwise every styled cell would lose its formatting.
 */
export function inlineSheetStyles(sheet: Record<string, unknown>, styles: Record<string, unknown>): Record<string, unknown> {
  const cellData = sheet.cellData as Record<string, Record<string, { s?: unknown } | null>> | undefined;
  if (!cellData) return sheet;
  const out: Record<number, Record<number, unknown>> = {};
  for (const [r, cols] of Object.entries(cellData)) {
    const row: Record<number, unknown> = {};
    for (const [c, cell] of Object.entries(cols)) {
      if (cell && typeof cell === "object" && typeof (cell as { s?: unknown }).s === "string") {
        row[Number(c)] = { ...cell, s: styles[(cell as { s: string }).s] ?? undefined };
      } else {
        row[Number(c)] = cell;
      }
    }
    out[Number(r)] = row;
  }
  return { ...sheet, cellData: out };
}

/** The first VISIBLE sheet of a converted snapshot (or the first, if all hidden). */
export function firstVisibleSheet(snapshot: WorkbookData): Record<string, unknown> | null {
  const order = (snapshot.sheetOrder as string[] | undefined) ?? [];
  const sheets = (snapshot.sheets as Record<string, Record<string, unknown>> | undefined) ?? {};
  for (const id of order) {
    const s = sheets[id];
    if (s && s.hidden !== 1) return s;
  }
  return order.length ? sheets[order[0]] ?? null : null;
}

/**
 * Build a rich in-place placement of ONE imported sheet at a (row,col) offset:
 * a sparse ICellData matrix (styles inlined) + shifted merge rectangles. Used to
 * bring imported content into an EXISTING sheet (replace-current / append /
 * at-cell) with full formatting — not just plain values.
 */
export function buildRichPlacement(
  sheet: Record<string, unknown>,
  styles: Record<string, unknown>,
  rowOffset: number,
  colOffset: number,
): { cells: Record<number, Record<number, unknown>>; merges: Array<{ row: number; col: number; numRows: number; numCols: number }>; maxRow: number; maxCol: number } {
  const resolved = inlineSheetStyles(sheet, styles);
  const cd = (resolved.cellData as Record<string, Record<string, unknown>> | undefined) ?? {};
  const cells: Record<number, Record<number, unknown>> = {};
  let maxRow = 0;
  let maxCol = 0;
  for (const [r, cols] of Object.entries(cd)) {
    const rr = Number(r) + rowOffset;
    for (const [c, cell] of Object.entries(cols)) {
      const cc = Number(c) + colOffset;
      (cells[rr] ??= {})[cc] = cell;
      if (rr > maxRow) maxRow = rr;
      if (cc > maxCol) maxCol = cc;
    }
  }
  const mergeData = (sheet.mergeData as Array<{ startRow: number; startColumn: number; endRow: number; endColumn: number }> | undefined) ?? [];
  const merges = mergeData.map((m) => ({
    row: m.startRow + rowOffset,
    col: m.startColumn + colOffset,
    numRows: m.endRow - m.startRow + 1,
    numCols: m.endColumn - m.startColumn + 1,
  }));
  return { cells, merges, maxRow, maxCol };
}

/* ---- Facade image placement (shared by whole-doc render + sheet insert) ---- */
interface FImgBuilder {
  setSource(src: string, type?: unknown): FImgBuilder;
  setColumn(c: number): FImgBuilder;
  setRow(r: number): FImgBuilder;
  setWidth(w: number): FImgBuilder;
  setHeight(h: number): FImgBuilder;
  buildAsync(): Promise<unknown>;
}
interface FImgSheet {
  getSheetName?: () => string;
  newOverGridImage?: () => FImgBuilder;
  insertImages?: (images: unknown[]) => void;
  insertImage?: (url: string, column?: number, row?: number, offsetX?: number, offsetY?: number) => Promise<boolean>;
}
interface FImgApi {
  getActiveWorkbook?: () => { getSheets?: () => FImgSheet[] } | null;
  Enum?: { ImageSourceType?: { BASE64?: unknown } };
}

/**
 * Place imported floating images on their target sheets via the Facade, matching
 * by sheet NAME (robust to index shifts when inserting into an existing
 * workbook). Best-effort — logs a summary and never throws.
 */
export async function placeImportImages(api: unknown, images: ImportImage[], indexToName: Record<number, string>): Promise<void> {
  try {
    const a = api as FImgApi;
    const sheets = a.getActiveWorkbook?.()?.getSheets?.() ?? [];
    const byName = new Map<string, FImgSheet>();
    for (const s of sheets) {
      const n = s.getSheetName?.();
      if (n) byName.set(n, s);
    }
    const sourceType = a.Enum?.ImageSourceType?.BASE64 ?? "BASE64";
    const groups = new Map<string, ImportImage[]>();
    for (const im of images) {
      const name = indexToName[im.sheetIndex];
      if (!name) continue;
      const list = groups.get(name) ?? [];
      list.push(im);
      groups.set(name, list);
    }
    let placed = 0;
    for (const [name, ims] of groups) {
      const fws = byName.get(name);
      if (!fws) continue;
      if (fws.newOverGridImage && fws.insertImages) {
        const built: unknown[] = [];
        for (const im of ims) {
          try {
            const b = await fws.newOverGridImage().setSource(im.base64, sourceType).setColumn(im.col).setRow(im.row).setWidth(im.width).setHeight(im.height).buildAsync();
            if (b) built.push(b);
          } catch (e) {
            console.warn("[levich] image build failed", e);
          }
        }
        if (built.length) {
          try {
            fws.insertImages(built);
            placed += built.length;
          } catch (e) {
            console.warn("[levich] insertImages failed", e);
          }
        }
      } else if (fws.insertImage) {
        for (const im of ims) {
          try {
            await fws.insertImage(im.base64, im.col, im.row, im.colOffset, im.rowOffset);
            placed += 1;
          } catch (e) {
            console.warn("[levich] insertImage failed", e);
          }
        }
      }
    }
    console.info(`[levich] imported images: ${images.length} found, ${placed} placed`);
  } catch (e) {
    console.warn("[levich] placeImportImages error", e);
  }
}

/* -------------------------------------------------------------------------- */
/* Public: convert a File / ArrayBuffer to a Univer snapshot                  */
/* -------------------------------------------------------------------------- */
export async function parseXlsxToSnapshot(file: File): Promise<ParsedSnapshot> {
  const ExcelJS = (await import("exceljs")).default;
  const wb = new ExcelJS.Workbook();
  // Keep the raw bytes: ExcelJS discards the pivot-table XML parts, so we re-read
  // them from the ZIP (jszip) to reconstruct interactive pivots (see pivot-import).
  const rawBytes = await file.arrayBuffer();
  await wb.xlsx.load(rawBytes);

  // One workbook-wide style registry; cells reference styles by id (keeps the
  // snapshot small when a style repeats — headers, fills, etc.).
  const styles: Record<string, UStyle> = {};
  const styleIndex = new Map<string, string>();
  let styleSeq = 0;
  const internStyle = (style: UStyle): string | undefined => {
    if (!style || Object.keys(style).length === 0) return undefined;
    const key = JSON.stringify(style);
    let id = styleIndex.get(key);
    if (!id) {
      id = `s${++styleSeq}`;
      styleIndex.set(key, id);
      styles[id] = style;
    }
    return id;
  };

  const WORKBOOK_ID = "levich-imported";
  const sheetOrder: string[] = [];
  const sheets: Record<string, unknown> = {};
  const images: ImportImage[] = [];
  // Native Univer drawing resource, keyed by sheetId → { data, order }. Embedded
  // in the snapshot so images render at LOAD time (the post-load Facade insert
  // proved unreliable).
  const drawingResource: Record<string, { data: Record<string, unknown>; order: string[] }> = {};
  // AutoFilter → SHEET_FILTER_PLUGIN resource, keyed by sheetId → { ref }.
  const filterResource: Record<string, { ref: URange }> = {};
  // Conditional formatting → SHEET_CONDITIONAL_FORMATTING_PLUGIN, keyed by sheetId.
  const cfResource: Record<string, Array<Record<string, unknown>>> = {};
  let cfSeq = 0;
  // Cell hyperlinks (emails / URLs) → SHEET_HYPER_LINK_PLUGIN, keyed by sheetId.
  const linkResource: Record<string, { links: Array<{ id: string; payload: string; display?: string; row: number; column: number }> }> = {};
  let linkSeq = 0;
  const wbx = wb as unknown as XWorkbook;
  const media = wbx.media ?? wbx.model?.media ?? [];

  (wb.worksheets as unknown as XWorksheet[]).forEach((ws, index) => {
    const sheetId = `sheet_${index + 1}`;
    sheetOrder.push(sheetId);

    const { merges, slaves } = readMerges(ws);
    // Top-left cells of merges that span MULTIPLE ROWS — these already have
    // vertical room across their rows, so we must NOT bump their row height for
    // a big font / wrapped text (that inflates the merged box, e.g. a title box).
    const multiRowMergeTL = new Set(merges.filter((m) => m.endRow > m.startRow).map((m) => `${m.startRow}:${m.startColumn}`));
    const cellData: Record<number, Record<number, UCell>> = {};
    // Wrapped cells (for auto-fit row-height estimation): text + font size + col.
    const wrapInfo: Array<{ r: number; c: number; text: string; fs: number }> = [];
    // Largest font size seen per row (for large-font row-height estimation — a
    // 20pt title in a default 24px row would otherwise clip at the top).
    const rowMaxFs: Record<number, number> = {};
    let maxRow = 0;
    let maxCol = 0;

    ws.eachRow({ includeEmpty: true }, (row) => {
      const r = row.number - 1;
      row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
        const c = colNumber - 1;
        if (slaves.has(`${r}:${c}`)) return; // merge slave — leave empty
        const { v, f, isDate } = cellValue(cell);
        const style = buildStyle(cell, !!isDate);
        // Embedded newlines (Alt+Enter / programmatic exports like Carta) are
        // multi-line even when wrapText isn't set — Excel/Sheets render them on
        // separate lines and grow the row. Force wrap so Univer honours the
        // breaks and the row-height estimator (below) makes room.
        if (typeof v === "string" && v.includes("\n")) style.tb = 3;
        const skipHeightBump = multiRowMergeTL.has(`${r}:${c}`); // merge already spans rows
        if (style.tb === 3 && v !== undefined && v !== "" && !skipHeightBump) {
          wrapInfo.push({ r, c, text: String(v), fs: style.fs ?? 11 }); // for auto-fit
        }
        if (v !== undefined && v !== "" && style.fs && !skipHeightBump) rowMaxFs[r] = Math.max(rowMaxFs[r] ?? 0, style.fs);
        // An EMPTY cell that still carries a date/number format makes Univer
        // format "nothing" → renders "NaN". Drop the format when there's no
        // value or formula to display.
        if (v === undefined && f === undefined && style.n) delete style.n;
        const sid = internStyle(style);
        // Skip cells that carry neither a value/formula nor any styling.
        if (v === undefined && f === undefined && !sid) return;
        const out: UCell = {};
        if (f !== undefined) out.f = f;
        if (v !== undefined) out.v = v;
        if (sid) out.s = sid;
        (cellData[r] ??= {})[c] = out;
        // Cell hyperlink → Univer link resource so it renders clickable like
        // Excel. Prefer the STORED link; otherwise auto-detect a plain-text
        // email / URL cell and linkify it (FinSheets capability beyond the file).
        const rawLink = typeof cell.hyperlink === "string" ? cell.hyperlink : cell.hyperlink?.hyperlink;
        const link = rawLink || autoLink(v);
        if (link) {
          (linkResource[sheetId] ??= { links: [] }).links.push({
            id: `${sheetId}_hl_${linkSeq++}`,
            payload: link,
            display: v !== undefined ? String(v) : link,
            row: r,
            column: c,
          });
        }
        if (r > maxRow) maxRow = r;
        if (c > maxCol) maxCol = c;
      });
    });

    // Column widths + hidden state. `hd: 1` hides the column (matches Excel/
    // Sheets) — critical for layout fidelity, since many models hide helper
    // columns and rendering them all distorts the sheet.
    const columnData: Record<number, { w?: number; hd?: number }> = {};
    (ws.columns ?? []).forEach((col, i) => {
      const idx = (col?.number ?? i + 1) - 1;
      const entry: { w?: number; hd?: number } = {};
      if (col?.width && col.width > 0) entry.w = charWidthToPx(col.width);
      if (col?.hidden) entry.hd = 1;
      if (entry.w !== undefined || entry.hd !== undefined) columnData[idx] = entry;
    });

    // Row heights. Rows with an explicit stored height are pinned as-is. For
    // rows whose wrapped text has NO stored height (Excel/Sheets auto-fit them
    // but don't persist the computed height), ESTIMATE the height from the text
    // length vs. the available column width — Univer's own `ia` auto-fit isn't
    // computed at snapshot-load, so an explicit height is the only reliable way
    // to stop multi-line headers / notes from rendering clipped.
    // includeEmpty:true so we capture the hidden/height state of EVERY row —
    // including fill-only "divider" rows (a solid fill, no text). With
    // includeEmpty:false those rows are skipped, so a hidden black-fill divider
    // leaks through as a black bar instead of staying hidden.
    const rowData: Record<number, { h?: number; hd?: number }> = {};
    ws.eachRow({ includeEmpty: true }, (row) => {
      const entry: { h?: number; hd?: number } = {};
      if (row.height && row.height > 0) entry.h = ptToPx(row.height);
      if (row.hidden) entry.hd = 1;
      if (entry.h !== undefined || entry.hd !== undefined) rowData[row.number - 1] = entry;
    });
    {
      const colW = (c: number) => columnData[c]?.w ?? DEFAULT_COL_WIDTH;
      const spanCols = new Map<string, number>();
      for (const m of merges) spanCols.set(`${m.startRow}:${m.startColumn}`, m.endColumn - m.startColumn + 1);
      const neededLines: Record<number, number> = {};
      const rowFs: Record<number, number> = {};
      for (const info of wrapInfo) {
        if (rowData[info.r]?.h) continue; // explicit height wins
        let widthPx = colW(info.c);
        const span = spanCols.get(`${info.r}:${info.c}`);
        if (span && span > 1) { widthPx = 0; for (let k = 0; k < span; k++) widthPx += colW(info.c + k); }
        const avgChar = info.fs * 0.56; // ≈ px per character
        const charsPerLine = Math.max(1, Math.floor((widthPx - 8) / avgChar));
        let lines = 0;
        for (const seg of info.text.split("\n")) lines += Math.max(1, Math.ceil(seg.length / charsPerLine));
        neededLines[info.r] = Math.max(neededLines[info.r] ?? 1, lines);
        rowFs[info.r] = Math.max(rowFs[info.r] ?? 11, info.fs);
      }
      for (const [rStr, lines] of Object.entries(neededLines)) {
        const r = Number(rStr);
        if (lines > 1 && !rowData[r]?.h) {
          const lineH = Math.round((rowFs[r] ?? 11) * 1.5); // comfortable line height
          rowData[r] = { ...rowData[r], h: Math.min(600, lines * lineH + 6) };
        }
      }
      // Large-font single-line rows (titles/headings): the default 24px row clips
      // a ~20pt+ title. Grow the row to fit the tallest font when no other height
      // has been set. ptToPx(fs) is the glyph box; + padding for comfort.
      for (const [rStr, fs] of Object.entries(rowMaxFs)) {
        const r = Number(rStr);
        if (rowData[r]?.h) continue; // explicit / wrap height already wins
        const needed = ptToPx(fs) + 8;
        if (needed > DEFAULT_ROW_HEIGHT_PX) rowData[r] = { ...rowData[r], h: Math.min(600, needed) };
      }
    }

    // Floating images (logos, equation graphics, etc.). Excel stores these as
    // drawings anchored to cells with EMU offsets; extract source + anchor +
    // pixel size so the Facade can place them after load.
    try {
      // Hidden rows/columns collapse to 0 in Univer (and Excel), so image anchor
      // math MUST treat them as 0 — otherwise the summed pixel position is
      // inflated by every hidden row/col above/left, scattering the images.
      const colPx = (c: number) => (columnData[c]?.hd ? 0 : columnData[c]?.w ?? DEFAULT_COL_WIDTH);
      const rowPx = (r: number) => (rowData[r]?.hd ? 0 : rowData[r]?.h ?? DEFAULT_ROW_HEIGHT_PX);
      const sumW = (upto: number) => { let s = 0; for (let c = 0; c < upto; c++) s += colPx(c); return s; };
      const sumH = (upto: number) => { let s = 0; for (let r = 0; r < upto; r++) s += rowPx(r); return s; };
      const posAtX = (px: number) => { let acc = 0, i = 0; while (acc + colPx(i) <= px && i < 16384) { acc += colPx(i); i++; } return { column: i, columnOffset: Math.round(px - acc) }; };
      const posAtY = (px: number) => { let acc = 0, i = 0; while (acc + rowPx(i) <= px && i < 1048576) { acc += rowPx(i); i++; } return { row: i, rowOffset: Math.round(px - acc) }; };
      let imgIdx = 0;
      for (const img of ws.getImages?.() ?? []) {
        const m = media[Number(img.imageId)];
        if (!m) continue;
        const ext = (m.extension || "png").toLowerCase();
        const mime = ext === "jpg" ? "jpeg" : ext;
        // Prefer an inline base64 if ExcelJS provides one, else encode the buffer.
        let base64: string;
        if (typeof m.base64 === "string" && m.base64) {
          base64 = m.base64.startsWith("data:") ? m.base64 : `data:image/${mime};base64,${m.base64}`;
        } else {
          const bytes = toU8(m.buffer);
          if (!bytes) continue;
          base64 = `data:image/${mime};base64,${bytesToBase64(bytes)}`;
        }
        const tl = img.range?.tl ?? {};
        const br = img.range?.br;
        const ext2 = img.range?.ext;
        const col = tl.nativeCol ?? tl.col ?? 0;
        const row = tl.nativeRow ?? tl.row ?? 0;
        const colOffset = (tl.nativeColOff ?? 0) / EMU_PER_PX;
        const rowOffset = (tl.nativeRowOff ?? 0) / EMU_PER_PX;
        let width = 120;
        let height = 60;
        if (br && typeof br.nativeCol === "number" && typeof br.nativeRow === "number") {
          // twoCellAnchor: size = span from top-left to bottom-right anchor.
          let w = 0;
          for (let c = col; c < br.nativeCol; c++) w += colPx(c);
          width = w - colOffset + (br.nativeColOff ?? 0) / EMU_PER_PX;
          let h = 0;
          for (let r = row; r < br.nativeRow; r++) h += rowPx(r);
          height = h - rowOffset + (br.nativeRowOff ?? 0) / EMU_PER_PX;
        } else if (ext2?.width && ext2?.height) {
          // oneCellAnchor: ExcelJS already converts `ext` from EMU to PIXELS
          // (unlike the col/row offsets, which stay in EMU). Use it directly.
          width = ext2.width;
          height = ext2.height;
        }
        const wPx = Math.max(8, Math.round(width));
        const hPx = Math.max(8, Math.round(height));
        const colOff = Math.max(0, Math.round(colOffset));
        const rowOff = Math.max(0, Math.round(rowOffset));
        images.push({ sheetIndex: index, base64, col, row, colOffset: colOff, rowOffset: rowOff, width: wPx, height: hPx });

        // Native drawing object (rendered at load via the SHEET_DRAWING_PLUGIN
        // resource). transform = absolute px box; sheetTransform = from/to cell
        // anchors so it tracks the grid.
        const left = Math.round(sumW(col) + colOffset);
        const top = Math.round(sumH(row) + rowOffset);
        const toX = posAtX(left + wPx);
        const toY = posAtY(top + hPx);
        const from = { column: col, columnOffset: colOff, row, rowOffset: rowOff };
        const to = { column: toX.column, columnOffset: toX.columnOffset, row: toY.row, rowOffset: toY.rowOffset };
        const drawingId = `img_${index}_${imgIdx++}`;
        const drawing = {
          unitId: WORKBOOK_ID,
          subUnitId: sheetId,
          drawingId,
          drawingType: 0, // DrawingTypeEnum.DRAWING_IMAGE
          imageSourceType: "BASE64",
          source: base64,
          transform: { left, top, width: wPx, height: hPx, angle: 0, scaleX: 1, scaleY: 1, skewX: 0, skewY: 0, flipX: false, flipY: false },
          sheetTransform: { from: { ...from }, to: { ...to } },
          axisAlignSheetTransform: { from: { ...from }, to: { ...to } },
          anchorType: "1", // SheetDrawingAnchorType.Both
        };
        (drawingResource[sheetId] ??= { data: {}, order: [] });
        drawingResource[sheetId].data[drawingId] = drawing;
        drawingResource[sheetId].order.push(drawingId);
      }
    } catch (e) {
      console.warn("[levich] image extraction failed", e);
    }

    // Frozen panes.
    const view = ws.views?.find((vw) => vw.state === "frozen");
    const xSplit = view?.xSplit ?? 0;
    const ySplit = view?.ySplit ?? 0;
    const freeze = { xSplit, ySplit, startRow: ySplit, startColumn: xSplit };

    // Tab colour + gridline visibility (Excel hides gridlines on dashboards).
    // Only honour an EXPLICIT, non-white argb tab colour — theme-based / white
    // values are usually a "no colour" artifact that would paint tabs white.
    const tc = ws.properties?.tabColor?.argb ? argbToHex(ws.properties.tabColor.argb) : null;
    const tabColor = tc && tc !== "#FFFFFF" ? tc : "";
    const showGridlines = ws.views?.some((vw) => vw.showGridLines === false) ? 0 : 1;

    // AutoFilter → filter funnels on the header row of the range.
    const afRef = parseAutoFilterRef(ws.autoFilter);
    if (afRef) filterResource[sheetId] = { ref: afRef };

    // Conditional formatting rules for this sheet.
    const cfRules = readConditionalFormats(ws, sheetId, () => cfSeq++);
    if (cfRules.length) cfResource[sheetId] = cfRules;

    const rowCount = Math.max(maxRow + 1, ws.rowCount || 0, MIN_ROWS) + ROW_HEADROOM;
    const columnCount = Math.max(maxCol + 1, ws.actualColumnCount || ws.columnCount || 0, MIN_COLS) + COL_HEADROOM;

    sheets[sheetId] = {
      id: sheetId,
      name: ws.name || `Sheet${index + 1}`,
      rowCount,
      columnCount,
      cellData,
      columnData,
      rowData,
      // Match the source's default row height / column width so rows & columns
      // with no explicit size render at the same pixel size as Excel/Sheets
      // (Univer's own default is taller → merged boxes came out oversized).
      defaultRowHeight: ws.properties?.defaultRowHeight ? ptToPx(ws.properties.defaultRowHeight) : DEFAULT_ROW_HEIGHT_PX,
      defaultColumnWidth: ws.properties?.defaultColWidth ? charWidthToPx(ws.properties.defaultColWidth) : DEFAULT_COL_WIDTH,
      mergeData: merges,
      freeze,
      tabColor,
      showGridlines,
      // Respect Excel's hidden/veryHidden sheets — imported (data preserved)
      // but not shown as tabs, matching how the source app displays them.
      hidden: ws.state && ws.state !== "visible" ? 1 : 0,
    };
  });

  // A workbook with no worksheets — hand back a single empty sheet.
  if (sheetOrder.length === 0) {
    const id = "sheet_1";
    sheetOrder.push(id);
    sheets[id] = { id, name: "Sheet1", rowCount: MIN_ROWS, columnCount: MIN_COLS, cellData: {} };
  }

  const baseName = file.name.replace(/\.(xlsx|xls)$/i, "") || "Imported";
  // Embed images as Univer's native drawing resource so they render at load.
  const resources: Array<{ name: string; data: string }> = [];
  if (Object.keys(drawingResource).length) resources.push({ name: "SHEET_DRAWING_PLUGIN", data: JSON.stringify(drawingResource) });
  if (Object.keys(filterResource).length) resources.push({ name: "SHEET_FILTER_PLUGIN", data: JSON.stringify(filterResource) });
  if (Object.keys(cfResource).length) resources.push({ name: "SHEET_CONDITIONAL_FORMATTING_PLUGIN", data: JSON.stringify(cfResource) });
  if (Object.keys(linkResource).length) resources.push({ name: "SHEET_HYPER_LINK_PLUGIN", data: JSON.stringify(linkResource) });

  const snapshot: ParsedSnapshot = {
    id: WORKBOOK_ID,
    name: baseName,
    sheetOrder,
    styles,
    sheets,
    resources,
    // Non-Univer field: floating images for the Facade path (used when inserting
    // imported sheets into an EXISTING workbook, which bypasses the snapshot
    // resource). Univer ignores unknown top-level keys.
    drawingsImport: images,
  } as ParsedSnapshot;

  // Non-Univer escape-hatch (mirrors `drawingsImport`): reconstruct any pivot
  // tables into interactive `PivotSource` + `PivotSpec` so a host can open them
  // in `pivotInteractive` mode. Best-effort — never blocks the import.
  try {
    const { parsePivotsFromXlsx } = await import("./pivot-import");
    const pivots = await parsePivotsFromXlsx(rawBytes, snapshot);
    if (pivots.length) snapshot.pivotsImport = pivots;
  } catch (e) {
    console.warn("[levich] pivot reconstruction failed", e);
  }

  return snapshot;
}
