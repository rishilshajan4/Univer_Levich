/**
 * LevichToolbar — our own branded operations toolbar (UI/UX only). Every control
 * drives Univer through its Facade API (the same operations Univer's built-in
 * toolbar triggers), so the engine/"BE" is untouched. Univer's toolbar is hidden
 * (`univerToolbar: false`); the formula bar + grid stay.
 *
 * Untitled UI icons · our portal dropdowns (no clipping in the scroll strip) ·
 * our tooltips · ‹ › horizontal paging.
 */
import { useEffect, useRef, useState, type CSSProperties, type ReactNode } from "react";
import { createPortal } from "react-dom";
import { Z_BASE } from "../core/z-index";
import {
  AlignBottom01,
  AlignCenter,
  AlignLeft,
  AlignRight,
  Bold01,
  ChevronDown,
  CurrencyDollar,
  Droplets01,
  FilterFunnel01,
  FlipBackward,
  FlipForward,
  Grid01,
  IntersectSquare,
  Italic01,
  Link01,
  PaintPour,
  Percent01,
  Printer,
  Rows01,
  SearchMd,
  Strikethrough01,
  Type01,
  Underline01,
} from "@untitledui/icons";
import type { UniverAPI } from "../core/create-sheet";
import { GOOGLE_FONTS, ensureGoogleFont, ensureGoogleFonts } from "../theme/google-fonts";
import { applyFilterView, clearAllFilters, ensureFilter, loadViews, saveViews, snapshotFilters, type FilterView, type FilterViewApi } from "../features/filter-views";
import { buildGroupView, clearGroupView, getGroupColumns, loadGroupViews, saveGroupViews, toggleGroup, type GroupByApi, type GroupRun, type GroupView } from "../features/group-by-view";
import { insertAggregate, insertFunctionTemplate, type Aggregate, type FunctionsApi } from "../features/functions";
import { printSheet } from "../core/print-sheet";
import { ViewsMenu } from "./views-menu";
import { FunctionsMenu } from "./functions-menu";

/* ---- Loose Facade views (avoid branded-type friction) --------------------- */
interface CellStyle {
  bl?: number;
  it?: number;
  ul?: { s?: number };
  st?: { s?: number };
  tr?: { a?: number; v?: number };
}
interface RangeOps {
  getCellStyleData(): CellStyle | null;
  getFontSize(): number | null;
  getFontFamily(): string | null;
  getRow?(): number;
  getColumn?(): number;
  getWrap(): boolean;
  getNumberFormat(): string;
  setFontWeight(w: string | null): RangeOps;
  setFontStyle(s: string | null): RangeOps;
  setFontLine(l: string | null): RangeOps;
  setFontFamily(f: string): RangeOps;
  setFontSize(n: number): RangeOps;
  setFontColor(c: string): RangeOps;
  setBackground(c: string): RangeOps;
  setHorizontalAlignment(a: string): RangeOps;
  setVerticalAlignment(a: string): RangeOps;
  setWrap(b: boolean): RangeOps;
  setWrapStrategy(s: number): RangeOps;
  setTextRotation(r: number): RangeOps;
  setNumberFormat(p: string): RangeOps;
  setBorder(type: unknown, style: unknown, color?: string): RangeOps;
  merge(): RangeOps;
  breakApart(): RangeOps;
  getValue(): unknown;
  setHyperLink(url: string, label?: string): Promise<boolean>;
}
interface PrintCell {
  v?: string | number | boolean | null;
  s?: string | { n?: { pattern?: string } };
}
interface PrintSnapshot {
  sheetOrder?: string[];
  sheets?: Record<string, { cellData?: Record<number, Record<number, PrintCell>> }>;
  styles?: Record<string, { n?: { pattern?: string } } | undefined>;
}
interface SheetOps {
  zoom(ratio: number): unknown;
  getSheetId(): string;
  getSheetName(): string;
  getFilter?(): unknown | null;
}
interface DefinedNameOps {
  getName(): string;
  getFormulaOrRefString(): string;
}
interface WorkbookOps {
  getActiveRange(): RangeOps | null;
  getActiveSheet(): SheetOps | null;
  getSheets(): SheetOps[];
  getDefinedNames(): DefinedNameOps[];
  getSnapshot(): PrintSnapshot;
}
interface ToolbarApi {
  undo(): unknown;
  redo(): unknown;
  executeCommand(id: string, params?: object): unknown;
  getActiveWorkbook(): WorkbookOps | null;
  addEvent?(event: string, cb: (params: unknown) => void): { dispose?: () => void } | undefined;
  Event?: Record<string, string>;
}

const SZ = 19;
// System/Office fonts first (built-in or metric substitutes), then the Google
// Fonts library — every one loads on demand via ../theme/google-fonts.
const FONT_FAMILIES = ["Arial", "Calibri", "Times New Roman", "Cambria", "Georgia", "Courier New", "Verdana", "Helvetica", ...GOOGLE_FONTS];

// Sticky "current font": once you pick a font, every cell you subsequently COMMIT
// adopts it (across sheets), until you pick a different one — matching the user's
// "pick once, type anywhere in that font" model. Module-level so it survives the
// per-sheet remount; resets on page reload (≈ opening a new document).
let stickyFontFamily: string | null = null;
const FONT_SIZES = [8, 9, 10, 11, 12, 14, 16, 18, 20, 24, 28, 36, 48];
const ZOOMS = [50, 75, 90, 100, 125, 150, 200];
// The Google Sheets "123" more-formats menu, grouped with samples.
const NUMBER_FORMAT_GROUPS: Array<Array<{ label: string; pattern: string; sample: string }>> = [
  [
    { label: "Automatic", pattern: "General", sample: "" },
    { label: "Plain text", pattern: "@", sample: "" },
  ],
  [
    { label: "Number", pattern: "#,##0.00", sample: "1,000.12" },
    { label: "Percent", pattern: "0.00%", sample: "10.12%" },
    { label: "Scientific", pattern: "0.00E+00", sample: "1.01E+03" },
  ],
  [
    { label: "Accounting", pattern: '_("$"* #,##0.00_);_("$"* (#,##0.00)', sample: "$ (1,000.12)" },
    { label: "Financial", pattern: "#,##0.00;(#,##0.00)", sample: "(1,000.12)" },
    { label: "Currency", pattern: '"$"#,##0.00', sample: "$1,000.12" },
    { label: "Currency rounded", pattern: '"$"#,##0', sample: "$1,000" },
  ],
  [
    { label: "Date", pattern: "yyyy-mm-dd", sample: "2026-09-26" },
    { label: "Time", pattern: "h:mm:ss am/pm", sample: "3:59:00 PM" },
    { label: "Date time", pattern: "yyyy-mm-dd h:mm:ss", sample: "2026-09-26 15:59:00" },
    { label: "Duration", pattern: "[h]:mm:ss", sample: "24:01:00" },
  ],
];
// The Google Sheets / Docs colour palette (10 cols × 8 rows) + standard row.
const GOOGLE_PALETTE: string[] = [
  "#000000", "#434343", "#666666", "#999999", "#b7b7b7", "#cccccc", "#d9d9d9", "#efefef", "#f3f3f3", "#ffffff",
  "#980000", "#ff0000", "#ff9900", "#ffff00", "#00ff00", "#00ffff", "#4a86e8", "#0000ff", "#9900ff", "#ff00ff",
  "#e6b8af", "#f4cccc", "#fce5cd", "#fff2cc", "#d9ead3", "#d0e0e3", "#c9daf8", "#cfe2f3", "#d9d2e9", "#ead1dc",
  "#dd7e6b", "#ea9999", "#f9cb9c", "#ffe599", "#b6d7a8", "#a2c4c9", "#a4c2f4", "#9fc5e8", "#b4a7d6", "#d5a6bd",
  "#cc4125", "#e06666", "#f6b26b", "#ffd966", "#93c47d", "#76a5af", "#6d9eeb", "#6fa8dc", "#8e7cc3", "#c27ba0",
  "#a61c00", "#cc0000", "#e69138", "#f1c232", "#6aa84f", "#45818e", "#3c78d8", "#3d85c6", "#674ea7", "#a64d79",
  "#85200c", "#990000", "#b45f06", "#bf9000", "#38761d", "#134f5c", "#1155cc", "#0b5394", "#351c75", "#741b47",
  "#5b0f00", "#660000", "#783f04", "#7f6000", "#274e13", "#0c343d", "#1c4587", "#073763", "#20124d", "#4c1130",
];
const STANDARD_COLORS = ["#000000", "#ffffff", "#4a86e8", "#e06666", "#ffd966", "#93c47d", "#f6b26b", "#76a5af"];
// Google-Sheets border positions. `type` is the literal Univer BorderType value;
// `seg` lists which cell-grid segments the icon lights up (t/b/l/r edges, mh/mv
// middle cross).
const BORDER_POSITIONS: Array<{ type: string; label: string; seg: string[] }> = [
  { type: "all", label: "All borders", seg: ["t", "b", "l", "r", "mh", "mv"] },
  { type: "inside", label: "Inner borders", seg: ["mh", "mv"] },
  { type: "horizontal", label: "Horizontal borders", seg: ["mh"] },
  { type: "vertical", label: "Vertical borders", seg: ["mv"] },
  { type: "outside", label: "Outer borders", seg: ["t", "b", "l", "r"] },
  { type: "left", label: "Left border", seg: ["l"] },
  { type: "top", label: "Top border", seg: ["t"] },
  { type: "right", label: "Right border", seg: ["r"] },
  { type: "bottom", label: "Bottom border", seg: ["b"] },
  { type: "none", label: "Clear borders", seg: [] },
];
// Border line styles → Univer BorderStyleTypes values.
const BORDER_STYLES: Array<{ v: number; label: string; w: number; dash: string; double?: boolean }> = [
  { v: 1, label: "Thin", w: 1, dash: "" },
  { v: 8, label: "Medium", w: 2, dash: "" },
  { v: 13, label: "Thick", w: 3, dash: "" },
  { v: 4, label: "Dashed", w: 1.4, dash: "4 2" },
  { v: 3, label: "Dotted", w: 1.4, dash: "1.5 2" },
  { v: 7, label: "Double", w: 1, dash: "", double: true },
];

/* ---- Styles --------------------------------------------------------------- */
const iconBtn: CSSProperties = {
  display: "inline-flex",
  alignItems: "center",
  justifyContent: "center",
  height: 31,
  minWidth: 28,
  padding: "0 5px",
  borderRadius: 8,
  border: "none",
  background: "transparent",
  color: "#475467",
  cursor: "pointer",
  transition: "background-color .12s ease, color .12s ease",
  flexShrink: 0,
};
const triggerBtn: CSSProperties = { ...iconBtn, gap: 2, padding: "0 5px" };
const divider: CSSProperties = { width: 1, height: 20, background: "#e4e7ec", margin: "0 2px", flexShrink: 0 };
const panelStyle: CSSProperties = {
  background: "#fff",
  border: "1px solid #eaecf0",
  borderRadius: 10,
  boxShadow: "0 8px 24px rgba(16,24,40,0.12)",
  padding: 6,
  zIndex: Z_BASE + 1000,
};
const menuItem: CSSProperties = {
  display: "flex",
  alignItems: "center",
  gap: 8,
  width: "100%",
  padding: "7px 10px",
  borderRadius: 6,
  border: "none",
  background: "transparent",
  color: "#344054",
  fontSize: 13,
  textAlign: "left",
  cursor: "pointer",
};
const tooltipStyle: CSSProperties = {
  position: "fixed",
  transform: "translateX(-50%)",
  background: "#0c111d",
  color: "#fff",
  fontSize: 12,
  fontWeight: 500,
  lineHeight: 1.3,
  padding: "4px 8px",
  borderRadius: 6,
  whiteSpace: "nowrap",
  pointerEvents: "none",
  zIndex: Z_BASE + 1100,
  boxShadow: "0 4px 12px rgba(16,24,40,0.18)",
};

/* ---- Tooltip wrapper (portaled so the scroll strip can't clip it) --------- */
function Tip({ label, children, disabled }: { label: string; children: ReactNode; disabled?: boolean }) {
  const ref = useRef<HTMLSpanElement>(null);
  const [pos, setPos] = useState<{ top: number; left: number } | null>(null);
  const show = () => {
    if (disabled) return; // e.g. while the dropdown is open
    const r = ref.current?.getBoundingClientRect();
    if (r) setPos({ top: r.bottom + 6, left: Math.max(72, Math.min(window.innerWidth - 72, r.left + r.width / 2)) });
  };
  return (
    <span ref={ref} style={{ position: "relative", display: "inline-flex" }} onMouseEnter={show} onMouseLeave={() => setPos(null)}>
      {children}
      {pos && createPortal(<span style={{ ...tooltipStyle, top: pos.top, left: pos.left }}>{label}</span>, document.body)}
    </span>
  );
}

/**
 * Keep the in-cell editor focused when a toolbar control is pressed. Without
 * this, mousedown moves focus off the editor → Univer commits the cell and
 * `EDITOR_ACTIVATED` turns false → the set-range commands fall back to whole-cell
 * styling instead of styling the in-cell rich text. preventDefault stops the
 * blur while still letting onClick run the command.
 */
const keepEditorFocus = (e: { preventDefault: () => void }) => e.preventDefault();

/* ---- Icon / text buttons -------------------------------------------------- */
function IconBtn({ label, onClick, active = false, children }: { label: string; onClick: () => void; active?: boolean; children: ReactNode }) {
  // When `active` (e.g. the selected cell is bold), the button stays filled —
  // mirroring Google Sheets' pressed B/I/U/S state — and hover only lightens it.
  const base: CSSProperties = active ? { ...iconBtn, background: "#fde68a", color: "#101828" } : iconBtn;
  return (
    <Tip label={label}>
      <button
        type="button"
        aria-label={label}
        aria-pressed={active}
        onMouseDown={keepEditorFocus}
        onClick={onClick}
        style={base}
        onMouseEnter={(e) => {
          e.currentTarget.style.background = "#fef3c7";
          e.currentTarget.style.color = "#101828";
        }}
        onMouseLeave={(e) => {
          e.currentTarget.style.background = active ? "#fde68a" : "transparent";
          e.currentTarget.style.color = active ? "#101828" : "#475467";
        }}
      >
        {children}
      </button>
    </Tip>
  );
}

/* ---- Icon with the current colour as a bar underneath (A / bucket) -------- */
function ColorIcon({ color, children }: { color: string; children: ReactNode }) {
  return (
    <span style={{ display: "inline-flex", flexDirection: "column", alignItems: "center", gap: 1, lineHeight: 0 }}>
      {children}
      <span style={{ width: SZ - 2, height: 3, borderRadius: 1, background: color, border: color.toLowerCase() === "#ffffff" ? "1px solid #d0d5dd" : "none", boxSizing: "border-box" }} />
    </span>
  );
}

/* ---- Portal dropdown (never clipped by the scroll strip) ------------------ */
function Dropdown({
  id,
  openId,
  setOpenId,
  label,
  trigger,
  width = 180,
  align = "left",
  children,
}: {
  id: string;
  openId: string | null;
  setOpenId: (v: string | null) => void;
  label: string;
  trigger: ReactNode;
  width?: number;
  /** "left" anchors the panel's left edge to the trigger; "right" anchors its
   *  right edge (opens leftward) — better for triggers near the screen edge. */
  align?: "left" | "right";
  children: (close: () => void) => ReactNode;
}) {
  const open = openId === id;
  const ref = useRef<HTMLButtonElement>(null);
  const [pos, setPos] = useState({ top: 0, left: 0 });
  const [hovered, setHovered] = useState(false);
  const toggle = () => {
    if (open) return setOpenId(null);
    const r = ref.current?.getBoundingClientRect();
    if (r) {
      // Clamp within the viewport so panels near the right edge aren't clipped.
      // Right-aligned panels open leftward and keep a gap from the screen edge.
      const left =
        align === "right"
          ? Math.max(8, Math.min(r.right, window.innerWidth - 16) - width)
          : Math.max(8, Math.min(r.left, window.innerWidth - width - 8));
      setPos({ top: r.bottom + 4, left });
    }
    setOpenId(id);
  };
  // Warm yellow fill when open (selected); lighter amber on hover.
  const triggerStyle: CSSProperties = {
    ...triggerBtn,
    boxSizing: "border-box",
    border: "1px solid transparent",
    background: open ? "#fde68a" : hovered ? "#fef3c7" : "transparent",
    color: open || hovered ? "#101828" : "#475467",
    boxShadow: "none",
  };
  return (
    <>
      <Tip label={label} disabled={open}>
        <button type="button" ref={ref} aria-label={label} onMouseDown={keepEditorFocus} onClick={toggle} style={triggerStyle} onMouseEnter={() => setHovered(true)} onMouseLeave={() => setHovered(false)}>
          {trigger}
          <ChevronDown size={14} />
        </button>
      </Tip>
      {open &&
        createPortal(
          <div data-levich-dd style={{ position: "fixed", top: pos.top, left: pos.left, width, ...panelStyle }}>
            {children(() => setOpenId(null))}
          </div>,
          document.body,
        )}
    </>
  );
}

function MenuRow({ onClick, children }: { onClick: () => void; children: ReactNode }) {
  return (
    <button type="button" onMouseDown={keepEditorFocus} onClick={onClick} style={menuItem} onMouseEnter={(e) => (e.currentTarget.style.background = "#f9fafb")} onMouseLeave={(e) => (e.currentTarget.style.background = "transparent")}>
      {children}
    </button>
  );
}

/** Font dropdown body — search bar + ~10 visible rows (scroll for more). Batch-
 *  loads the Google Fonts list on open so each row previews in its own typeface. */
function FontMenu({ onPick }: { onPick: (f: string) => void }) {
  const [q, setQ] = useState("");
  useEffect(() => { void ensureGoogleFonts(GOOGLE_FONTS); }, []);
  const filtered = FONT_FAMILIES.filter((f) => f.toLowerCase().includes(q.trim().toLowerCase()));
  return (
    <div>
      <div style={{ display: "flex", alignItems: "center", gap: 8, padding: "6px 10px", margin: "2px 4px 6px", border: "1px solid #e4e7ec", borderRadius: 8, color: "#667085" }}>
        <SearchMd size={15} />
        <input
          autoFocus value={q} onChange={(e) => setQ(e.target.value)} placeholder="Search fonts"
          style={{ border: "none", outline: "none", flex: 1, fontSize: 13, color: "#101828", background: "transparent", fontFamily: "inherit" }}
        />
      </div>
      {/* ~10 rows tall, then scroll */}
      <div style={{ maxHeight: 10 * 34, overflowY: "auto" }}>
        {filtered.map((f) => (
          <MenuRow key={f} onClick={() => onPick(f)}>
            <span style={{ fontFamily: `"${f}", sans-serif` }}>{f}</span>
          </MenuRow>
        ))}
        {filtered.length === 0 && <div style={{ padding: "10px 12px", color: "#98a2b3", fontSize: 13 }}>No fonts found</div>}
      </div>
    </div>
  );
}

/* ---- Color panel — the full Google Sheets palette + custom ---------------- */
function Swatch({ c, selected, onPick }: { c: string; selected?: boolean; onPick: (c: string) => void }) {
  const isWhite = c.toLowerCase() === "#ffffff";
  return (
    <button
      type="button"
      aria-label={c}
      aria-pressed={selected}
      title={c}
      onMouseDown={keepEditorFocus}
      onClick={() => onPick(c)}
      style={{
        width: 16,
        height: 16,
        borderRadius: 3,
        border: selected ? "2px solid #101828" : isWhite ? "1px solid #d0d5dd" : "1px solid rgba(0,0,0,0.12)",
        background: c,
        cursor: "pointer",
        padding: 0,
        transition: "transform .1s ease, box-shadow .1s ease",
      }}
      onMouseEnter={(e) => {
        e.currentTarget.style.transform = "scale(1.25)";
        e.currentTarget.style.boxShadow = "0 0 0 2px #fde68a";
      }}
      onMouseLeave={(e) => {
        e.currentTarget.style.transform = "none";
        e.currentTarget.style.boxShadow = "none";
      }}
    />
  );
}

const sectionLabelStyle: CSSProperties = { fontSize: 11, fontWeight: 700, letterSpacing: 0.3, color: "#98a2b3", margin: "10px 2px 6px" };

/* ---- Colour maths (for the in-panel custom picker) ------------------------ */
const clamp01 = (n: number) => Math.max(0, Math.min(1, n));
function hexToRgb(hex: string): [number, number, number] {
  const m = /^#?([0-9a-f]{6})$/i.exec(hex.trim());
  if (!m) return [16, 24, 40];
  const n = parseInt(m[1], 16);
  return [(n >> 16) & 255, (n >> 8) & 255, n & 255];
}
const rgbToHex = (r: number, g: number, b: number) => "#" + [r, g, b].map((x) => Math.round(x).toString(16).padStart(2, "0")).join("");
function rgbToHsv(r: number, g: number, b: number): [number, number, number] {
  r /= 255; g /= 255; b /= 255;
  const max = Math.max(r, g, b), min = Math.min(r, g, b), d = max - min;
  let h = 0;
  if (d) {
    if (max === r) h = ((g - b) / d) % 6;
    else if (max === g) h = (b - r) / d + 2;
    else h = (r - g) / d + 4;
    h *= 60;
    if (h < 0) h += 360;
  }
  return [h, max ? d / max : 0, max];
}
function hsvToRgb(h: number, s: number, v: number): [number, number, number] {
  const c = v * s, x = c * (1 - Math.abs(((h / 60) % 2) - 1)), m = v - c;
  const [r, g, b] = h < 60 ? [c, x, 0] : h < 120 ? [x, c, 0] : h < 180 ? [0, c, x] : h < 240 ? [0, x, c] : h < 300 ? [x, 0, c] : [c, 0, x];
  return [(r + m) * 255, (g + m) * 255, (b + m) * 255];
}

/**
 * In-panel custom colour picker (saturation/value square + hue slider + hex).
 * Everything is hoverable and lives inside our panel — no OS colour dialog.
 * `onChange` is the live, applied value.
 */
function CustomColorPicker({ value, onChange }: { value: string; onChange: (hex: string) => void }) {
  const initial = rgbToHsv(...hexToRgb(value));
  const [hsv, setHsv] = useState<[number, number, number]>(initial);
  const [hex, setHex] = useState(value);
  const boxRef = useRef<HTMLDivElement>(null);
  const hueRef = useRef<HTMLDivElement>(null);
  const hsvRef = useRef(hsv);
  hsvRef.current = hsv;

  const commit = (h: number, s: number, v: number) => {
    const hx = rgbToHex(...hsvToRgb(h, s, v));
    setHex(hx);
    onChange(hx);
  };
  const drag = (ref: typeof boxRef, e: PointerEvent | { clientX: number; clientY: number }, hue: boolean) => {
    const r = ref.current?.getBoundingClientRect();
    if (!r) return;
    const [h] = hsvRef.current;
    if (hue) {
      const nh = clamp01((e.clientX - r.left) / r.width) * 360;
      const next: [number, number, number] = [nh, hsvRef.current[1], hsvRef.current[2]];
      setHsv(next);
      commit(...next);
    } else {
      const s = clamp01((e.clientX - r.left) / r.width);
      const v = 1 - clamp01((e.clientY - r.top) / r.height);
      const next: [number, number, number] = [h, s, v];
      setHsv(next);
      commit(...next);
    }
  };
  const start = (ref: typeof boxRef, hue: boolean) => (e: { clientX: number; clientY: number; preventDefault: () => void }) => {
    e.preventDefault();
    drag(ref, e, hue);
    const mv = (ev: PointerEvent) => drag(ref, ev, hue);
    const up = () => {
      window.removeEventListener("pointermove", mv);
      window.removeEventListener("pointerup", up);
    };
    window.addEventListener("pointermove", mv);
    window.addEventListener("pointerup", up);
  };

  const [h, s, v] = hsv;
  const hueColor = rgbToHex(...hsvToRgb(h, 1, 1));
  return (
    <div style={{ marginTop: 6 }} onMouseDown={keepEditorFocus}>
      <div
        ref={boxRef}
        onPointerDown={start(boxRef, false)}
        style={{ position: "relative", height: 116, borderRadius: 8, cursor: "crosshair", background: `linear-gradient(to top, #000, rgba(0,0,0,0)), linear-gradient(to right, #fff, rgba(255,255,255,0)), ${hueColor}` }}
      >
        <span style={{ position: "absolute", left: `${s * 100}%`, top: `${(1 - v) * 100}%`, width: 12, height: 12, marginLeft: -6, marginTop: -6, borderRadius: "50%", border: "2px solid #fff", boxShadow: "0 0 0 1px rgba(0,0,0,.4)", background: hex, pointerEvents: "none" }} />
      </div>
      <div
        ref={hueRef}
        onPointerDown={start(hueRef, true)}
        style={{ position: "relative", height: 12, borderRadius: 6, marginTop: 10, cursor: "ew-resize", background: "linear-gradient(to right, #f00, #ff0, #0f0, #0ff, #00f, #f0f, #f00)" }}
      >
        <span style={{ position: "absolute", left: `${(h / 360) * 100}%`, top: "50%", width: 14, height: 14, marginLeft: -7, marginTop: -7, borderRadius: "50%", border: "2px solid #fff", boxShadow: "0 0 0 1px rgba(0,0,0,.4)", background: hueColor, pointerEvents: "none" }} />
      </div>
      <div style={{ display: "flex", alignItems: "center", gap: 8, marginTop: 10 }}>
        <span style={{ width: 28, height: 28, borderRadius: 6, border: "1px solid #d0d5dd", background: hex, flexShrink: 0 }} />
        <input
          value={hex}
          onMouseDown={keepEditorFocus}
          onChange={(e) => {
            const t = e.target.value;
            setHex(t);
            if (/^#?[0-9a-f]{6}$/i.test(t)) {
              const norm = t.startsWith("#") ? t : "#" + t;
              setHsv(rgbToHsv(...hexToRgb(norm)));
              onChange(norm);
            }
          }}
          style={{ flex: 1, height: 30, borderRadius: 8, border: "1px solid #d0d5dd", padding: "0 10px", fontSize: 13, color: "#101828", textTransform: "uppercase", fontVariantNumeric: "tabular-nums", outline: "none" }}
        />
      </div>
    </div>
  );
}

function ColorPanel({ current, onPick, onApply }: { current?: string; onPick: (c: string) => void; onApply: (c: string) => void }) {
  // Swatches apply immediately and close (Google-Sheets style); "Custom…" opens
  // the in-panel hoverable picker which applies live but keeps the panel open.
  const [showCustom, setShowCustom] = useState(false);
  const isSel = (c: string) => !!current && current.toLowerCase() === c.toLowerCase();
  return (
    <div style={{ width: 196 }}>
      <div style={{ display: "grid", gridTemplateColumns: "repeat(10, 1fr)", gap: 3 }}>
        {GOOGLE_PALETTE.map((c, i) => (
          <Swatch key={i} c={c} selected={isSel(c)} onPick={onPick} />
        ))}
      </div>
      <div style={sectionLabelStyle}>STANDARD</div>
      <div style={{ display: "grid", gridTemplateColumns: "repeat(10, 1fr)", gap: 3 }}>
        {STANDARD_COLORS.map((c, i) => (
          <Swatch key={i} c={c} selected={isSel(c)} onPick={onPick} />
        ))}
      </div>
      <div style={sectionLabelStyle}>CUSTOM</div>
      <button
        type="button"
        onMouseDown={keepEditorFocus}
        onClick={() => setShowCustom((x) => !x)}
        style={{ display: "flex", alignItems: "center", gap: 8, width: "100%", border: "none", background: "transparent", cursor: "pointer", fontSize: 13, color: "#475467", padding: "2px 2px" }}
      >
        <span style={{ width: 18, height: 18, borderRadius: 4, border: "1px solid #d0d5dd", display: "inline-flex", alignItems: "center", justifyContent: "center", fontSize: 14, lineHeight: 1, color: "#667085" }}>{showCustom ? "–" : "+"}</span>
        Custom…
      </button>
      {showCustom && <CustomColorPicker value={current && /^#/.test(current) ? current : "#101828"} onChange={onApply} />}
    </div>
  );
}

/* ---- Borders panel (Google-Sheets style: positions + colour + style) ------ */
const BORDER_SEGMENTS: Record<string, [number, number, number, number]> = {
  t: [3.5, 3.5, 14.5, 3.5],
  b: [3.5, 14.5, 14.5, 14.5],
  l: [3.5, 3.5, 3.5, 14.5],
  r: [14.5, 3.5, 14.5, 14.5],
  mh: [3.5, 9, 14.5, 9],
  mv: [9, 3.5, 9, 14.5],
};
function BorderIcon({ active }: { active: string[] }) {
  return (
    <svg width="18" height="18" viewBox="0 0 18 18" aria-hidden>
      {Object.entries(BORDER_SEGMENTS).map(([k, [x1, y1, x2, y2]]) => {
        const on = active.includes(k);
        const inner = k === "mh" || k === "mv";
        return <line key={k} x1={x1} y1={y1} x2={x2} y2={y2} stroke={on ? "#344054" : "#e4e7ec"} strokeWidth={on ? 1.6 : 1} strokeLinecap="square" strokeDasharray={inner ? "1.6 1.6" : undefined} />;
      })}
    </svg>
  );
}
/** A horizontal preview line for a given border style. */
function LinePreview({ style, color = "#344054", width = 40 }: { style: { w: number; dash: string; double?: boolean }; color?: string; width?: number }) {
  if (style.double) {
    return (
      <svg width={width} height="10" viewBox={`0 0 ${width} 10`} aria-hidden>
        <line x1="0" y1="3.5" x2={width} y2="3.5" stroke={color} strokeWidth="1" />
        <line x1="0" y1="6.5" x2={width} y2="6.5" stroke={color} strokeWidth="1" />
      </svg>
    );
  }
  return (
    <svg width={width} height="10" viewBox={`0 0 ${width} 10`} aria-hidden>
      <line x1="0" y1="5" x2={width} y2="5" stroke={color} strokeWidth={style.w} strokeDasharray={style.dash || undefined} />
    </svg>
  );
}

function BordersPanel({ apply, onPick }: { apply: (type: string, style: number, color: string) => void; onPick: () => void }) {
  const [color, setColor] = useState("#000000");
  const [styleV, setStyleV] = useState(1);
  const [sub, setSub] = useState<null | "color" | "style">(null);
  const style = BORDER_STYLES.find((s) => s.v === styleV) ?? BORDER_STYLES[0];
  const cellBtn: CSSProperties = { display: "inline-flex", alignItems: "center", justifyContent: "center", width: "100%", height: 30, border: "1px solid transparent", borderRadius: 6, background: "transparent", cursor: "pointer" };
  const ctrlBtn: CSSProperties = { display: "inline-flex", alignItems: "center", gap: 6, flex: 1, height: 32, padding: "0 8px", border: "1px solid #eaecf0", borderRadius: 8, background: "#fff", cursor: "pointer" };
  return (
    <div style={{ width: 220 }}>
      <div style={{ display: "grid", gridTemplateColumns: "repeat(5, 1fr)", gap: 2 }}>
        {BORDER_POSITIONS.map((p) => (
          <Tip key={p.type} label={p.label}>
            <button
              type="button"
              aria-label={p.label}
              onMouseDown={keepEditorFocus}
              onClick={() => apply(p.type, p.type === "none" ? 0 : styleV, color)}
              style={cellBtn}
              onMouseEnter={(e) => (e.currentTarget.style.background = "#fef3c7")}
              onMouseLeave={(e) => (e.currentTarget.style.background = "transparent")}
            >
              <BorderIcon active={p.seg} />
            </button>
          </Tip>
        ))}
      </div>
      <div style={{ height: 1, background: "#eaecf0", margin: "8px 0" }} />
      <div style={{ display: "flex", gap: 6 }}>
        <Tip label="Border color">
          <button type="button" onMouseDown={keepEditorFocus} onClick={() => setSub((s) => (s === "color" ? null : "color"))} style={ctrlBtn}>
            <span style={{ width: 18, height: 18, borderRadius: 4, border: color.toLowerCase() === "#ffffff" ? "1px solid #d0d5dd" : "1px solid rgba(0,0,0,.15)", background: color }} />
            <ChevronDown size={14} color="#667085" />
          </button>
        </Tip>
        <Tip label="Border style">
          <button type="button" onMouseDown={keepEditorFocus} onClick={() => setSub((s) => (s === "style" ? null : "style"))} style={ctrlBtn}>
            <LinePreview style={style} width={34} />
            <ChevronDown size={14} color="#667085" />
          </button>
        </Tip>
      </div>
      {sub === "color" && (
        <div style={{ marginTop: 8 }}>
          <ColorPanel current={color} onApply={setColor} onPick={(c) => { setColor(c); setSub(null); }} />
        </div>
      )}
      {sub === "style" && (
        <div style={{ marginTop: 8 }}>
          {BORDER_STYLES.map((s) => (
            <MenuRow key={s.v} onClick={() => { setStyleV(s.v); setSub(null); }}>
              <LinePreview style={s} width={90} />
              <span style={{ marginLeft: "auto", color: "#98a2b3", fontSize: 12 }}>{s.label}</span>
              {s.v === styleV && <span style={{ marginLeft: 8, color: "#101828" }}>✓</span>}
            </MenuRow>
          ))}
        </div>
      )}
      <div style={{ height: 1, background: "#eaecf0", margin: "8px 0 4px" }} />
      <MenuRow onClick={onPick}><span style={{ color: "#667085", fontSize: 12 }}>Close</span></MenuRow>
    </div>
  );
}

/* ---- Vertical-align / wrap / rotation glyphs ------------------------------ */
function VAlignGlyph({ pos }: { pos: "top" | "middle" | "bottom" }) {
  const y = pos === "top" ? 3.5 : pos === "middle" ? 7 : 10.5;
  return (
    <svg width="16" height="16" viewBox="0 0 16 16" fill="none" aria-hidden>
      <rect x="2.5" y="2.5" width="11" height="11" rx="1.5" stroke="#d0d5dd" />
      <rect x="4.5" y={y} width="7" height="2.4" rx="1.2" fill="currentColor" />
    </svg>
  );
}
const AlignTopGlyph = () => <VAlignGlyph pos="top" />;
const AlignMiddleGlyph = () => <VAlignGlyph pos="middle" />;

function WrapGlyph({ mode }: { mode: "overflow" | "wrap" | "clip" }) {
  return (
    <svg width="18" height="18" viewBox="0 0 18 18" fill="none" aria-hidden stroke="currentColor" strokeWidth="1.4" strokeLinecap="round" strokeLinejoin="round">
      <rect x="2.5" y="3" width="13" height="12" rx="1.5" stroke="#d0d5dd" />
      {mode === "overflow" && <line x1="5" y1="9" x2="17.5" y2="9" />}
      {mode === "clip" && <line x1="5" y1="9" x2="12.5" y2="9" />}
      {mode === "wrap" && (
        <>
          <path d="M5 6.5H11.5a2 2 0 0 1 0 4H7" />
          <path d="M8.5 9l-1.8 1.5L8.5 12" fill="none" />
        </>
      )}
    </svg>
  );
}

/** An "A" rotated to preview a text-rotation angle (deg = counter-clockwise). */
function RotateGlyph({ deg }: { deg: number }) {
  return (
    <span style={{ display: "inline-flex", width: 18, height: 18, alignItems: "center", justifyContent: "center", fontSize: 13, fontWeight: 700, lineHeight: 1, transform: `rotate(${-deg}deg)` }}>A</span>
  );
}
function VerticalStackGlyph() {
  return (
    <span style={{ display: "inline-flex", flexDirection: "column", width: 18, height: 18, alignItems: "center", justifyContent: "center", fontSize: 8, fontWeight: 700, lineHeight: 0.95, letterSpacing: 0 }}>
      <span>A</span>
      <span>B</span>
    </span>
  );
}

/* ---- Text rotation panel (Google-Sheets style) ---------------------------- */
const ROTATIONS: Array<{ label: string; v: number | string; deg?: number; vertical?: boolean }> = [
  { label: "None", v: 0, deg: 0 },
  { label: "Tilt up", v: 45, deg: 45 },
  { label: "Tilt down", v: -45, deg: -45 },
  { label: "Stack vertically", v: "vertical", vertical: true },
  { label: "Rotate up", v: 90, deg: 90 },
  { label: "Rotate down", v: -90, deg: -90 },
];
const ROTATION_ANGLES = [-90, -75, -60, -45, -30, -15, 0, 15, 30, 45, 60, 75, 90];
function RotationPanel({ current, apply, onPick }: { current: number | string; apply: (v: number | string) => void; onPick: () => void }) {
  const [showAngles, setShowAngles] = useState(false);
  const curAngle = typeof current === "number" ? current : 0;
  const isSel = (o: (typeof ROTATIONS)[number]) => (o.vertical ? current === "vertical" : current === o.v);
  return (
    <div style={{ width: 196 }}>
      {ROTATIONS.map((o) => (
        <MenuRow key={o.label} onClick={() => { apply(o.v); onPick(); }}>
          {o.vertical ? <VerticalStackGlyph /> : <RotateGlyph deg={o.deg ?? 0} />}
          <span style={{ marginLeft: 2 }}>{o.label}</span>
          {isSel(o) && <span style={{ marginLeft: "auto", color: "#101828" }}>✓</span>}
        </MenuRow>
      ))}
      <div style={{ height: 1, background: "#eaecf0", margin: "6px 4px" }} />
      {/* Discrete angle picker (−90°…90°), using our own dropdown style. */}
      <button
        type="button"
        onMouseDown={keepEditorFocus}
        onClick={() => setShowAngles((x) => !x)}
        style={{ display: "flex", alignItems: "center", gap: 8, width: "100%", height: 32, padding: "0 8px", border: "1px solid #eaecf0", borderRadius: 8, background: "#fff", cursor: "pointer", color: "#344054" }}
      >
        <RotateGlyph deg={current === "vertical" ? 0 : curAngle} />
        <span style={{ fontSize: 13 }}>{current === "vertical" ? "Custom angle" : `${curAngle}°`}</span>
        <ChevronDown size={14} color="#667085" style={{ marginLeft: "auto" }} />
      </button>
      {showAngles && (
        <div style={{ marginTop: 4, maxHeight: 200, overflowY: "auto" }}>
          {ROTATION_ANGLES.map((a) => (
            <MenuRow key={a} onClick={() => { apply(a); onPick(); }}>
              <RotateGlyph deg={a} />
              <span style={{ marginLeft: 2 }}>{a}°</span>
              {current !== "vertical" && a === curAngle && <span style={{ marginLeft: "auto", color: "#101828" }}>✓</span>}
            </MenuRow>
          ))}
        </div>
      )}
    </div>
  );
}

/* ---- Insert-link panel (Google-Sheets style: URL · Range · Named range) ---- */
type LinkType = "url" | "range" | "name";
const LINK_TABS: Array<{ key: LinkType; label: string }> = [
  { key: "url", label: "Link" },
  { key: "range", label: "Range" },
  { key: "name", label: "Named range" },
];
function LinkPanel({
  initialText,
  definedNames,
  onApply,
  onCancel,
}: {
  initialText: string;
  definedNames: Array<{ name: string; ref: string }>;
  onApply: (type: LinkType, label: string, value: string) => void;
  onCancel: () => void;
}) {
  const [text, setText] = useState(initialText);
  const [type, setType] = useState<LinkType>("url");
  const [url, setUrl] = useState("");
  const [range, setRange] = useState("");
  const [defIdx, setDefIdx] = useState(0);
  const inputStyle: CSSProperties = { width: "100%", height: 36, borderRadius: 8, border: "1px solid #d0d5dd", padding: "0 10px", fontSize: 13, outline: "none", boxSizing: "border-box", color: "#101828" };
  const sel = definedNames[defIdx];

  const value = type === "url" ? url.trim() : type === "range" ? range.trim() : sel?.ref ?? "";
  const label = text.trim() || (type === "name" ? sel?.name ?? "" : value);
  const valid = type === "name" ? !!sel : value.length > 0;
  const submit = () => { if (valid) onApply(type, label, value); };

  return (
    <div style={{ width: 300, padding: 4 }}>
      {/* Type tabs */}
      <div style={{ display: "flex", gap: 4, background: "#f2f4f7", borderRadius: 8, padding: 3, marginBottom: 10 }}>
        {LINK_TABS.map((t) => (
          <button
            key={t.key}
            type="button"
            onClick={() => setType(t.key)}
            style={{ flex: 1, height: 28, borderRadius: 6, border: "none", cursor: "pointer", fontSize: 12.5, fontWeight: 600, background: type === t.key ? "#fff" : "transparent", color: type === t.key ? "#101828" : "#667085", boxShadow: type === t.key ? "0 1px 2px rgba(16,24,40,0.08)" : "none" }}
          >
            {t.label}
          </button>
        ))}
      </div>
      {/* Display text */}
      <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 8 }}>
        <Type01 size={16} color="#667085" />
        <input value={text} placeholder="Text to display" onChange={(e) => setText(e.target.value)} style={inputStyle} />
      </div>
      {/* Per-type input */}
      {type === "url" && (
        <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
          <Link01 size={16} color="#667085" />
          <input value={url} placeholder="Paste or type a link" autoFocus onChange={(e) => setUrl(e.target.value)} onKeyDown={(e) => e.key === "Enter" && submit()} style={inputStyle} />
        </div>
      )}
      {type === "range" && (
        <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
          <Grid01 size={16} color="#667085" />
          <input value={range} placeholder="Range in this sheet, e.g. A1:B10" autoFocus onChange={(e) => setRange(e.target.value)} onKeyDown={(e) => e.key === "Enter" && submit()} style={inputStyle} />
        </div>
      )}
      {type === "name" && (
        definedNames.length ? (
          <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
            <Grid01 size={16} color="#667085" />
            <select value={defIdx} onChange={(e) => setDefIdx(Number(e.target.value))} style={{ ...inputStyle, cursor: "pointer", background: "#fff" }}>
              {definedNames.map((d, i) => (
                <option key={d.name} value={i}>{d.name} — {d.ref}</option>
              ))}
            </select>
          </div>
        ) : (
          <div style={{ fontSize: 12.5, color: "#98a2b3", padding: "6px 2px" }}>No named ranges yet. Create one via the Name Box (top-left).</div>
        )
      )}
      <div style={{ display: "flex", gap: 8, marginTop: 12 }}>
        <button type="button" onClick={onCancel} style={{ flex: 1, height: 34, borderRadius: 8, border: "1px solid #d0d5dd", background: "#fff", color: "#344054", fontSize: 13, fontWeight: 600, cursor: "pointer" }}>
          Cancel
        </button>
        <button type="button" onClick={submit} disabled={!valid} style={{ flex: 1, height: 34, borderRadius: 8, border: "none", background: valid ? "#0a0a0a" : "#d0d5dd", color: "#fff", fontSize: 13, fontWeight: 600, cursor: valid ? "pointer" : "default" }}>
          Apply
        </button>
      </div>
    </div>
  );
}

export interface LevichToolbarProps {
  api: UniverAPI | null;
  /** Open the Levich Find & Replace modal (replaces Univer's native panel). */
  onOpenFind?: () => void;
}

export function LevichToolbar({ api, onOpenFind }: LevichToolbarProps) {
  const stripRef = useRef<HTMLDivElement | null>(null);
  const [openId, setOpenId] = useState<string | null>(null);
  const [zoomPct, setZoomPct] = useState(100);
  const [fontPt, setFontPt] = useState(11);
  const [fontFamily, setFontFamily] = useState("Arial");
  // Active formatting of the currently-selected cell (reflected in the toolbar).
  const [active, setActive] = useState({ bold: false, italic: false, underline: false, strike: false });
  // Last-used text / fill colour (shown as the bar under the A / bucket icons).
  const [textColor, setTextColor] = useState("#101828");
  const [fillColor, setFillColor] = useState("#fde68a");
  // Current text rotation (degrees, or "vertical" for stacked) — drives the
  // toolbar trigger icon, like Google Sheets.
  const [rotation, setRotation] = useState<number | "vertical">(0);
  // Whether the sheet currently has a filter (lights the Filter button).
  const [filterOn, setFilterOn] = useState(false);
  // Saved Filter Views (Google-style) + the active one. State only — non
  // destructive snapshots of the per-column filter criteria.
  const [views, setViews] = useState<FilterView[]>([]);
  const [groupViews, setGroupViews] = useState<GroupView[]>([]);
  const [activeView, setActiveView] = useState<string | null>(null);
  const [editingViewId, setEditingViewId] = useState<string | null>(null);
  // Active group-by view: the banner label + the live inserted-row run.
  const [groupBanner, setGroupBanner] = useState<string | null>(null);
  const [viewSaved, setViewSaved] = useState(false);
  const groupRunRef = useRef<GroupRun | null>(null);
  const lastToggleRef = useRef<{ t: number; row: number }>({ t: 0, row: -1 });
  const viewSeq = useRef(0);
  const viewsLoaded = useRef(false);
  // Last sort applied (full sort-range params) — captured so a saved Filter
  // View can reproduce the sorted order, like Google's sortSpecs.
  const lastSortRef = useRef<object | null>(null);
  // True while a cell is in edit mode (tracked via Univer's edit events).
  const editingRef = useRef(false);

  const apiOf = () => api as unknown as ToolbarApi | null;
  const fvApi = () => apiOf() as unknown as FilterViewApi | null;
  const gbApi = () => apiOf() as unknown as GroupByApi | null;
  const fnApi = () => apiOf() as unknown as FunctionsApi | null;

  // ---- Views (filter + group) handlers --------------------------------------
  // Remove an active group view's inserted rows (if any) before doing anything
  // else — only one view is active at a time.
  const clearGroupRun = () => {
    if (groupRunRef.current) {
      try {
        clearGroupView(gbApi(), groupRunRef.current);
      } catch {
        /* best effort */
      }
      groupRunRef.current = null;
    }
    setGroupBanner(null);
  };
  const createFilterView = () => {
    clearGroupRun();
    ensureFilter(fvApi());
    const cols = snapshotFilters(fvApi());
    const id = `view_${++viewSeq.current}`;
    setViews((v) => [...v, { id, name: `Filter view ${v.length + 1}`, cols, sort: lastSortRef.current }]);
    setActiveView(id);
    setViewSaved(false);
    setFilterOn(true);
  };
  const applyView = (v: FilterView) => {
    clearGroupRun();
    applyFilterView(fvApi(), v);
    setActiveView(v.id);
    setViewSaved(true);
    setFilterOn(true);
  };
  const startGroupBy = (colIndex: number, label: string) => {
    clearGroupRun();
    const run = buildGroupView(gbApi(), colIndex);
    groupRunRef.current = run;
    setGroupBanner(run && run.groups.length ? label : null);
    setFilterOn(true);
  };
  const createGroupView = (colIndex: number, label: string) => {
    startGroupBy(colIndex, label);
    const id = `group_${++viewSeq.current}`;
    setGroupViews((g) => [...g, { id, name: `Group by ${label}`, colIndex }]);
    setActiveView(id);
    setViewSaved(false); // freshly created → "Temporary" until Save view
  };
  const applyGroupView = (v: GroupView) => {
    const label = v.name.replace(/^Group by\s*/, "") || "column";
    startGroupBy(v.colIndex, label);
    setActiveView(v.id);
    setViewSaved(true); // applying an already-saved view
  };
  const exitFilterView = () => {
    clearGroupRun();
    clearAllFilters(fvApi());
    setActiveView(null);
  };
  const refreshView = () => {
    if (!activeView) return;
    const fv = views.find((v) => v.id === activeView);
    if (fv) {
      clearGroupRun();
      applyFilterView(fvApi(), fv);
      return;
    }
    const gv = groupViews.find((v) => v.id === activeView);
    if (gv) applyGroupView(gv);
  };
  const saveView = () => {
    if (!activeView) return;
    // Filter view → re-snapshot current filters + sort into it. Group view →
    // already persisted on creation; Save just confirms it (drops "Temporary").
    setViews((vs) => vs.map((v) => (v.id === activeView ? { ...v, cols: snapshotFilters(fvApi()), sort: lastSortRef.current } : v)));
    setViewSaved(true);
  };
  const deleteView = (id: string) => {
    setViews((vs) => vs.filter((x) => x.id !== id));
    setGroupViews((gs) => gs.filter((x) => x.id !== id));
    if (activeView === id) exitFilterView();
  };
  const renameView = (id: string, name: string) => {
    const clean = name.trim();
    setViews((vs) => vs.map((x) => (x.id === id ? { ...x, name: clean || x.name } : x)));
    setGroupViews((gs) => gs.map((x) => (x.id === id ? { ...x, name: clean || x.name } : x)));
    setEditingViewId(null);
  };
  const withRange = (fn: (r: RangeOps) => void) => {
    const r = apiOf()?.getActiveWorkbook()?.getActiveRange();
    if (r) fn(r);
  };
  // B / I / U / S behave exactly like Univer's own toolbar. While a cell is being
  // edited we dispatch the *docs* inline-format command directly (styling the
  // in-cell rich text: the selected characters, or the caret's stored format for
  // everything typed next). Otherwise we dispatch the *range* command (whole
  // cell). Univer tracks underline & strikethrough independently. We flip our own
  // lit state optimistically; the selection listener resyncs it when you move.
  const fireMark = (rangeCmd: string, inlineCmd: string, key: "bold" | "italic" | "underline" | "strike") => {
    const api = apiOf();
    if (editingRef.current) {
      // Editing: style the WHOLE in-cell text (select-all → inline format → collapse the caret) so the
      // characters already typed change immediately — Univer's inline command alone only affects the
      // current selection/caret, so B/I/U/S looked like it "did nothing" while a cell was being edited.
      try {
        api?.executeCommand("doc.command.select-all");
        api?.executeCommand(inlineCmd);
        api?.executeCommand("doc.operation.move-cursor", { direction: "right" });
      } catch {
        /* in-cell editor selection is best-effort */
      }
    }
    // Always set the CELL (range) format too, so the committed cell keeps the style and text typed
    // next inherits it (also the correct path when not editing).
    api?.executeCommand(rangeCmd);
    setActive((a) => ({ ...a, [key]: !a[key] }));
  };
  const toggleBold = () => fireMark("sheet.command.set-range-bold", "doc.command.set-inline-format-bold", "bold");
  const toggleItalic = () => fireMark("sheet.command.set-range-italic", "doc.command.set-inline-format-italic", "italic");
  const toggleUnderline = () => fireMark("sheet.command.set-range-underline", "doc.command.set-inline-format-underline", "underline");
  const toggleStrike = () => fireMark("sheet.command.set-range-stroke", "doc.command.set-inline-format-strikethrough", "strike");
  // Font size / family / text colour: while a cell is being edited we dispatch
  // the *docs* inline-format command directly (styling the selected characters /
  // the caret's stored format for what's typed next); otherwise the *range*
  // command (whole cell). Same edit-state-aware approach as B/I/U/S — the
  // range command's own EDITOR_ACTIVATED check is unreliable from an external
  // toolbar, which is why setting e.g. green didn't colour the typed text.
  const fireStyled = (rangeCmd: string, inlineCmd: string, value: string | number) =>
    apiOf()?.executeCommand(editingRef.current ? inlineCmd : rangeCmd, { value });
  const applyFontSize = (s: number) => {
    setFontPt(s);
    fireStyled("sheet.command.set-range-fontsize", "doc.command.set-inline-format-fontsize", s);
  };
  const changeFont = (d: number) => applyFontSize(Math.max(6, Math.min(96, fontPt + d)));
  const applyFontFamily = (f: string) => {
    setFontFamily(f); // update the toolbar indicator immediately
    stickyFontFamily = f; // make it the sticky current font for future typing
    // Apply AFTER the web font is loaded — Univer measures & caches text on the
    // canvas at paint time, so setting `ff` before the font is ready bakes in a
    // fallback that never updates. The dropdown batch-loads on open, so by pick
    // time the font is usually already cached (applies instantly).
    const apply = () => {
      const api = apiOf();
      if (editingRef.current) {
        // Editing (Excel/Sheets behaviour): style the WHOLE cell text live by
        // selecting all in the in-cell editor, applying the font, then collapsing
        // the cursor to the end so continued typing appends in the new font.
        try {
          api?.executeCommand("doc.command.select-all");
          api?.executeCommand("doc.command.set-inline-format-font-family", { value: f });
          api?.executeCommand("doc.operation.move-cursor", { direction: "right" });
        } catch { /* editor selection is best-effort */ }
      }
      // Set the CELL (range) font — this applies to the SELECTED cell(s) only, which is what the
      // user expects. (We intentionally do NOT change the sheet's default style here; doing so
      // re-fonted every default-font cell in the whole sheet.)
      api?.executeCommand("sheet.command.set-range-font-family", { value: f });
      // Recompute the render skeleton so text RE-MEASURES with the now-loaded
      // font (a plain resize doesn't invalidate Univer's cached glyph metrics).
      try { (api?.getActiveWorkbook?.()?.getActiveSheet?.() as unknown as { refreshCanvas?: () => void })?.refreshCanvas?.(); }
      catch { try { window.dispatchEvent(new Event("resize")); } catch { /* */ } }
    };
    void ensureGoogleFont(f).then(apply);
  };
  const applyTextColor = (c: string) => {
    setTextColor(c);
    // Colour the WHOLE cell's text. The inline (in-cell editor) command only
    // recolours the current selection / the next-typed character — so text you
    // already typed stays its old colour, which reads as "it didn't work". The
    // range command colours everything in the cell (and the newly-typed text on
    // commit), matching what people expect from a spreadsheet colour picker.
    // Apply inline too while editing, so the change is visible live.
    if (editingRef.current) apiOf()?.executeCommand("doc.command.set-inline-format-text-color", { value: c });
    apiOf()?.executeCommand("sheet.command.set-range-text-color", { value: c });
  };
  const applyFill = (c: string) => {
    setFillColor(c);
    withRange((r) => r.setBackground(c));
  };
  const setNum = (p: string) => withRange((r) => r.setNumberFormat(p));
  const bumpDec = (d: number) => withRange((r) => r.setNumberFormat(adjustDecimals(r.getNumberFormat?.() || "General", d)));
  // Apply a border position with the chosen style (BorderStyleTypes) + colour.
  // Borders are a cell-level op with no in-editor equivalent, so if a cell is
  // mid-edit we commit it first (Keyboard/Enter) — otherwise the editor overlay
  // hides the border and the commit discards it.
  // Cell-level format ops (border, align, wrap, rotation) have no in-editor
  // equivalent, so commit any in-progress edit first (Enter), then apply to the
  // settled cell — otherwise the editor overlay hides/discards the change.
  const commitEdit = () => {
    if (editingRef.current) {
      apiOf()?.executeCommand("sheet.operation.set-cell-edit-visible", { visible: false, eventType: 4, keycode: 13 });
      editingRef.current = false;
    }
  };
  const cellOp = (fn: (r: RangeOps) => void) => {
    commitEdit();
    withRange(fn);
  };
  const applyBorder = (type: string, style: number, color: string) => cellOp((r) => r.setBorder(type, style, color));
  // Toggle the table filter (free-tier). Univer auto-expands the current
  // selection to the contiguous data range and adds filter dropdowns to every
  // header cell (sort + filter by condition/values). Click again to remove.
  const toggleFilter = () => {
    apiOf()?.executeCommand("sheet.command.smart-toggle-filter");
    setFilterOn((x) => !x);
  };
  // Rotation: numeric angle, or "vertical" (non-number → Univer vertical text).
  const applyRotation = (v: number | string) => {
    setRotation(v === "vertical" ? "vertical" : Number(v));
    cellOp((r) => (r.setTextRotation as (x: unknown) => RangeOps)(v));
  };
  // Current cell's text — prefills the link panel's display text.
  const getCellText = () => {
    const v = apiOf()?.getActiveWorkbook()?.getActiveRange()?.getValue();
    return v == null ? "" : String(v);
  };
  // Existing defined names → the "Named range" picker options.
  const getDefinedNamesList = () =>
    (apiOf()?.getActiveWorkbook()?.getDefinedNames?.() ?? []).map((d) => ({ name: d.getName(), ref: d.getFormulaOrRefString() }));
  // Insert link (free-tier hyperlink). URL → external link; Range / Named range
  // → Univer internal-navigation links (`#gid=<sheetId>&range=<A1:B2>`). A named
  // range is resolved to its cells so the link jumps there (no internal id
  // needed). Then setHyperLink stores the payload on the (editable) cell.
  const applyLink = (type: LinkType, label: string, value: string) => {
    const wb = apiOf()?.getActiveWorkbook();
    let payload = value;
    if (type === "url") {
      payload = /^(https?:|mailto:|tel:|#)/i.test(value) ? value : `https://${value}`;
    } else if (type === "range") {
      const sid = wb?.getActiveSheet()?.getSheetId() ?? "";
      payload = `#gid=${sid}&range=${value.replace(/\$/g, "").trim()}`;
    } else {
      // value = a defined name's ref, e.g. "Sheet1!$G$2:$G$141".
      const bang = value.indexOf("!");
      const sheetName = bang >= 0 ? value.slice(0, bang) : "";
      const rangePart = (bang >= 0 ? value.slice(bang + 1) : value).replace(/\$/g, "").trim();
      const sheet = (wb?.getSheets?.() ?? []).find((s) => s.getSheetName() === sheetName) ?? wb?.getActiveSheet();
      payload = `#gid=${sheet?.getSheetId() ?? ""}&range=${rangePart}`;
    }
    cellOp((r) => { void r.setHyperLink(payload, label || payload); });
  };

  // Close any open dropdown on outside click.
  useEffect(() => {
    if (!openId) return;
    const h = (ev: MouseEvent) => {
      const t = ev.target as HTMLElement;
      if (!t.closest("[data-levich-dd]")) setOpenId(null);
    };
    document.addEventListener("mousedown", h);
    return () => document.removeEventListener("mousedown", h);
  }, [openId]);

  // Reflect the selected cell's formatting in the toolbar (B/I/U/S pressed
  // state + font family/size), like Google Sheets, by listening to Univer's
  // selection-change events on the Facade. We also track whether a cell is being
  // edited so B/I/U/S can target the in-cell rich text directly.
  useEffect(() => {
    const a = apiOf();
    if (!a?.addEvent) return;
    // Cross-sheet sticky font: when this sheet mounts and a sticky font is active,
    // make it the sheet default so first-char typing is correct here too.
    if (stickyFontFamily) {
      try {
        const ws = a.getActiveWorkbook?.()?.getActiveSheet?.() as unknown as { getDefaultStyle?: () => unknown; setDefaultStyle?: (s: unknown) => void };
        const cur = ws?.getDefaultStyle?.();
        ws?.setDefaultStyle?.({ ...(cur && typeof cur === "object" ? cur : {}), ff: stickyFontFamily });
      } catch { /* */ }
    }
    const sync = () => {
      const r = a.getActiveWorkbook()?.getActiveRange();
      if (!r) return;
      const s = r.getCellStyleData?.() ?? {};
      setActive({ bold: s.bl === 1, italic: s.it === 1, underline: s.ul?.s === 1, strike: s.st?.s === 1 });
      setRotation(s.tr?.v === 1 ? "vertical" : s.tr?.a ?? 0);
      setFilterOn(!!a.getActiveWorkbook()?.getActiveSheet()?.getFilter?.());
      const fs = r.getFontSize?.();
      if (fs) setFontPt(fs);
      const ff = r.getFontFamily?.();
      if (ff) setFontFamily(ff);
      // Pre-arm the active EMPTY cell with the sticky font BEFORE editing starts,
      // so the in-cell editor inherits it and even the FIRST typed character uses
      // it (priming on edit-start runs after the trigger char). Only empty cells,
      // so existing content keeps its font.
      if (stickyFontFamily && ff !== stickyFontFamily) {
        try {
          const v = (r as unknown as { getValue?: () => unknown }).getValue?.();
          if (v === "" || v === null || v === undefined) {
            a.executeCommand("sheet.command.set-range-font-family", { value: stickyFontFamily });
            setFontFamily(stickyFontFamily);
          }
        } catch { /* */ }
      }
    };
    const ev = a.Event ?? {};
    const disposers = [
      a.addEvent?.(ev.SelectionChanged ?? "SelectionChanged", sync),
      a.addEvent?.(ev.SelectionMoveEnd ?? "SelectionMoveEnd", sync),
      a.addEvent?.(ev.SheetEditStarted ?? "SheetEditStarted", () => {
        editingRef.current = true;
        // Sticky font while TYPING: prime the in-cell editor with the current font
        // so characters appear in it live (not just after commit). Deferred a tick
        // so the editor's document/selection is ready to accept the format.
        if (!stickyFontFamily) return;
        const val = stickyFontFamily;
        setTimeout(() => { try { a.executeCommand("doc.command.set-inline-format-font-family", { value: val }); } catch { /* */ } }, 0);
      }),
      a.addEvent?.(ev.SheetEditEnded ?? "SheetEditEnded", () => {
        editingRef.current = false;
        // Sticky current font: the just-committed cell adopts the picked font, so
        // typing in ANY cell uses it until the font is changed. Font is already
        // loaded (it was picked), then re-measure so it paints correctly.
        if (!stickyFontFamily) return;
        try {
          a.executeCommand("sheet.command.set-range-font-family", { value: stickyFontFamily });
          (a.getActiveWorkbook()?.getActiveSheet() as unknown as { refreshCanvas?: () => void })?.refreshCanvas?.();
        } catch { /* best-effort */ }
      }),
      // Remember the last sort applied (full params) so a saved Filter View can
      // reproduce the sorted order.
      a.addEvent?.(ev.CommandExecuted ?? "CommandExecuted", (e) => {
        const evt = e as { id?: string; params?: object };
        if (evt?.id === "sheet.command.sort-range" && evt.params) lastSortRef.current = evt.params;
      }),
      // Keep the toolbar zoom indicator in sync when zoom changes from anywhere
      // (View ▸ Zoom menu, the bottom-right control, or the toolbar itself).
      a.addEvent?.(ev.SheetZoomChanged ?? "SheetZoomChanged", (e) => {
        const z = (e as { zoom?: number })?.zoom;
        if (typeof z === "number") setZoomPct(Math.round(z * 100));
      }),
      // Group-by collapse: clicking a group header's first cell folds/unfolds it.
      a.addEvent?.(ev.SelectionMoveEnd ?? "SelectionMoveEnd", () => {
        const run = groupRunRef.current;
        if (!run) return;
        const r = a.getActiveWorkbook()?.getActiveRange();
        const row = r?.getRow?.();
        const col = r?.getColumn?.();
        if (row == null || col == null || col !== run.startCol) return;
        if (!run.groups.some((g) => g.headerRow === row)) return;
        const now = Date.now();
        if (lastToggleRef.current.row === row && now - lastToggleRef.current.t < 350) return;
        lastToggleRef.current = { t: now, row };
        try {
          toggleGroup(gbApi(), run, row);
        } catch {
          /* best effort */
        }
      }),
    ];
    return () => disposers.forEach((d) => d?.dispose?.());
  }, [api]);

  // Load persisted Filter Views once the API is ready (Google auto-saves views
  // with the doc; we persist to localStorage per sheet).
  useEffect(() => {
    if (!api) return;
    const savedFilters = loadViews(fvApi());
    const savedGroups = loadGroupViews(gbApi());
    setViews(savedFilters);
    setGroupViews(savedGroups);
    const seqOf = (id: string) => Number(id.replace(/^(view|group)_/, "")) || 0;
    viewSeq.current = [...savedFilters, ...savedGroups].reduce((m, v) => Math.max(m, seqOf(v.id)), 0);
    viewsLoaded.current = true;
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [api]);

  // Auto-save views whenever they change (after the initial load).
  useEffect(() => {
    if (viewsLoaded.current) saveViews(fvApi(), views);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [views]);
  useEffect(() => {
    if (viewsLoaded.current) saveGroupViews(gbApi(), groupViews);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [groupViews]);

  const dd = { openId, setOpenId };

  // Left padding lines the Search icon up under "File" in the menu bar above:
  // the menu bar container is padded 10px and each label has an 8px inset, so a
  // 6px container pad + the icon button's own ~9px inset puts the magnifier
  // glyph directly under the "File" text.
  return (
    <>
    <div style={{ display: "flex", alignItems: "center", gap: 3, width: "100%", background: "#f9fafb", borderBottom: "1px solid #eaecf0", padding: "4px 8px 4px 9px", boxSizing: "border-box" }}>
      <div ref={stripRef} style={{ display: "flex", alignItems: "center", gap: 3, overflowX: "auto", flex: 1, scrollbarWidth: "none" }}>
        <IconBtn label="Search (find & replace)" onClick={() => onOpenFind?.()}><SearchMd size={SZ} /></IconBtn>
        <IconBtn label="Undo" onClick={() => apiOf()?.undo()}><FlipBackward size={SZ} /></IconBtn>
        <IconBtn label="Redo" onClick={() => apiOf()?.redo()}><FlipForward size={SZ} /></IconBtn>
        <IconBtn label="Print" onClick={() => printSheet(apiOf()?.getActiveWorkbook() ?? null)}><Printer size={SZ} /></IconBtn>
        <IconBtn label="Format painter" onClick={() => apiOf()?.executeCommand("sheet.command.set-once-format-painter")}><PaintPour size={SZ} /></IconBtn>
        <Dropdown id="zoom" {...dd} label="Zoom" trigger={<span style={{ fontSize: 13, minWidth: 38, textAlign: "left" }}>{zoomPct}%</span>} width={92}>
          {(close) =>
            ZOOMS.map((z) => (
              <MenuRow
                key={z}
                onClick={() => {
                  apiOf()?.getActiveWorkbook()?.getActiveSheet()?.zoom(z / 100);
                  setZoomPct(z);
                  close();
                }}
              >
                {z}%
              </MenuRow>
            ))
          }
        </Dropdown>
        <span style={divider} />

        <IconBtn label="Currency" onClick={() => setNum('"$"#,##0.00;("$"#,##0.00)')}><CurrencyDollar size={SZ} /></IconBtn>
        <IconBtn label="Percent" onClick={() => setNum("0.00%")}><Percent01 size={SZ} /></IconBtn>
        <IconBtn label="Decrease decimal places" onClick={() => bumpDec(-1)}><span style={{ fontSize: 12, fontWeight: 600 }}>.0﹣</span></IconBtn>
        <IconBtn label="Increase decimal places" onClick={() => bumpDec(1)}><span style={{ fontSize: 12, fontWeight: 600 }}>.0﹢</span></IconBtn>
        <Dropdown id="numFmt" {...dd} label="More formats" trigger={<span style={{ fontSize: 13, minWidth: 22, textAlign: "left" }}>123</span>} width={224}>
          {(close) =>
            NUMBER_FORMAT_GROUPS.map((group, gi) => (
              <div key={gi}>
                {gi > 0 && <div style={{ height: 1, background: "#eaecf0", margin: "4px 4px" }} />}
                {group.map((n) => (
                  <MenuRow key={n.label} onClick={() => { setNum(n.pattern); close(); }}>
                    <span>{n.label}</span>
                    {n.sample && <span style={{ marginLeft: "auto", color: "#98a2b3", fontSize: 12 }}>{n.sample}</span>}
                  </MenuRow>
                ))}
              </div>
            ))
          }
        </Dropdown>
        <span style={divider} />

        <Dropdown id="font" {...dd} label="Font" trigger={<span style={{ fontSize: 13, minWidth: 54, maxWidth: 80, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap", textAlign: "left" }}>{fontFamily}</span>} width={200}>
          {(close) => <FontMenu onPick={(f) => { applyFontFamily(f); close(); }} />}
        </Dropdown>
        <IconBtn label="Decrease font size" onClick={() => changeFont(-1)}><span style={{ fontSize: 18, fontWeight: 500 }}>−</span></IconBtn>
        <Dropdown id="size" {...dd} label="Font size" trigger={<span style={{ fontSize: 13, minWidth: 18, textAlign: "center" }}>{fontPt}</span>} width={84}>
          {(close) => FONT_SIZES.map((s) => <MenuRow key={s} onClick={() => { applyFontSize(s); close(); }}>{s}</MenuRow>)}
        </Dropdown>
        <IconBtn label="Increase font size" onClick={() => changeFont(1)}><span style={{ fontSize: 18, fontWeight: 500 }}>+</span></IconBtn>
        <span style={divider} />

        <IconBtn label="Bold" active={active.bold} onClick={toggleBold}><Bold01 size={SZ} /></IconBtn>
        <IconBtn label="Italic" active={active.italic} onClick={toggleItalic}><Italic01 size={SZ} /></IconBtn>
        <IconBtn label="Underline" active={active.underline} onClick={toggleUnderline}><Underline01 size={SZ} /></IconBtn>
        <IconBtn label="Strikethrough" active={active.strike} onClick={toggleStrike}><Strikethrough01 size={SZ} /></IconBtn>
        <span style={divider} />

        <Dropdown id="textColor" {...dd} label="Text color" trigger={<ColorIcon color={textColor}><Type01 size={SZ} /></ColorIcon>} width={212}>
          {(close) => <ColorPanel current={textColor} onApply={applyTextColor} onPick={(c) => { applyTextColor(c); close(); }} />}
        </Dropdown>
        <Dropdown id="fillColor" {...dd} label="Fill color" trigger={<ColorIcon color={fillColor}><Droplets01 size={SZ} /></ColorIcon>} width={212}>
          {(close) => <ColorPanel current={fillColor} onApply={applyFill} onPick={(c) => { applyFill(c); close(); }} />}
        </Dropdown>
        <Dropdown id="borders" {...dd} label="Borders" trigger={<Grid01 size={SZ} />} width={232}>
          {(close) => <BordersPanel apply={applyBorder} onPick={close} />}
        </Dropdown>
        <Dropdown id="merge" {...dd} label="Merge cells" trigger={<IntersectSquare size={SZ} />} width={160}>
          {(close) => (
            <>
              <MenuRow onClick={() => { cellOp((r) => r.merge()); close(); }}>Merge cells</MenuRow>
              <MenuRow onClick={() => { cellOp((r) => r.breakApart()); close(); }}>Unmerge</MenuRow>
            </>
          )}
        </Dropdown>
        <Dropdown id="link" {...dd} label="Insert link" trigger={<Link01 size={SZ} />} width={312}>
          {(close) => <LinkPanel initialText={getCellText()} definedNames={getDefinedNamesList()} onCancel={close} onApply={(type, label, value) => { applyLink(type, label, value); close(); }} />}
        </Dropdown>
        <span style={divider} />

        <Dropdown id="hAlign" {...dd} label="Horizontal align" trigger={<AlignLeft size={SZ} />} width={150}>
          {(close) => (
            <>
              <MenuRow onClick={() => { cellOp((r) => r.setHorizontalAlignment("left")); close(); }}><AlignLeft size={16} /> Left</MenuRow>
              <MenuRow onClick={() => { cellOp((r) => r.setHorizontalAlignment("center")); close(); }}><AlignCenter size={16} /> Center</MenuRow>
              {/* Univer's facade quirk: "right" throws; "normal" maps to RIGHT. */}
              <MenuRow onClick={() => { cellOp((r) => r.setHorizontalAlignment("normal")); close(); }}><AlignRight size={16} /> Right</MenuRow>
            </>
          )}
        </Dropdown>
        <Dropdown id="vAlign" {...dd} label="Vertical align" trigger={<AlignBottom01 size={SZ} />} width={150}>
          {(close) => (
            <>
              <MenuRow onClick={() => { cellOp((r) => r.setVerticalAlignment("top")); close(); }}><AlignTopGlyph /> Top</MenuRow>
              <MenuRow onClick={() => { cellOp((r) => r.setVerticalAlignment("middle")); close(); }}><AlignMiddleGlyph /> Middle</MenuRow>
              <MenuRow onClick={() => { cellOp((r) => r.setVerticalAlignment("bottom")); close(); }}><AlignBottom01 size={16} /> Bottom</MenuRow>
            </>
          )}
        </Dropdown>
        <Dropdown id="wrap" {...dd} label="Text wrapping" trigger={<WrapGlyph mode="wrap" />} width={160}>
          {(close) => (
            <>
              <MenuRow onClick={() => { cellOp((r) => r.setWrapStrategy(1)); close(); }}><WrapGlyph mode="overflow" /> Overflow</MenuRow>
              <MenuRow onClick={() => { cellOp((r) => r.setWrapStrategy(3)); close(); }}><WrapGlyph mode="wrap" /> Wrap</MenuRow>
              <MenuRow onClick={() => { cellOp((r) => r.setWrapStrategy(2)); close(); }}><WrapGlyph mode="clip" /> Clip</MenuRow>
            </>
          )}
        </Dropdown>
        <Dropdown id="rotate" {...dd} label="Text rotation" trigger={rotation === "vertical" ? <VerticalStackGlyph /> : <RotateGlyph deg={rotation} />} width={208}>
          {(close) => <RotationPanel current={rotation} apply={(v) => { applyRotation(v); }} onPick={close} />}
        </Dropdown>
        {/* Filter funnel + Filter views — next to Text rotation. */}
        <span style={divider} />
        <IconBtn label={filterOn ? "Remove filter" : "Create a filter"} active={filterOn} onClick={toggleFilter}><FilterFunnel01 size={SZ} /></IconBtn>
        <Dropdown id="filterviews" {...dd} label="Views (filter & group)" trigger={<Rows01 size={SZ} />} width={252}>
          {(close) => (
            <ViewsMenu
              columns={getGroupColumns(gbApi())}
              filterViews={views}
              groupViews={groupViews}
              activeViewId={activeView}
              editingViewId={editingViewId}
              onCreateFilterView={createFilterView}
              onCreateGroupBy={createGroupView}
              onApplyFilterView={applyView}
              onApplyGroupView={applyGroupView}
              onSaveView={saveView}
              onRefreshView={refreshView}
              onExitView={exitFilterView}
              onStartRename={setEditingViewId}
              onRename={renameView}
              onDelete={deleteView}
              close={close}
            />
          )}
        </Dropdown>
        <Dropdown id="functions" {...dd} label="Functions (Σ)" align="right" trigger={<span style={{ fontSize: 17, fontWeight: 600, lineHeight: 1 }}>Σ</span>} width={230}>
          {(close) => (
            <FunctionsMenu
              onQuick={(fn: Aggregate) => insertAggregate(fnApi(), fn)}
              onInsert={(name: string) => insertFunctionTemplate(fnApi(), name)}
              close={close}
            />
          )}
        </Dropdown>
      </div>
    </div>
    {groupBanner && (
      <div style={{ display: "flex", alignItems: "center", gap: 10, padding: "6px 14px", background: "#e6f4ea", borderBottom: "1px solid #b7e1c5", fontSize: 13, color: "#1e4620" }}>
        <Rows01 size={16} />
        <span style={{ fontWeight: 600 }}>{viewSaved ? "Group by" : "Temporary group by"} {groupBanner}</span>
        <span style={{ color: "#4b7e57" }}>· click a group header to fold/unfold</span>
        <button
          type="button"
          onMouseDown={keepEditorFocus}
          onClick={saveView}
          disabled={viewSaved}
          onMouseEnter={(e) => { if (!viewSaved) e.currentTarget.style.background = "#000"; }}
          onMouseLeave={(e) => { if (!viewSaved) e.currentTarget.style.background = "#101828"; }}
          style={{ marginLeft: "auto", padding: "5px 14px", borderRadius: 7, border: viewSaved ? "1px solid #b7e1c5" : "none", background: viewSaved ? "#fff" : "#101828", color: viewSaved ? "#1f7a3d" : "#fff", fontSize: 12, fontWeight: 600, cursor: viewSaved ? "default" : "pointer", fontFamily: "inherit", transition: "background-color .15s ease" }}
        >
          {viewSaved ? "✓ Saved" : "Save view"}
        </button>
        <button
          type="button"
          aria-label="Exit group view"
          onMouseDown={keepEditorFocus}
          onClick={exitFilterView}
          style={{ width: 26, height: 26, borderRadius: 7, border: "1px solid #b7e1c5", background: "#fff", color: "#1e4620", cursor: "pointer", fontSize: 14, lineHeight: 1, display: "inline-flex", alignItems: "center", justifyContent: "center" }}
        >
          ✕
        </button>
      </div>
    )}
    </>
  );
}

/**
 * Add/remove one decimal place from a number-format pattern, preserving any
 * currency symbol, `%`, parentheses, and multiple sections (Google-Sheets-like).
 * Adjusts each numeric token (`#,##0` or `0`) with its optional `.0…` run.
 */
function adjustDecimals(pattern: string, delta: number): string {
  if (!pattern || pattern === "General") return delta > 0 ? "0.0" : "0";
  return pattern.replace(/(#,##0|0)(\.0+)?/g, (_m, intPart: string, dec?: string) => {
    const cur = dec ? dec.length - 1 : 0; // dec includes the leading "."
    const next = Math.max(0, Math.min(10, cur + delta));
    return next > 0 ? `${intPart}.${"0".repeat(next)}` : intPart;
  });
}

export default LevichToolbar;
