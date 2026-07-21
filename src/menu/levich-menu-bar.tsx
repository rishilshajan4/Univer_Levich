/**
 * LevichMenuBar — a Google-Sheets-style menu bar (File · Edit · View · Insert ·
 * Format · Data · Tools · Extensions · Help). UI/UX only: each action drives
 * Univer via its Facade API. Items free Univer can't do are shown disabled.
 *
 * Untitled UI icons · portal dropdowns + nested submenu flyouts · our styling.
 */
import { useEffect, useRef, useState, type CSSProperties } from "react";
import { createPortal } from "react-dom";
import { Z_BASE } from "../core/z-index";
import { ChevronRight } from "@untitledui/icons";
import { exportToXlsx, type SnapshotSource } from "../core/export-xlsx";
import { downloadCsv, downloadHtml, downloadTsv } from "../core/export-data";
import { parseFileToGrid, pickFile, stashImportPayload, stashSnapshotPayload } from "../core/import-data";
import { parseXlsxToSnapshot, inlineSheetStyles, placeImportImages, sheetIndexToName, buildRichPlacement, firstVisibleSheet, type ImportImage } from "../core/xlsx-to-snapshot";
import type { WorkbookData } from "../core/types";
import { printSheet } from "../core/print-sheet";
import { attachKeyboardShortcuts, type ShortcutContext } from "../features/keyboard-shortcuts";
import { ImportModal, type ImportLocation } from "../components/import-modal";
import { RenameModal } from "../components/rename-modal";
import type { UniverAPI } from "../core/create-sheet";

/* ---- Loose Facade views --------------------------------------------------- */
interface RangeOps {
  getRow(): number;
  getColumn(): number;
  setValues(v: (string | number)[][]): unknown;
  setFontWeight(w: string | null): RangeOps;
  setFontStyle(s: string | null): RangeOps;
  setFontLine(l: string | null): RangeOps;
  setFontSize(n: number): RangeOps;
  setHorizontalAlignment(a: string): RangeOps;
  setVerticalAlignment(a: string): RangeOps;
  setWrap(b: boolean): RangeOps;
  setTextRotation(n: number): RangeOps;
  setNumberFormat(p: string): RangeOps;
  merge(): RangeOps;
  breakApart(): RangeOps;
}
interface SheetOps {
  getActiveRange(): RangeOps | null;
  getRange(row: number, column: number, numRows?: number, numColumns?: number): RangeOps | null;
  setName(name: string): unknown;
  getSheetName?(): string;
  insertRowBefore(i: number): unknown;
  insertRowAfter(i: number): unknown;
  insertColumnBefore(i: number): unknown;
  insertColumnAfter(i: number): unknown;
  deleteRows(i: number, n: number): unknown;
  deleteColumns(i: number, n: number): unknown;
  setFrozenRows(n: number): unknown;
  setFrozenColumns(n: number): unknown;
  cancelFreeze(): unknown;
  zoom(ratio: number): unknown;
  getMaxRows?(): number;
  getMaxColumns?(): number;
  setRowCount?(n: number): unknown;
  setColumnCount?(n: number): unknown;
  setColumnWidth?(column: number, width: number): unknown;
  setRowHeightsForced?(startRow: number, numRows: number, height: number): unknown;
  hideSheet?(): unknown;
  showSheet?(): unknown;
  isSheetHidden?(): boolean;
  activate?(): unknown;
}
interface WorkbookOps extends SnapshotSource {
  getName?(): string;
  setName?(name: string): unknown;
  getActiveSheet(): SheetOps | null;
  getActiveRange(): RangeOps | null;
  create(name: string, rows: number, columns: number, options?: { index?: number; sheet?: Record<string, unknown> }): SheetOps;
  getSheets(): SheetOps[];
  setActiveSheet(sheet: SheetOps | string): unknown;
  deleteSheet(sheet: SheetOps | string): unknown;
}
interface MenuApi {
  undo(): unknown;
  redo(): unknown;
  executeCommand(id: string, params?: object): unknown;
  getActiveWorkbook(): WorkbookOps | null;
  addEvent?(event: string, cb: (params: unknown) => void): { dispose?: () => void } | undefined;
  Event?: Record<string, string>;
}

/* ---- Menu model ----------------------------------------------------------- */
interface MItem {
  label: string;
  shortcut?: string;
  onClick?: () => void;
  items?: MItem[];
  disabled?: boolean;
  sep?: boolean; // render a divider before this item
  checked?: boolean; // when defined, renders a left check column (✓ when true)
}
interface Menu {
  label: string;
  items: MItem[];
}

/* ---- Styles --------------------------------------------------------------- */
const NUMBER_FMTS: Array<[string, string]> = [
  ["General", "General"],
  ["Number", "#,##0.00"],
  ["Currency", '"$"#,##0.00;("$"#,##0.00)'],
  ["Accounting", '_("$"* #,##0.00_);_("$"* (#,##0.00);_("$"* "-"??_)'],
  ["Percent", "0.00%"],
  ["Date", "yyyy-mm-dd"],
  ["Time", "h:mm:ss"],
];
const FONT_SIZES = [8, 9, 10, 11, 12, 14, 18, 24, 36];
const ZOOM_LEVELS = [50, 75, 90, 100, 125, 150, 200];

function toggleFullScreen(): void {
  if (typeof document === "undefined") return;
  if (document.fullscreenElement) void document.exitFullscreen?.();
  else void document.documentElement.requestFullscreen?.();
}

/** 0-based column index → A1 letters (0 → A, 25 → Z, 26 → AA). */
function colLetter(index: number): string {
  let n = index;
  let s = "";
  do {
    s = String.fromCharCode(65 + (n % 26)) + s;
    n = Math.floor(n / 26) - 1;
  } while (n >= 0);
  return s;
}

const menuLabel: CSSProperties = { padding: "2px 8px", borderRadius: 6, fontSize: 14, color: "#1f2937", cursor: "pointer", lineHeight: "22px" };
const panel: CSSProperties = { background: "#fff", border: "1px solid #eaecf0", borderRadius: 10, boxShadow: "0 8px 24px rgba(16,24,40,0.14)", padding: 6, minWidth: 220, zIndex: Z_BASE + 1000 };
const row: CSSProperties = { display: "flex", alignItems: "center", gap: 10, width: "100%", padding: "7px 10px", borderRadius: 6, border: "none", background: "transparent", fontSize: 13, color: "#344054", textAlign: "left", cursor: "pointer", whiteSpace: "nowrap" };

function MenuItemRow({ item, onClose }: { item: MItem; onClose: () => void }) {
  const [open, setOpen] = useState(false);
  const ref = useRef<HTMLDivElement>(null);
  const [pos, setPos] = useState({ top: 0, left: 0 });
  const hasSub = !!item.items?.length;

  const enter = () => {
    if (!hasSub || item.disabled) return;
    const r = ref.current?.getBoundingClientRect();
    if (r) setPos({ top: r.top - 6, left: r.right + 2 });
    setOpen(true);
  };

  return (
    <div ref={ref} style={{ position: "relative" }} onMouseEnter={enter} onMouseLeave={() => setOpen(false)}>
      {item.sep && <div style={{ height: 1, background: "#eaecf0", margin: "4px 0" }} />}
      <button
        type="button"
        disabled={item.disabled}
        onClick={() => {
          if (item.disabled || hasSub) return;
          item.onClick?.();
          onClose();
        }}
        style={{ ...row, color: item.disabled ? "#98a2b3" : row.color, cursor: item.disabled ? "default" : "pointer" }}
        onMouseEnter={(e) => !item.disabled && (e.currentTarget.style.background = "#f9fafb")}
        onMouseLeave={(e) => (e.currentTarget.style.background = "transparent")}
      >
        {item.checked !== undefined && <span style={{ width: 16, marginLeft: -2, display: "inline-flex", justifyContent: "center", color: "#101828" }}>{item.checked ? "✓" : ""}</span>}
        <span>{item.label}</span>
        <span style={{ marginLeft: "auto", display: "flex", alignItems: "center", gap: 6, color: "#98a2b3", fontSize: 12 }}>
          {item.shortcut}
          {hasSub && <ChevronRight size={15} />}
        </span>
      </button>
      {open &&
        hasSub &&
        createPortal(
          <div data-levich-menu style={{ position: "fixed", top: pos.top, left: pos.left, ...panel }}>
            {item.items!.map((sub, i) => (
              <MenuItemRow key={i} item={sub} onClose={onClose} />
            ))}
          </div>,
          document.body,
        )}
    </div>
  );
}

export interface LevichMenuBarProps {
  api: UniverAPI | null;
  /** Download handler (defaults to exporting the live sheet to .xlsx). */
  onDownload?: () => void;
  /** Open the Levich Find & Replace modal (replaces Univer's native panel). */
  onOpenFind?: () => void;
  /**
   * File ▸ Save / ⌘S. Receives the live workbook snapshot to persist. Return
   * `true` if the host handled it; falsy → built-in localStorage fallback +
   * "Saved" toast. The browser save-page dialog is always suppressed.
   */
  onSave?: (snapshot: WorkbookData) => boolean | void;
  /** File ▸ New — host-defined "blank document". Defaults to clearing the sheet. */
  onNew?: () => void;
  /**
   * File ▸ Import hook. Called with the parsed grid + the chosen destination.
   * Return `true` to signal the HOST handled it (e.g. FinOpz BE created or
   * replaced a document) — the built-in local behavior is then skipped. Return
   * falsy (or omit the prop) to let the sheet apply its built-in behavior.
   *
   * Document-level destinations ("new-spreadsheet", "replace-spreadsheet") are
   * exactly the cases a real backend owns; the in-document ones (insert sheet,
   * replace current sheet, append, at-cell) can be left to the sheet.
   */
  onImport?: (grid: (string | number)[][], location: ImportLocation) => boolean | void;
  /**
   * File ▸ Import RAW-FILE hook — fired with the picked File BEFORE parsing.
   * Return `true` (sync/async) to let the HOST own the whole import (e.g. upload
   * to the backend) and skip the built-in client-side parse. Falsy → fall through.
   */
  onImportFile?: (file: File) => boolean | Promise<boolean>;
  /** File ▸ Make a copy. Defaults to downloading a copy as .xlsx. */
  onMakeCopy?: () => void;
  /** File ▸ Rename. Defaults to renaming the active sheet (prompt). */
  onRename?: (name: string) => void;
  /** View ▸ Hide sheet — host hook (hides the ACTIVE sheet). Defaults to the Facade. */
  onHideActiveSheet?: () => void;
  /** View ▸ Show sheets ▸ — host hook to unhide + open a sheet by id. */
  onShowSheet?: (sheetId: string) => void;
  /** Hidden sheets for the Show-sheets submenu (host-driven). */
  hiddenSheetList?: Array<{ sheetId: string; name: string }>;
  /** Whether the active sheet can be hidden (host-driven). */
  canHideActiveSheet?: boolean;
}

export function LevichMenuBar({ api, onDownload, onOpenFind, onSave, onNew, onImport, onImportFile, onMakeCopy, onRename, onHideActiveSheet, onShowSheet, hiddenSheetList, canHideActiveSheet }: LevichMenuBarProps) {
  const [openMenu, setOpenMenu] = useState<number | null>(null);
  // Transient "Saved" toast (File ▸ Save / ⌘S).
  const [saved, setSaved] = useState(false);
  const savedTimer = useRef<ReturnType<typeof setTimeout> | null>(null);
  // File ▸ Rename modal (replaces window.prompt). Holds the name being edited +
  // the set of other sheet names to block duplicates.
  const [renameState, setRenameState] = useState<{ current: string; taken?: Set<string> } | null>(null);
  const barRef = useRef<HTMLDivElement>(null);
  const [pos, setPos] = useState({ top: 0, left: 0 });
  // View ▸ Show toggle state.
  const [formulaBarOn, setFormulaBarOn] = useState(true);
  const [gridlinesOn, setGridlinesOn] = useState(true);
  // Active selection row/col (0-based) for "Freeze up to row/column" — captured
  // when a menu opens so the labels reflect the current selection.
  const [selRow, setSelRow] = useState(0);
  const [selCol, setSelCol] = useState(0);
  // Import-file modal state (the parsed grid awaiting a destination choice).
  const [importGrid, setImportGrid] = useState<(string | number)[][] | null>(null);
  const [importFileName, setImportFileName] = useState("");
  // The picked File (kept so whole-document xlsx destinations can be re-read
  // into a rich Univer snapshot — styles / merges / formats / multiple sheets).
  const importFileRef = useRef<File | null>(null);

  const toggleFormulaBar = () => {
    const el = typeof document !== "undefined" ? (document.querySelector('[data-u-comp="formula-bar"]') as HTMLElement | null) : null;
    const next = !formulaBarOn;
    if (el) el.style.display = next ? "" : "none";
    setFormulaBarOn(next);
  };
  const toggleGridlines = () => {
    apiOf()?.executeCommand("sheet.command.toggle-gridlines");
    setGridlinesOn((v) => !v);
  };

  const apiOf = () => api as unknown as MenuApi | null;
  const wb = () => apiOf()?.getActiveWorkbook() ?? null;
  const sheet = () => wb()?.getActiveSheet() ?? null;
  const withRange = (fn: (r: RangeOps) => void) => {
    const r = wb()?.getActiveRange();
    if (r) fn(r);
  };
  const exec = (id: string) => apiOf()?.executeCommand(id);

  const download = () => {
    if (onDownload) return onDownload();
    const w = wb();
    if (w) void exportToXlsx(w as SnapshotSource, "levich-sheet.xlsx");
  };
  const downloadAs = (fmt: "csv" | "tsv" | "html") => {
    const w = wb();
    if (!w) return;
    if (fmt === "csv") return downloadCsv(w as SnapshotSource, "levich-sheet.csv");
    if (fmt === "tsv") return downloadTsv(w as SnapshotSource, "levich-sheet.tsv");
    downloadHtml(w as SnapshotSource, "levich-sheet.html");
  };
  const doPrint = () => printSheet(wb());

  // File ▸ Save / ⌘S. Hands the live snapshot to the host; if the host doesn't
  // claim it, fall back to persisting in localStorage so a standalone sheet
  // still "saves". A brief toast confirms either way.
  const doSave = () => {
    const w = wb();
    if (!w) return;
    const snapshot = w.getSnapshot() as WorkbookData;
    if (onSave?.(snapshot)) return flashSaved(); // host handled it
    try {
      const id = String((snapshot as { id?: unknown }).id ?? "workbook");
      if (typeof localStorage !== "undefined") localStorage.setItem(`levich:save:${id}`, JSON.stringify(snapshot));
    } catch {
      /* quota / serialization — still flash so the shortcut feels responsive */
    }
    flashSaved();
  };
  const flashSaved = () => {
    setSaved(true);
    if (savedTimer.current) clearTimeout(savedTimer.current);
    savedTimer.current = setTimeout(() => setSaved(false), 1600);
  };

  /** {maxRow, maxCol} of the sheet's NON-EMPTY used range from the live snapshot
   *  (-1/-1 if empty). Empty cells left behind by clearing (v: "") are ignored,
   *  so Append lands right after the real data instead of below a blank gap. */
  const usedBounds = (): { maxR: number; maxC: number } => {
    const w = wb();
    if (!w) return { maxR: -1, maxC: -1 };
    const snap = w.getSnapshot();
    const sid = snap.sheetOrder?.[0] ?? Object.keys(snap.sheets ?? {})[0];
    const cd = (sid ? snap.sheets?.[sid]?.cellData : undefined) ?? {};
    let maxR = -1;
    let maxC = -1;
    for (const [r, cols] of Object.entries(cd)) {
      for (const [c, cell] of Object.entries(cols ?? {})) {
        const cc = cell as { v?: unknown; f?: unknown } | undefined;
        const hasContent = !!cc && ((cc.v !== undefined && cc.v !== null && String(cc.v).trim() !== "") || (cc.f != null && cc.f !== ""));
        if (hasContent) {
          maxR = Math.max(maxR, Number(r));
          maxC = Math.max(maxC, Number(c));
        }
      }
    }
    return { maxR, maxC };
  };
  /** Write a value grid into the active sheet at (row0, col0). Expands the sheet
   *  first if the grid runs past the current row/column count (otherwise
   *  getRange throws "out of bounds" and the import silently does nothing). */
  const writeGridAt = (grid: (string | number)[][], row0: number, col0: number) => {
    const s = sheet();
    if (!s || !grid.length) return;
    const cols = grid.reduce((m, r) => Math.max(m, r.length), 0);
    const norm = grid.map((r) => {
      const a = r.slice();
      while (a.length < cols) a.push("");
      return a;
    });
    try {
      const needRows = row0 + norm.length;
      const needCols = col0 + cols;
      if ((s.getMaxRows?.() ?? Infinity) < needRows) s.setRowCount?.(needRows);
      if ((s.getMaxColumns?.() ?? Infinity) < needCols) s.setColumnCount?.(needCols);
      s.getRange(row0, col0, norm.length, cols)?.setValues(norm);
    } catch {
      /* best effort — never crash the import */
    }
  };
  /** Write ONE imported sheet's rich content (values + inline styles + merges)
   *  into the active sheet at (row0, col0). Expands the sheet if needed. Sparse
   *  ICellData only touches the imported cells, so Append doesn't clobber the
   *  rest of the sheet. */
  const writeRichSheetAt = (sheetSnap: Record<string, unknown> | null, styles: Record<string, unknown>, row0: number, col0: number, applyLayout: boolean) => {
    const s = sheet();
    if (!s || !sheetSnap) return;
    const { cells, merges, maxRow, maxCol } = buildRichPlacement(sheetSnap, styles, row0, col0);
    try {
      if ((s.getMaxRows?.() ?? Infinity) < maxRow + 1) s.setRowCount?.(maxRow + 1);
      if ((s.getMaxColumns?.() ?? Infinity) < maxCol + 1) s.setColumnCount?.(maxCol + 1);
      (s.getRange(0, 0, maxRow + 1, maxCol + 1) as unknown as { setValues: (v: unknown) => unknown } | null)?.setValues(cells);
      for (const m of merges) {
        try {
          (s.getRange(m.row, m.col, m.numRows, m.numCols) as unknown as { merge?: () => void } | null)?.merge?.();
        } catch {
          /* a conflicting/duplicate merge — skip */
        }
      }
      // Match the source layout — apply the imported column widths + row heights.
      // Without this, narrow default columns wrap the text and Univer auto-grows
      // rows, throwing every merge out of alignment. `setRowHeightsForced` also
      // pins heights so they don't auto-grow. Only for whole-sheet replacement
      // (offset 0,0) — append/at-cell must not disturb the existing layout.
      if (applyLayout && col0 === 0 && row0 === 0) {
        const columnData = (sheetSnap.columnData as Record<string, { w?: number }> | undefined) ?? {};
        for (const [c, cw] of Object.entries(columnData)) {
          if (cw?.w && cw.w > 0) s.setColumnWidth?.(Number(c), cw.w);
        }
        const rowData = (sheetSnap.rowData as Record<string, { h?: number }> | undefined) ?? {};
        for (const [r, rh] of Object.entries(rowData)) {
          if (rh?.h && rh.h > 0) s.setRowHeightsForced?.(Number(r), 1, rh.h);
        }
      }
    } catch {
      /* best effort — never crash the import */
    }
  };
  /** Blank every cell in the current used range. */
  const clearUsedRange = () => {
    const s = sheet();
    const { maxR, maxC } = usedBounds();
    if (!s || maxR < 0) return;
    const empties = Array.from({ length: maxR + 1 }, () => Array.from({ length: maxC + 1 }, () => ""));
    s.getRange(0, 0, maxR + 1, maxC + 1)?.setValues(empties);
  };

  // File ▸ Import: pick an .xlsx/.csv, parse it, then ask where to put it.
  const importFile = async () => {
    const file = await pickFile(".xlsx,.csv");
    if (!file) return;
    // Raw-file host hook first: if the backend owns the import (uploads + converts
    // server-side), it returns true and we skip the heavy in-browser parse entirely.
    if (onImportFile && (await onImportFile(file))) return;
    const grid = await parseFileToGrid(file).catch(() => null);
    if (!grid || !grid.length) return;
    importFileRef.current = file;
    setImportGrid(grid);
    setImportFileName(file.name);
  };
  // Whole-document destinations (new / replace spreadsheet) carry a RICH Univer
  // snapshot when the source is an .xlsx — preserving styles, colours, merges,
  // number formats and every worksheet. CSV (or a failed parse) falls back to
  // the flat-grid payload. Returns true if a rich snapshot was stashed.
  const stashWholeDoc = async (grid: (string | number)[][], snapshot: WorkbookData | null): Promise<void> => {
    if (snapshot && stashSnapshotPayload(snapshot)) return; // rich path
    stashImportPayload(grid);
  };

  const applyImport = async (location: ImportLocation) => {
    const grid = importGrid;
    setImportGrid(null);
    if (!grid) return;
    // Host hook first. If FinOpz BE handles this destination (typically
    // "new-spreadsheet" / "replace-spreadsheet" = create/replace a real
    // document), it returns true and we stop. Otherwise fall through to the
    // built-in local behavior.
    if (onImport?.(grid, location)) return;

    // Parse the RICH snapshot ONCE for .xlsx sources — every destination below
    // uses it for full fidelity (styles / merges / formats / images), falling
    // back to the flat grid for CSV or a parse failure.
    const file = importFileRef.current;
    const isXlsx = !!file && /\.(xlsx|xls)$/i.test(file.name);
    const richSnap: WorkbookData | null = isXlsx && file ? await parseXlsxToSnapshot(file).catch(() => null) : null;
    const richStyles = (richSnap?.styles as Record<string, unknown> | undefined) ?? {};

    if (location === "new-spreadsheet") {
      // Built-in fallback (no backend): open the imported data as a fresh
      // spreadsheet in a NEW browser tab at the same app URL.
      await stashWholeDoc(grid, richSnap);
      if (typeof window !== "undefined") window.open(window.location.href, "_blank");
      return;
    }

    if (location === "replace-spreadsheet") {
      // Replace the WHOLE document: reload THIS tab as a fresh spreadsheet of
      // the imported data. Guaranteed clean — no leftover sheets, no residue.
      await stashWholeDoc(grid, richSnap);
      if (typeof window !== "undefined") window.location.reload();
      return;
    }

    if (location === "new-sheets") {
      // Add NEW sheet tab(s) with the imported data; keep existing sheets.
      const w = wb();
      if (!w) return;

      // Rich path (.xlsx): recreate each VISIBLE source sheet with full
      // fidelity — styles, merges, number formats, widths + floating images.
      if (richSnap) {
        const order = (richSnap.sheetOrder as string[] | undefined) ?? [];
        const sheetsMap = (richSnap.sheets as Record<string, Record<string, unknown>> | undefined) ?? {};
        const existing = new Set(w.getSheets().map((s) => s.getSheetName?.() ?? ""));
        let firstName = "";
        for (const id of order) {
          const sheetSnap = sheetsMap[id];
          if (!sheetSnap || sheetSnap.hidden === 1) continue; // only visible source sheets
          const baseName = String(sheetSnap.name ?? "Imported");
          let name = baseName;
          let n = 2;
          while (existing.has(name)) name = `${baseName} ${n++}`;
          existing.add(name);
          if (!firstName) firstName = name;
          const { id: _omitId, ...resolved } = inlineSheetStyles(sheetSnap, richStyles);
          const rows = Number(sheetSnap.rowCount ?? 100);
          const cols = Number(sheetSnap.columnCount ?? 26);
          try {
            w.create(name, rows, cols, { sheet: { ...resolved, name } });
          } catch {
            /* skip a sheet that fails to create */
          }
        }
        if (firstName) w.setActiveSheet(firstName);
        const imgs = (richSnap as { drawingsImport?: ImportImage[] }).drawingsImport;
        if (imgs?.length) await placeImportImages(api, imgs, sheetIndexToName(richSnap));
        return;
      }

      // Flat path (CSV / xlsx parse failure): single sheet of plain values.
      const numCols = Math.max(grid.reduce((m, r) => Math.max(m, r.length), 0) + 2, 26);
      const numRows = Math.max(grid.length + 20, 100);
      const existing = new Set(w.getSheets().map((s) => s.getSheetName?.() ?? ""));
      let name = "Imported";
      let i = 2;
      while (existing.has(name)) name = `Imported ${i++}`;
      try {
        w.create(name, numRows, numCols);
        w.setActiveSheet(name);
        writeGridAt(grid, 0, 0);
      } catch {
        /* sheet creation differs — best effort */
      }
      return;
    }

    // In-place destinations: bring ONE imported sheet (the first visible one)
    // into the ACTIVE sheet with full formatting when .xlsx; else plain values.
    // (Multi-sheet workbooks: use "Insert new sheet(s)" or "Replace spreadsheet"
    // to keep every sheet.)
    const richSheet = richSnap ? firstVisibleSheet(richSnap) : null;

    if (location === "append") {
      const at = usedBounds().maxR + 1;
      if (richSheet) writeRichSheetAt(richSheet, richStyles, at, 0, false);
      else writeGridAt(grid, at, 0);
      return;
    }
    if (location === "at-cell") {
      if (richSheet) writeRichSheetAt(richSheet, richStyles, selRow, selCol, false);
      else writeGridAt(grid, selRow, selCol);
      return;
    }

    // replace-sheet → clear + write the ACTIVE sheet only (keeps other tabs).
    clearUsedRange();
    if (richSheet) writeRichSheetAt(richSheet, richStyles, 0, 0, true);
    else writeGridAt(grid, 0, 0);
  };
  const makeCopy = () => {
    if (onMakeCopy) return onMakeCopy();
    const w = wb();
    if (w) void exportToXlsx(w as SnapshotSource, "levich-sheet-copy.xlsx");
  };
  // Open the styled Rename modal for the SPREADSHEET (document) name — seeded
  // with the current workbook name. (Renaming an individual sheet is done from
  // the sheet-tab menu.) No duplicate check — spreadsheet names aren't unique.
  const rename = () => {
    const current = wb()?.getName?.() || "Untitled spreadsheet";
    setRenameState({ current });
  };
  const applyRename = (name: string) => {
    setRenameState(null);
    try {
      wb()?.setName?.(name); // update the Univer workbook name
    } catch {
      /* ignore */
    }
    onRename?.(name); // let the host update its displayed document title
  };
  const newDoc = () => {
    if (onNew) return onNew();
    if (typeof window !== "undefined" && !window.confirm("Clear all cells and start a blank sheet?")) return;
    clearUsedRange();
  };

  // --- Sheet visibility (View ▸ Hide sheet / Show sheets) --------------------
  const visibleSheets = (): SheetOps[] => (wb()?.getSheets() ?? []).filter((s) => !s.isSheetHidden?.());
  const hiddenSheets = (): SheetOps[] => (wb()?.getSheets() ?? []).filter((s) => !!s.isSheetHidden?.());
  const hideActiveSheet = () => {
    if (visibleSheets().length <= 1) return; // never hide the last visible sheet
    sheet()?.hideSheet?.();
  };
  const showHiddenSheet = (s: SheetOps) => {
    try {
      s.showSheet?.();
      s.activate?.();
    } catch {
      /* best effort */
    }
  };
  // "Show sheets ▸" submenu — host-driven (manifest) when provided, else the Facade.
  const showSheetItems = (): MItem[] => {
    if (onShowSheet) {
      const h = hiddenSheetList ?? [];
      if (!h.length) return [{ label: "No hidden sheets", disabled: true }];
      return h.map((s) => ({ label: s.name, onClick: () => onShowSheet(s.sheetId) }));
    }
    const h = hiddenSheets();
    if (!h.length) return [{ label: "No hidden sheets", disabled: true }];
    return h.map((s) => ({ label: s.getSheetName?.() ?? "Sheet", onClick: () => showHiddenSheet(s) }));
  };

  const menus: Menu[] = [
    {
      label: "File",
      items: [
        { label: "New", onClick: newDoc },
        { label: "Import", shortcut: "⌘O", onClick: importFile, sep: true },
        { label: "Make a copy", onClick: makeCopy },
        { label: "Save", shortcut: "⌘S", onClick: doSave, sep: true },
        {
          label: "Download",
          sep: true,
          items: [
            { label: "Microsoft Excel (.xlsx)", onClick: download },
            { label: "Comma-separated values (.csv)", onClick: () => downloadAs("csv") },
            { label: "Tab-separated values (.tsv)", onClick: () => downloadAs("tsv") },
            { label: "Web page (.html)", onClick: () => downloadAs("html") },
          ],
        },
        { label: "Rename", onClick: rename, sep: true },
        { label: "Print", shortcut: "⌘P", onClick: doPrint },
      ],
    },
    {
      label: "Edit",
      items: [
        { label: "Undo", shortcut: "⌘Z", onClick: () => apiOf()?.undo() },
        { label: "Redo", shortcut: "⌘Y", onClick: () => apiOf()?.redo() },
        { label: "Cut", shortcut: "⌘X", onClick: () => exec("sheet.command.cut"), sep: true },
        { label: "Copy", shortcut: "⌘C", onClick: () => exec("sheet.command.copy") },
        { label: "Paste", shortcut: "⌘V", onClick: () => exec("sheet.command.paste") },
        {
          label: "Delete",
          sep: true,
          items: [
            { label: "Values", onClick: () => exec("sheet.command.clear-selection-content") },
            { label: "Row", onClick: () => withRange((r) => sheet()?.deleteRows(r.getRow(), 1)), sep: true },
            { label: "Column", onClick: () => withRange((r) => sheet()?.deleteColumns(r.getColumn(), 1)) },
          ],
        },
        { label: "Find and replace", shortcut: "⌘⇧H", onClick: () => onOpenFind?.(), sep: true },
      ],
    },
    {
      label: "View",
      items: [
        {
          label: "Show",
          items: [
            { label: "Formula bar", checked: formulaBarOn, onClick: toggleFormulaBar },
            { label: "Gridlines", checked: gridlinesOn, onClick: toggleGridlines },
            { label: "Formulae", shortcut: "⌃`", checked: false, disabled: true },
            { label: "Protected ranges", checked: false, disabled: true },
          ],
        },
        {
          label: "Freeze",
          sep: true,
          items: [
            { label: "No rows", onClick: () => sheet()?.setFrozenRows(0) },
            { label: "1 row", onClick: () => sheet()?.setFrozenRows(1) },
            { label: "2 rows", onClick: () => sheet()?.setFrozenRows(2) },
            { label: `Up to row ${selRow + 1}`, onClick: () => sheet()?.setFrozenRows(selRow + 1) },
            { label: "No columns", onClick: () => sheet()?.setFrozenColumns(0), sep: true },
            { label: "1 column", onClick: () => sheet()?.setFrozenColumns(1) },
            { label: "2 columns", onClick: () => sheet()?.setFrozenColumns(2) },
            { label: `Up to column ${colLetter(selCol)}`, onClick: () => sheet()?.setFrozenColumns(selCol + 1) },
          ],
        },
        {
          label: "Zoom",
          sep: true,
          items: ZOOM_LEVELS.map((z) => ({ label: `${z}%`, onClick: () => sheet()?.zoom(z / 100) })),
        },
        { label: "Hide sheet", sep: true, disabled: onHideActiveSheet ? canHideActiveSheet === false : visibleSheets().length <= 1, onClick: onHideActiveSheet ?? hideActiveSheet },
        { label: "Show sheets", items: showSheetItems() },
        { label: "Full screen", onClick: toggleFullScreen, sep: true },
      ],
    },
    {
      label: "Insert",
      items: [
        { label: "Row above", onClick: () => withRange((r) => sheet()?.insertRowBefore(r.getRow())) },
        { label: "Row below", onClick: () => withRange((r) => sheet()?.insertRowAfter(r.getRow())) },
        { label: "Column left", onClick: () => withRange((r) => sheet()?.insertColumnBefore(r.getColumn())), sep: true },
        { label: "Column right", onClick: () => withRange((r) => sheet()?.insertColumnAfter(r.getColumn())) },
        { label: "Note", onClick: () => exec("sheet.command.toggle-note-popup"), sep: true },
        { label: "Link", onClick: () => exec("sheet.operation.insert-hyper-link") },
        { label: "Chart", disabled: true, sep: true },
        { label: "Pivot table", disabled: true },
        { label: "Image", disabled: true },
      ],
    },
    {
      label: "Format",
      items: [
        { label: "Theme", disabled: true },
        {
          label: "Number",
          sep: true,
          items: NUMBER_FMTS.map(([label, p]) => ({ label, onClick: () => withRange((r) => r.setNumberFormat(p)) })),
        },
        {
          label: "Text",
          items: [
            { label: "Bold", shortcut: "⌘B", onClick: () => withRange((r) => r.setFontWeight("bold")) },
            { label: "Italic", shortcut: "⌘I", onClick: () => withRange((r) => r.setFontStyle("italic")) },
            { label: "Underline", shortcut: "⌘U", onClick: () => withRange((r) => r.setFontLine("underline")) },
            { label: "Strikethrough", onClick: () => withRange((r) => r.setFontLine("line-through")) },
          ],
        },
        {
          label: "Alignment",
          items: [
            { label: "Left", onClick: () => withRange((r) => r.setHorizontalAlignment("left")) },
            { label: "Center", onClick: () => withRange((r) => r.setHorizontalAlignment("center")) },
            { label: "Right", onClick: () => withRange((r) => r.setHorizontalAlignment("right")) },
            { label: "Top", onClick: () => withRange((r) => r.setVerticalAlignment("top")), sep: true },
            { label: "Middle", onClick: () => withRange((r) => r.setVerticalAlignment("middle")) },
            { label: "Bottom", onClick: () => withRange((r) => r.setVerticalAlignment("bottom")) },
          ],
        },
        {
          label: "Wrapping",
          items: [
            { label: "Overflow", onClick: () => withRange((r) => r.setWrap(false)) },
            { label: "Wrap", onClick: () => withRange((r) => r.setWrap(true)) },
          ],
        },
        {
          label: "Rotation",
          items: [
            { label: "None", onClick: () => withRange((r) => r.setTextRotation(0)) },
            { label: "Tilt up (45°)", onClick: () => withRange((r) => r.setTextRotation(-45)) },
            { label: "Tilt down (45°)", onClick: () => withRange((r) => r.setTextRotation(45)) },
            { label: "Rotate up (90°)", onClick: () => withRange((r) => r.setTextRotation(-90)) },
            { label: "Rotate down (90°)", onClick: () => withRange((r) => r.setTextRotation(90)) },
          ],
        },
        {
          label: "Font size",
          items: FONT_SIZES.map((s) => ({ label: String(s), onClick: () => withRange((r) => r.setFontSize(s)) })),
        },
        {
          label: "Merge cells",
          items: [
            { label: "Merge all", onClick: () => withRange((r) => r.merge()) },
            { label: "Unmerge", onClick: () => withRange((r) => r.breakApart()) },
          ],
        },
        { label: "Convert to table", onClick: () => exec("sheet.command.add-table"), sep: true },
        { label: "Conditional formatting", onClick: () => exec("sheet.operation.open-conditional-formatting-panel") },
        { label: "Alternating colours", disabled: true },
        { label: "Clear formatting", shortcut: "⌘\\", onClick: () => exec("sheet.command.clear-selection-format"), sep: true },
      ],
    },
    {
      label: "Data",
      items: [
        { label: "Sort range A → Z", onClick: () => exec("sheet.command.sort-range-asc") },
        { label: "Sort range Z → A", onClick: () => exec("sheet.command.sort-range-desc") },
        { label: "Create a filter", onClick: () => exec("sheet.command.smart-toggle-filter"), sep: true },
        { label: "Data validation", onClick: () => apiOf()?.executeCommand("data-validation.operation.open-validation-panel", { isAdd: true }) },
      ],
    },
  ];

  // Track the live selection so "Freeze up to row/column" always reflects the
  // current cell. Reading it on menu-open is unreliable: clicking a menu label
  // can blur the sheet, making getActiveRange() return null (stale labels).
  useEffect(() => {
    const a = apiOf();
    if (!a?.addEvent) return;
    const sync = () => {
      const cell = a.getActiveWorkbook()?.getActiveRange();
      if (!cell) return;
      setSelRow(cell.getRow());
      setSelCol(cell.getColumn());
    };
    const ev = a.Event ?? {};
    const disposers = [a.addEvent(ev.SelectionChanged ?? "SelectionChanged", sync), a.addEvent(ev.SelectionMoveEnd ?? "SelectionMoveEnd", sync)];
    sync();
    return () => disposers.forEach((d) => d?.dispose?.());
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [api]);

  // Global keyboard shortcuts (cross-platform ⌘/Ctrl). A ref carries the latest
  // handlers so the once-attached listener never closes over a stale api.
  const kbdCtxRef = useRef<ShortcutContext>({ api });
  kbdCtxRef.current = {
    api,
    onFind: () => onOpenFind?.(),
    onImport: () => void importFile(),
    onPrint: doPrint,
    onSave: doSave, // ⌘S / Ctrl+S — same handler as File ▸ Save
  };
  useEffect(() => attachKeyboardShortcuts(() => kbdCtxRef.current), []);
  useEffect(() => () => { if (savedTimer.current) clearTimeout(savedTimer.current); }, []);

  const openAt = (idx: number, el: HTMLElement) => {
    const r = el.getBoundingClientRect();
    setPos({ top: r.bottom + 2, left: r.left });
    setOpenMenu(idx);
  };

  useEffect(() => {
    if (openMenu === null) return;
    const h = (ev: MouseEvent) => {
      const t = ev.target as HTMLElement;
      if (!t.closest("[data-levich-menu]") && !t.closest("[data-levich-menubar]")) setOpenMenu(null);
    };
    document.addEventListener("mousedown", h);
    return () => document.removeEventListener("mousedown", h);
  }, [openMenu]);

  return (
    <div ref={barRef} data-levich-menubar style={{ display: "flex", alignItems: "center", gap: 2, padding: "2px 10px", background: "#fff", borderBottom: "1px solid #f2f4f7" }}>
      {menus.map((m, i) => (
        <span
          key={m.label}
          style={{ ...menuLabel, background: openMenu === i ? "#fef3c7" : "transparent", color: openMenu === i ? "#101828" : menuLabel.color }}
          onClick={(e) => (openMenu === i ? setOpenMenu(null) : openAt(i, e.currentTarget))}
          onMouseEnter={(e) => openMenu !== null && openAt(i, e.currentTarget)}
        >
          {m.label}
        </span>
      ))}
      {openMenu !== null &&
        createPortal(
          <div data-levich-menu style={{ position: "fixed", top: pos.top, left: pos.left, ...panel }}>
            {menus[openMenu].items.map((it, i) => (
              <MenuItemRow key={i} item={it} onClose={() => setOpenMenu(null)} />
            ))}
          </div>,
          document.body,
        )}
      {importGrid &&
        createPortal(<ImportModal open fileName={importFileName} onCancel={() => setImportGrid(null)} onImport={applyImport} />, document.body)}
      {renameState &&
        createPortal(
          <RenameModal
            open
            title="Rename spreadsheet"
            current={renameState.current}
            taken={renameState.taken}
            onCancel={() => setRenameState(null)}
            onRename={applyRename}
          />,
          document.body,
        )}
      {saved &&
        createPortal(
          <div
            role="status"
            style={{
              position: "fixed", bottom: 24, left: "50%", transform: "translateX(-50%)",
              background: "#101828", color: "#fff", fontSize: 13, fontWeight: 500,
              padding: "8px 16px", borderRadius: 8, boxShadow: "0 4px 12px rgba(16,24,40,.18)",
              zIndex: 2147483647, pointerEvents: "none",
            }}
          >
            Saved ✓
          </div>,
          document.body,
        )}
    </div>
  );
}

export default LevichMenuBar;
