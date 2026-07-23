/**
 * <LevichSheet> — the package's single public component.
 *
 * Compiles `data` + `columns` (+ layout config) into a Univer workbook, renders
 * it fully Levich-branded, and (in later stages) attaches configurable behaviors
 * and full-fidelity export. All `@univerjs` imports stay in `core/create-sheet`
 * so this component is `React.lazy`-friendly.
 */
import { forwardRef, useEffect, useImperativeHandle, useRef, useState } from "react";
import { buildWorkbook } from "./core/build-workbook";
import { createSheet, forceCanvasResize, type UniverAPI } from "./core/create-sheet";
import { FindReplaceModal } from "./components/find-replace-modal";
import { LevichMenuBar } from "./menu/levich-menu-bar";
import { LevichToolbar } from "./toolbar/levich-toolbar";
import type { Disposer } from "./core/facade";
import { exportToXlsx, type SnapshotSource } from "./core/export-xlsx";
import { ensureFontsForSnapshot } from "./theme/google-fonts";
import { attachColumnWidths } from "./features/column-widths";
import { attachComments } from "./features/comments";
import { attachFilterPanel } from "./features/filter-panel";
import { attachLockColumns } from "./features/lock-columns";
import { buildPivotCells, computePivot } from "./features/pivot";
import { computePivotModel, renderPivotModel } from "./features/pivot-model";
import { PivotPanel } from "./features/pivot-panel";
import { SheetTabMenu } from "./features/sheet-tab-menu";
import { emptyFormulaCells, findHashCell } from "./core/snapshot-scan";
import type { Cell, LevichSheetHandle, LevichSheetProps, PivotSource, PivotSpec } from "./core/types";

function ribbonFor(toolbar: LevichSheetProps["toolbar"]): "collapsed" | "simple" | "classic" {
  if (toolbar === "full") return "classic";
  if (toolbar === "none") return "collapsed";
  return "simple";
}

/**
 * A sensible starting pivot layout when the host doesn't supply one: the first
 * text-like field becomes the row grouping, and the first numeric field is
 * summed. Falls back to counting the first field when nothing looks numeric.
 */
function defaultPivotSpec(_source: PivotSource): PivotSpec {
  // Google-Sheets behavior: a freshly-inserted pivot is EMPTY — no rows/columns/values.
  // The user configures it from the editor; nothing renders until they add a field.
  return { rows: [], columns: [], values: [] };
}

/**
 * Map the ABSOLUTE sheet row of every collapsible group-label cell → that node's
 * collapse `path`. Mirrors `renderPivotModel`'s row walk exactly so clicking a
 * ▸/▾ label can toggle the right node. Only nodes WITH children are collapsible.
 */
function collapsibleRowPaths(model: ReturnType<typeof computePivotModel>): Map<number, string> {
  const { spec } = model;
  const collapsed = new Set(spec.collapsed ?? []);
  const showRowSubtotals = spec.showRowSubtotals ?? spec.rows.length > 1;
  const colDepth = spec.columns.length;
  const headerRows = (colDepth > 0 ? 1 : 0) + 1; // column-group line (opt) + value-label line
  const out = new Map<number, string>();
  let r = headerRows;
  const walk = (nodes: typeof model.rowTree) => {
    for (const node of nodes) {
      const hasChildren = node.children.length > 0;
      const isCollapsed = collapsed.has(node.path);
      if (hasChildren) out.set(r, node.path);
      r++;
      if (hasChildren && !isCollapsed) {
        walk(node.children);
        if (showRowSubtotals) r++; // subtotal row
      }
    }
  };
  walk(model.rowTree);
  return out;
}

export const LevichSheet = forwardRef<LevichSheetHandle, LevichSheetProps>(function LevichSheet(props, ref) {
  const { data, columns, snapshot, anchorCell, freeze, pivot, pivotInteractive, footer, currencySymbol, comments, columnWidths, getRowKey, toolbar, sheetBar, readOnly, className, onCellEdit, onColumnWidthsChange, onReady, onImport, onImportFile, onSave, onDownload, onNew, onMakeCopy, onRename, onCopyToExisting, onHideActiveSheet, onShowSheet, hiddenSheetList, canHideActiveSheet, onInsertPivot } = props;

  const containerRef = useRef<HTMLDivElement | null>(null);
  const univerRef = useRef<{ dispose: () => void } | null>(null);
  const apiRef = useRef<UniverAPI | null>(null);
  const [toolbarApi, setToolbarApi] = useState<UniverAPI | null>(null);
  const [findOpen, setFindOpen] = useState(false);

  // ── Interactive pivot state ────────────────────────────────────────────────
  // `spec` is the live pivot layout (driven by the drawer). It is applied to the
  // grid via IN-PLACE Facade writes (no remount → no flicker), so it must NOT be
  // in the build effect's deps. `specRef` mirrors it for the collapse click
  // handler + the in-place writer; a ref of the last-rendered rectangle lets us
  // clear stale cells before each redraw.
  const [pivotSpec, setPivotSpec] = useState<PivotSpec | null>(() => (pivotInteractive ? pivotInteractive.initialSpec ?? defaultPivotSpec(pivotInteractive.source) : null));
  const [pivotOpen, setPivotOpen] = useState<boolean>(!!pivotInteractive);
  const [addRowsN, setAddRowsN] = useState<number>(1000); // "Add N more rows at the bottom" input
  const specRef = useRef<PivotSpec | null>(pivotSpec);
  specRef.current = pivotSpec;
  const lastPivotRectRef = useRef<{ rows: number; cols: number }>({ rows: 0, cols: 0 });
  // True until the first in-place write, so the initial render (done by the build
  // effect) isn't redundantly re-written.
  const pivotBuiltRef = useRef(false);

  useEffect(() => {
    const container = containerRef.current;
    if (!container) return;

    const rowKeys = data.map((record, i) => (getRowKey ? getRowKey(record, i) : String(i)));

    // Pivot mode renders a computed pivot region instead of the raw grid.
    let workbookData;
    let behaviorColumnCount = columns.length;
    let footerRowIndex: number | undefined;
    if (snapshot) {
      // Rich-import mode: render the pre-built snapshot verbatim (styles,
      // merges, formats, multiple sheets). Skip the data/columns pipeline.
      workbookData = snapshot;
      const firstId = (snapshot.sheetOrder as string[] | undefined)?.[0];
      const firstSheet = firstId ? (snapshot.sheets as Record<string, { columnCount?: number }> | undefined)?.[firstId] : undefined;
      behaviorColumnCount = firstSheet?.columnCount ?? 26;
    } else if (pivotInteractive) {
      // Interactive pivot: render the current spec's region into a fresh sheet.
      // Subsequent spec changes are written IN PLACE (see the pivot-apply effect
      // below) — this build only lays down the initial grid + freeze geometry.
      const spec = specRef.current ?? defaultPivotSpec(pivotInteractive.source);
      const region = renderPivotModel(computePivotModel(pivotInteractive.source, spec));
      behaviorColumnCount = Math.max(region.columnCount, 4);
      lastPivotRectRef.current = { rows: region.rowCount, cols: region.columnCount };
      pivotBuiltRef.current = true;
      workbookData = buildWorkbook([], [], {
        extraCells: region.cells,
        extraRows: region.rowCount,
        extraColumns: region.columnCount,
        freeze: freeze ?? { rows: 1 },
        columnWidths,
      }).workbookData;
    } else if (pivot) {
      const result = computePivot(data, pivot);
      const region = buildPivotCells(result);
      behaviorColumnCount = region.columnCount;
      workbookData = buildWorkbook([], [], {
        extraCells: region.cells,
        extraRows: region.rowCount,
        extraColumns: region.columnCount,
        freeze: freeze ?? { rows: 1 },
        columnWidths,
      }).workbookData;
    } else {
      const commentColumnKey = columns.find((c) => c.editable)?.key;
      const built = buildWorkbook(data, columns, {
        freeze,
        currencySymbol,
        comments,
        columnWidths,
        rowKeys,
        commentColumnKey,
        footer,
      });
      workbookData = built.workbookData;
      footerRowIndex = built.footerRowIndex;
    }

    const { univer, univerAPI } = createSheet({
      container,
      workbookData,
      ribbonType: ribbonFor(toolbar),
      univerToolbar: false, // hide Univer's toolbar; we render the Levich toolbar
      // sheetBar:false → hide Univer's native footer tabs; the host renders its own
      // <SheetTabBar> (sole controller of the active sheet in shell-workbook mode).
      footer: sheetBar === false ? false : undefined,
    });
    univerRef.current = univer;
    apiRef.current = univerAPI;
    setToolbarApi(univerAPI);
    onReady?.(univerAPI);

    // Blank-grid fix: keep Univer's canvas sized to its container by re-measuring
    // DIRECTLY (forceCanvasResize → engine.resize()), bypassing Univer's own
    // ResizeObserver → requestIdleCallback recovery which can be starved in a busy
    // app and leave the canvas stuck at the 0×0 it measured at build time. Force a
    // re-measure now, across the next few frames (covers the box settling one frame
    // late), and on every genuine container size change. engine.resize() early-returns
    // on unchanged size, so this is cheap.
    let rafId = 0;
    const nudge = () => {
      cancelAnimationFrame(rafId);
      rafId = requestAnimationFrame(() => forceCanvasResize(univerAPI));
    };
    const sizeObserver = typeof ResizeObserver !== "undefined" ? new ResizeObserver(nudge) : null;
    sizeObserver?.observe(container);
    nudge();
    const nudgeTimers = [50, 160, 400, 800].map((ms) => setTimeout(nudge, ms));

    // Load any Google Fonts the snapshot uses (Univer paints on a canvas, so a
    // font must be in the FontFaceSet or text falls back to serif). Once loaded,
    // recompute the render skeleton so text re-measures in the right typeface
    // (refreshCanvas → SheetSkeletonManagerService.reCalculate, size-INDEPENDENT),
    // AND re-size the canvas in case it was built at 0×0 (forceCanvasResize →
    // engine.resize, which is size-GATED). These address different failure modes —
    // stale text metrics vs a 0×0 canvas — so run BOTH, not one instead of the other.
    if (snapshot) {
      void ensureFontsForSnapshot(snapshot).then(() => {
        try { (univerAPI.getActiveWorkbook()?.getActiveSheet() as unknown as { refreshCanvas?: () => void })?.refreshCanvas?.(); }
        catch { /* facade surface differs — best-effort */ }
        forceCanvasResize(univerAPI);
      });
    }

    // --- Configurable behaviors (all opt-in, driven by the column config) ----
    const disposers: Disposer[] = [];

    // Read-only preview: veto every edit at start AND end (paste/fill paths),
    // and try the workbook permission API as a belt-and-suspenders. Formatting/
    // navigation stay allowed — only value entry is blocked. Best-effort.
    if (readOnly) {
      try {
        const f = univerAPI as unknown as {
          Event?: Record<string, string>;
          addEvent?: (event: string, cb: (p: { cancel?: boolean }) => void) => Disposer;
        };
        const veto = (p: { cancel?: boolean }) => { p.cancel = true; };
        const startEvent = f.Event?.BeforeSheetEditStart;
        const endEvent = f.Event?.BeforeSheetEditEnd;
        if (startEvent && f.addEvent) disposers.push(f.addEvent(startEvent, veto));
        if (endEvent && f.addEvent) disposers.push(f.addEvent(endEvent, veto));
      } catch { /* event surface differs — veto is best-effort */ }
      try {
        (univerAPI.getActiveWorkbook() as unknown as { setEditable?: (v: boolean) => void })?.setEditable?.(false);
      } catch { /* permission API differs — best-effort */ }
    }

    if (!pivot && !pivotInteractive && !snapshot) {
      const lockedColumns = columns.flatMap((c, i) => (c.locked ? [i] : []));
      const editableColumn = columns.findIndex((c) => c.editable);
      const rowKeyByIndex = new Map<number, string>();
      rowKeys.forEach((key, i) => rowKeyByIndex.set(i + 1, key));
      disposers.push(
        ...attachLockColumns(univerAPI, { lockedColumns, rowCount: data.length, footerRowIndex }),
        ...attachComments(univerAPI, { editableColumn, rowKeyByIndex, onCellEdit }),
      );
    }
    disposers.push(...attachColumnWidths(univerAPI, { columnCount: behaviorColumnCount, onColumnWidthsChange }));
    // Replace Univer's header-funnel dropdown with our Google-style filter menu
    // (drives the same filter engine via the public Facade).
    disposers.push(...attachFilterPanel(univerAPI));

    // Open with the first DATA cell (A2) selected rather than Univer's default
    // A1 — A1 is the bold header, which made the toolbar's B/I/U/S show
    // "pressed" on every load. Done at the Steady (3) lifecycle so it isn't
    // overwritten by the engine's initial A1 selection.
    if (!pivot && !pivotInteractive && !snapshot) {
      try {
        const lifeEvent = (univerAPI as unknown as { Event?: Record<string, string> }).Event?.LifeCycleChanged;
        if (lifeEvent) {
          const d = (univerAPI as unknown as { addEvent?: (e: string, cb: (p: { stage?: number }) => void) => Disposer }).addEvent?.(lifeEvent, (p) => {
            if (p?.stage === 3) {
              try {
                (univerAPI.getActiveWorkbook() as unknown as { getActiveSheet?: () => { getRange?: (r: number, c: number) => { activate?: () => void } | undefined } | undefined })?.getActiveSheet?.()?.getRange?.(1, 0)?.activate?.();
              } catch {
                /* selection set is best-effort */
              }
            }
          });
          if (d) disposers.push(d);
        }
      } catch {
        /* lifecycle surface differs — best-effort */
      }
    }

    // Interactive-pivot collapse/expand: clicking a group-label cell that carries
    // a ▸/▾ chevron toggles that node's `path` in spec.collapsed (via specRef, so
    // this handler doesn't need to be re-registered per spec change). Best-effort
    // through the Facade selection event; the panel remains the source of truth.
    if (pivotInteractive) {
      try {
        const f = univerAPI as unknown as { Event?: Record<string, string>; addEvent?: (e: string, cb: (p?: unknown) => void) => Disposer };
        const ev = f.Event ?? {};
        const onClick = () => {
          const spec = specRef.current;
          if (!spec) return;
          try {
            const wb = univerAPI.getActiveWorkbook() as unknown as {
              getActiveRange?: () => { getRow?: () => number; getColumn?: () => number } | null;
            };
            const range = wb?.getActiveRange?.();
            const row = range?.getRow?.();
            const col = range?.getColumn?.();
            if (col !== 0 || row == null) return; // only the row-label column (col 0)
            const model = computePivotModel(pivotInteractive.source, spec);
            const path = collapsibleRowPaths(model).get(row);
            if (!path) return;
            const collapsed = new Set(spec.collapsed ?? []);
            if (collapsed.has(path)) collapsed.delete(path);
            else collapsed.add(path);
            setPivotSpec({ ...spec, collapsed: [...collapsed] });
          } catch {
            /* click resolution is best-effort */
          }
        };
        if (f.addEvent) {
          const d1 = f.addEvent(ev.SelectionMoveEnd ?? "SelectionMoveEnd", onClick);
          const d2 = f.addEvent(ev.SelectionChanged ?? "SelectionChanged", onClick);
          if (d1) disposers.push(d1);
          if (d2) disposers.push(d2);
        }
      } catch {
        /* selection surface differs — collapse falls back to a no-op */
      }
    }

    // Snapshot / rich-import path (host editor + xlsx import): recompute any truly-empty
    // formula cells (companion to NO_CALCULATION — feature #12) and open at the anchor /
    // "#" cell (feature #2). Both run at the Steady (3) lifecycle so they aren't
    // overwritten by the engine's initial A1 selection.
    if (snapshot) {
      const target = anchorCell ?? findHashCell(snapshot);
      const empties = emptyFormulaCells(snapshot);
      try {
        const lifeEvent = (univerAPI as unknown as { Event?: Record<string, string> }).Event?.LifeCycleChanged;
        if (lifeEvent) {
          const d = (univerAPI as unknown as { addEvent?: (e: string, cb: (p: { stage?: number }) => void) => Disposer }).addEvent?.(lifeEvent, (p) => {
            if (p?.stage !== 3) return;
            try {
              const sheet = (univerAPI.getActiveWorkbook() as unknown as {
                getActiveSheet?: () => {
                  getRange?: (r: number, c: number) => { activate?: () => void; setValue?: (v: unknown) => void } | undefined;
                  scrollToCell?: (r: number, c: number) => void;
                } | undefined;
              })?.getActiveSheet?.();
              if (!sheet?.getRange) return;
              // #2: jump to + select the anchor / "#" cell FIRST (else the first data
              // cell A2) so the view settles immediately, before the recompute below.
              const anchor = target ?? { row: 1, column: 0 };
              sheet.getRange(anchor.row, anchor.column)?.activate?.();
              try { sheet.scrollToCell?.(anchor.row, anchor.column); } catch { /* scroll best-effort */ }
              // #12: fill truly-empty formula cells by re-applying their formula (a
              // targeted recompute; cached-value cells — incl. genuine zeros — are never
              // touched). Written in CHUNKS that yield to the event loop between batches:
              // a tight synchronous loop over thousands of empty-formula cells (every
              // total blank under NO_CALCULATION) dispatched one Facade command each and
              // froze the tab on open of a large workbook. Chunking keeps it responsive
              // (totals fill in progressively) instead of a multi-second hang.
              if (empties.length) {
                const CHUNK = 200;
                let i = 0;
                const schedule = typeof requestAnimationFrame === "function" ? requestAnimationFrame : (cb: () => void) => setTimeout(cb, 0);
                const writeChunk = () => {
                  if (!apiRef.current) return; // component disposed mid-recompute — stop
                  const end = Math.min(i + CHUNK, empties.length);
                  for (; i < end; i++) {
                    const { row, column, formula } = empties[i];
                    try { sheet.getRange?.(row, column)?.setValue?.({ f: formula }); } catch { /* per-cell best-effort */ }
                  }
                  if (i < empties.length) schedule(writeChunk);
                };
                writeChunk();
              }
            } catch {
              /* post-load pass is best-effort */
            }
          });
          if (d) disposers.push(d);
        }
      } catch {
        /* lifecycle surface differs — best-effort */
      }
    }

    // Rich-import images are embedded in the snapshot as the SHEET_DRAWING_PLUGIN
    // resource, so Univer renders them at load — no post-load work needed here.

    return () => {
      sizeObserver?.disconnect();
      cancelAnimationFrame(rafId);
      nudgeTimers.forEach(clearTimeout);
      disposers.forEach((d) => {
        try {
          d.dispose();
        } catch {
          /* ignore */
        }
      });
      try {
        univer.dispose();
      } catch {
        /* ignore disposal errors */
      }
      univerRef.current = null;
      apiRef.current = null;
      setToolbarApi(null);
    };
    // The component is remounted per dataset via `key` upstream, so this is
    // effectively mount-once; deps cover the rebuild-on-change case.
  }, [data, columns, snapshot, anchorCell, freeze, pivot, pivotInteractive, footer, currencySymbol, comments, columnWidths, getRowKey, toolbar, readOnly, onCellEdit, onColumnWidthsChange]);

  // Interactive pivot: apply spec changes IN PLACE via Facade `setValue`, with no
  // remount → no flicker. The build effect lays down the FIRST render (guarded by
  // pivotBuiltRef so this effect skips that same paint). Each redraw clears the
  // previous pivot rectangle before writing the new cells so stale rows/columns
  // from a larger prior layout never linger.
  useEffect(() => {
    if (!pivotInteractive || !pivotSpec) return;
    if (pivotBuiltRef.current) {
      // The build effect just painted this exact spec — don't double-write.
      pivotBuiltRef.current = false;
      return;
    }
    const api = apiRef.current;
    if (!api) return;
    try {
      const sheet = (api.getActiveWorkbook() as unknown as {
        getActiveSheet?: () => {
          getRange?: (r: number, c: number, numRows?: number, numColumns?: number) => { setValue?: (v: unknown) => void; setValues?: (v: unknown) => unknown } | undefined;
          getMaxRows?: () => number;
          getMaxColumns?: () => number;
          insertRows?: (rowIndex: number, numRows: number) => void;
          insertColumns?: (columnIndex: number, numColumns: number) => void;
          setRowCount?: (count: number) => void;
          setColumnCount?: (count: number) => void;
        } | undefined;
      })?.getActiveSheet?.();
      if (!sheet?.getRange) return;

      const region = renderPivotModel(computePivotModel(pivotInteractive.source, pivotSpec));
      const prev = lastPivotRectRef.current;
      const clearRows = Math.max(prev.rows, region.rowCount);
      const clearCols = Math.max(prev.cols, region.columnCount);

      // Grow the pivot sheet to fit BEFORE writing. The host pre-sizes the sheet to the
      // source's cardinality (a hard upper bound — a pivot can't have more rows/cols than
      // the source has rows), so in practice the region ALWAYS fits and this is a no-op.
      // This block is a belt-and-suspenders backstop only.
      //
      // IMPORTANT: use `setRowCount`/`setColumnCount` (a worksheet-config mutation) — NEVER
      // `insertRows`/`insertColumns`. The pivot output sheet is read-only (users can't edit
      // pivot cells), and the Facade's `insertRows` runs through Univer's permission plugin,
      // which pops a BLOCKING modal alert ("The range is protected…") on a locked sheet.
      // `setRowCount` bypasses that permission check, so growth (if ever needed) is silent.
      const growN = (max: number, need: number, setCount?: (n: number) => void) => {
        if (max >= need) return;
        if (typeof setCount === "function") setCount(need + 8);
      };
      try {
        growN(sheet.getMaxRows?.() ?? 0, clearRows, sheet.setRowCount?.bind(sheet));
        growN(sheet.getMaxColumns?.() ?? 0, clearCols, sheet.setColumnCount?.bind(sheet));
      } catch {
        /* best-effort growth; getRange below may still clamp */
      }

      // Write the union of the old + new rectangle in a SINGLE bulk `setValues`
      // command (one recalc/render) instead of one Facade `setValue` per cell.
      // The per-cell loop dispatched clearRows×clearCols synchronous commands and
      // froze the tab on a large pivot / every spec change. Cells outside the new
      // region are written as blank ({ v:"", s:null }) so a shrinking layout leaves
      // no stale rows/columns behind.
      if (clearRows > 0 && clearCols > 0) {
        const matrix: unknown[][] = new Array(clearRows);
        for (let r = 0; r < clearRows; r++) {
          const rowCells = region.cells[r] ?? {};
          const row: unknown[] = new Array(clearCols);
          for (let c = 0; c < clearCols; c++) row[c] = (rowCells[c] as Cell | undefined) ?? { v: "", s: null };
          matrix[r] = row;
        }
        const range = sheet.getRange(0, 0, clearRows, clearCols);
        if (range?.setValues) range.setValues(matrix);
        else
          // Fallback for a Facade without bulk setValues: per-cell (rare/old build).
          for (let r = 0; r < clearRows; r++) for (let c = 0; c < clearCols; c++) { try { sheet.getRange(r, c)?.setValue?.((matrix[r] as unknown[])[c]); } catch { /* best-effort */ } }
      }
      lastPivotRectRef.current = { rows: region.rowCount, cols: region.columnCount };
    } catch {
      /* in-place apply is best-effort; the panel state stays authoritative */
    }
  }, [pivotSpec, pivotInteractive]);

  useImperativeHandle(
    ref,
    () => ({
      exportXlsx: async (fileName?: string) => {
        const api = apiRef.current;
        if (!api) return 0;
        const workbook = api.getActiveWorkbook();
        return exportToXlsx(workbook as unknown as SnapshotSource | null, fileName ?? "sheet.xlsx");
      },
    }),
    [],
  );

  // "Add N more rows at the bottom" — Google-Sheets-style row growth. Routes through the
  // Facade `insertRows` (which has NO count clamp, unlike Univer's native right-click dialog
  // that caps large inserts), so adding 10,000 rows adds exactly 10,000.
  const addRowsAtBottom = () => {
    const n = Math.max(1, Math.floor(Number(addRowsN)) || 0);
    try {
      const sheet = apiRef.current?.getActiveWorkbook()?.getActiveSheet() as unknown as
        | {
            getMaxRows?: () => number;
            insertRows?: (rowIndex: number, numRows: number) => void;
            insertRowsAfter?: (afterPosition: number, howMany: number) => void;
            setRowCount?: (count: number) => void;
          }
        | undefined;
      if (!sheet) return;
      const max = sheet.getMaxRows?.() ?? 0;
      // The Univer Facade exposes different row-growth methods across versions; try each
      // (none clamp the count, unlike the native right-click dialog that caps large inserts)
      // and confirm growth so the button never silently no-ops.
      if (typeof sheet.insertRows === "function") sheet.insertRows(max, n);
      else if (typeof sheet.insertRowsAfter === "function") sheet.insertRowsAfter(Math.max(0, max - 1), n);
      else if (typeof sheet.setRowCount === "function") sheet.setRowCount(max + n);
      if ((sheet.getMaxRows?.() ?? max) <= max && typeof sheet.setRowCount === "function") sheet.setRowCount(max + n);
    } catch {
      /* best-effort */
    }
  };
  const showAddRows = !readOnly && !pivotInteractive && !pivot;

  return (
    <div className={className ?? "levich-sheet"} style={{ display: "flex", flexDirection: "column", width: "100%", height: "100%", minHeight: 0 }}>
      <LevichMenuBar api={toolbarApi} onOpenFind={() => setFindOpen(true)} onImport={onImport} onImportFile={onImportFile} onSave={onSave} onDownload={onDownload} onNew={onNew} onMakeCopy={onMakeCopy} onRename={onRename} onHideActiveSheet={onHideActiveSheet} onShowSheet={onShowSheet} hiddenSheetList={hiddenSheetList} canHideActiveSheet={canHideActiveSheet} onInsertPivot={onInsertPivot} />
      <LevichToolbar api={toolbarApi} onOpenFind={() => setFindOpen(true)} />
      <div style={{ position: "relative", flex: 1, minHeight: 0 }}>
        <div ref={containerRef} style={{ position: "absolute", inset: 0 }} />
        <FindReplaceModal api={toolbarApi} open={findOpen} onClose={() => setFindOpen(false)} />
        {/* Injected-caret tab menu only for the NATIVE footer. When the host hides
            the native bar (sheetBar:false) it renders its own <SheetTabBar>. */}
        {sheetBar !== false && <SheetTabMenu api={toolbarApi} onCopyToExisting={onCopyToExisting} />}
        {/* Interactive-pivot fields drawer + a floating toggle when it's closed. */}
        {pivotInteractive && pivotSpec && pivotOpen && (
          <PivotPanel
            fields={pivotInteractive.source.fields}
            spec={pivotSpec}
            distinctValues={(field) => {
              const seen = new Set<string>();
              for (const row of pivotInteractive.source.rows) {
                const v = row[field];
                seen.add(v == null ? "" : String(v));
              }
              return [...seen].sort((a, b) => a.localeCompare(b, undefined, { numeric: true }));
            }}
            onChange={(next) => {
              setPivotSpec(next);
              pivotInteractive.onSpecChange?.(next); // let the host persist the layout
            }}
            onClose={() => setPivotOpen(false)}
          />
        )}
        {pivotInteractive && !pivotOpen && (
          <button
            type="button"
            onClick={() => setPivotOpen(true)}
            aria-label="Show PivotTable fields"
            style={{
              position: "absolute",
              top: 12,
              right: 12,
              zIndex: 50,
              height: 34,
              padding: "0 14px",
              borderRadius: 8,
              border: "1px solid #eaecf0",
              background: "#fff",
              color: "#344054",
              fontSize: 13,
              fontWeight: 600,
              cursor: "pointer",
              boxShadow: "0 2px 8px rgba(16,24,40,0.10)",
            }}
          >
            PivotTable Fields
          </button>
        )}
      </div>
      {showAddRows && (
        <div
          style={{
            display: "flex",
            alignItems: "center",
            gap: 8,
            padding: "6px 14px",
            borderTop: "1px solid #eaecf0",
            background: "#f9fafb",
            fontSize: 13,
            color: "#475467",
            flexShrink: 0,
          }}
        >
          <button
            type="button"
            onClick={addRowsAtBottom}
            style={{ color: "#155eef", fontWeight: 600, background: "none", border: "none", cursor: "pointer", padding: 0 }}
          >
            Add
          </button>
          <input
            type="number"
            min={1}
            value={addRowsN}
            onChange={(e) => setAddRowsN(Math.max(1, Math.floor(Number(e.target.value)) || 1))}
            aria-label="Number of rows to add"
            style={{ width: 72, height: 28, borderRadius: 6, border: "1px solid #d0d5dd", padding: "0 8px", fontSize: 13, color: "#101828", boxSizing: "border-box" }}
          />
          more rows at the bottom
        </div>
      )}
    </div>
  );
});

export default LevichSheet;
