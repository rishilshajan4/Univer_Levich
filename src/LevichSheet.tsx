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
import { createSheet, type UniverAPI } from "./core/create-sheet";
import { FindReplaceModal } from "./components/find-replace-modal";
import { LevichMenuBar } from "./menu/levich-menu-bar";
import { LevichToolbar } from "./toolbar/levich-toolbar";
import type { Disposer } from "./core/facade";
import { exportToXlsx, type SnapshotSource } from "./core/export-xlsx";
import { attachColumnWidths } from "./features/column-widths";
import { attachComments } from "./features/comments";
import { attachFilterPanel } from "./features/filter-panel";
import { attachLockColumns } from "./features/lock-columns";
import { buildPivotCells, computePivot } from "./features/pivot";
import { SheetTabMenu } from "./features/sheet-tab-menu";
import type { LevichSheetHandle, LevichSheetProps } from "./core/types";

function ribbonFor(toolbar: LevichSheetProps["toolbar"]): "collapsed" | "simple" | "classic" {
  if (toolbar === "full") return "classic";
  if (toolbar === "none") return "collapsed";
  return "simple";
}

export const LevichSheet = forwardRef<LevichSheetHandle, LevichSheetProps>(function LevichSheet(props, ref) {
  const { data, columns, snapshot, freeze, pivot, footer, currencySymbol, comments, columnWidths, getRowKey, toolbar, className, onCellEdit, onColumnWidthsChange, onReady, onImport, onSave, onNew, onMakeCopy, onRename, onCopyToExisting } = props;

  const containerRef = useRef<HTMLDivElement | null>(null);
  const univerRef = useRef<{ dispose: () => void } | null>(null);
  const apiRef = useRef<UniverAPI | null>(null);
  const [toolbarApi, setToolbarApi] = useState<UniverAPI | null>(null);
  const [findOpen, setFindOpen] = useState(false);

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
    });
    univerRef.current = univer;
    apiRef.current = univerAPI;
    setToolbarApi(univerAPI);
    onReady?.(univerAPI);

    // --- Configurable behaviors (all opt-in, driven by the column config) ----
    const disposers: Disposer[] = [];
    if (!pivot && !snapshot) {
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
    if (!pivot && !snapshot) {
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

    // Rich-import images are embedded in the snapshot as the SHEET_DRAWING_PLUGIN
    // resource, so Univer renders them at load — no post-load work needed here.

    return () => {
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
  }, [data, columns, snapshot, freeze, pivot, footer, currencySymbol, comments, columnWidths, getRowKey, toolbar, onCellEdit, onColumnWidthsChange]);

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

  return (
    <div className={className ?? "levich-sheet"} style={{ display: "flex", flexDirection: "column", width: "100%", height: "100%", minHeight: 0 }}>
      <LevichMenuBar api={toolbarApi} onOpenFind={() => setFindOpen(true)} onImport={onImport} onSave={onSave} onNew={onNew} onMakeCopy={onMakeCopy} onRename={onRename} />
      <LevichToolbar api={toolbarApi} onOpenFind={() => setFindOpen(true)} />
      <div style={{ position: "relative", flex: 1, minHeight: 0 }}>
        <div ref={containerRef} style={{ position: "absolute", inset: 0 }} />
        <FindReplaceModal api={toolbarApi} open={findOpen} onClose={() => setFindOpen(false)} />
        <SheetTabMenu api={toolbarApi} onCopyToExisting={onCopyToExisting} />
      </div>
    </div>
  );
});

export default LevichSheet;
