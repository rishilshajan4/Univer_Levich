/**
 * Public + internal types for @levich/univer-sheets.
 * Public types are re-exported from `src/index.ts` (the package contract);
 * internal types support the compile/export pipeline and are not part of the API.
 */
import type { UniverAPI } from "./create-sheet";
import type { ImportLocation } from "../components/import-modal";

/** Where an imported file's data is placed. Re-exported for host `onImport`. */
export type { ImportLocation };

/** The consumer's tabular data: an array of records keyed by field name. */
export type SheetData = Array<Record<string, unknown>>;

/**
 * How a column's values are displayed. Currency/number values stay numeric so
 * formulas + export remain correct. Date values are treated as display text
 * (pass ISO strings); they are not converted to spreadsheet date serials.
 */
export type ColumnFormat = "currency" | "date" | "number" | "text";

/** Describes one on-sheet column. Drives header, formatting, and edit behavior. */
export interface ColumnDef {
  /** Field read from each record. */
  key: string;
  /** Header label shown in row 0. */
  header: string;
  /** Display format. Defaults to "text". */
  format?: ColumnFormat;
  /** Read-only: edits to this column are vetoed. Wins over `editable`. */
  locked?: boolean;
  /** Explicitly editable (e.g. a comment column). */
  editable?: boolean;
  /** Column width in px. Falls back to the sheet default. */
  width?: number;
}

/** Consumer-driven freezing. `false` freezes nothing. Omitted defaults to 1 row. */
export type FreezeConfig = false | { rows?: number; columns?: number };

/** Aggregation used by the pivot engine. */
export type PivotAggregate = "sum" | "count" | "average" | "min" | "max";

/** Configuration-based pivot, computed from the supplied data. */
export interface PivotConfig {
  /** Field(s) grouped down the side. */
  rows: string[];
  /** Field(s) spread across the top. */
  columns: string[];
  /** Field(s) summarized. v1 aggregates the first value field only. */
  values: string[];
  /** Aggregation method. */
  aggregate: PivotAggregate;
}

/** Optional totals/footer row with a live SUM over one column. */
export interface FooterConfig {
  /** Footer label text (e.g. "TOTAL"). */
  label: string;
  /** Column key the label is placed in. Defaults to the column before the sum. */
  labelColumnKey?: string;
  /** Column key to SUM over the data rows. */
  sumColumnKey: string;
}

/** Emitted when an editable cell (e.g. a comment) is committed. */
export interface CellEditEvent {
  /** Stable identity of the row that was edited. */
  rowKey: string;
  row: number;
  column: number;
  value: string;
}

/** Public props for the `<LevichSheet>` component. */
export interface LevichSheetProps {
  data: SheetData;
  columns: ColumnDef[];
  /**
   * Render a raw Univer `IWorkbookData` snapshot directly, bypassing the
   * `data + columns` pipeline. Used for rich `.xlsx` import (full styles /
   * merges / number formats / multiple sheets). When set, `data`/`columns`/
   * `footer`/`pivot` are ignored. Produce one via `parseXlsxToSnapshot`.
   */
  snapshot?: WorkbookData;
  /** Freeze rows/columns. Omitted → freeze the top row. `false` → none. */
  freeze?: FreezeConfig;
  /** Optional pivot; renders a computed pivot region instead of the raw grid. */
  pivot?: PivotConfig;
  /** Optional totals row with a live SUM over one column. */
  footer?: FooterConfig;
  /** Currency symbol for currency columns. Defaults to "$". */
  currencySymbol?: string | null;
  /** Pre-fill for the editable comment column, keyed by row key. */
  comments?: Record<string, string>;
  /** Restore saved column widths, keyed by column index. */
  columnWidths?: Record<number, number>;
  /** Persistence hook fired when an editable cell is committed. */
  onCellEdit?: (event: CellEditEvent) => void;
  /** Fired (debounced) when column widths change. */
  onColumnWidthsChange?: (widths: Record<number, number>) => void;
  /** Toolbar style. Defaults to "simple". */
  toolbar?: "simple" | "full" | "none";
  /** Optional className for the host container. */
  className?: string;
  /** Stable row-key resolver for comment persistence. Defaults to row index. */
  getRowKey?: (record: Record<string, unknown>, index: number) => string;
  /**
   * Fired once after the sheet engine is ready, with the Univer Facade API.
   * Use it for advanced setup (e.g. seeding defined names for Defined-Name
   * hyperlinks, registering extra commands). Type-only import keeps the public
   * types runtime-engine-free.
   */
  onReady?: (api: UniverAPI) => void;
  /**
   * File ▸ Import hook. Receives the parsed grid + the chosen destination.
   * Return `true` if your app/backend handled it (e.g. FinOpz BE created or
   * replaced a document) — the built-in local behavior is then skipped. Return
   * falsy (or omit) to use the built-in behavior. Document-level destinations
   * ("new-spreadsheet" / "replace-spreadsheet") are the cases a real backend
   * typically owns.
   */
  onImport?: (grid: (string | number)[][], location: ImportLocation) => boolean | void;
  /**
   * File ▸ Save / ⌘S (Ctrl+S). Receives the LIVE Univer workbook snapshot so
   * the host can persist it to its backend. Return `true` if handled; return
   * falsy (or omit) to use the built-in fallback, which persists the snapshot
   * to `localStorage["levich:save:<workbookId>"]` and flashes a "Saved" toast.
   * Either way, the browser's own "Save page" dialog is always suppressed.
   */
  onSave?: (snapshot: WorkbookData) => boolean | void;
  /** File ▸ New — host hook for a fresh document. Defaults to clearing the sheet. */
  onNew?: () => void;
  /** File ▸ Make a copy — host hook. Defaults to downloading a `.xlsx` copy. */
  onMakeCopy?: () => void;
  /** File ▸ Rename — host hook for the document title. Defaults to renaming the active sheet. */
  onRename?: (name: string) => void;
  /**
   * Sheet tab ▸ Copy to ▸ Existing spreadsheet. Cross-spreadsheet copy is owned
   * by the host/backend (it must choose a target document). Receives the sheet
   * name + a single-sheet snapshot. Omitted → logs on localhost.
   */
  onCopyToExisting?: (sheetName: string, sheetSnapshot: unknown) => void;
}

/* -------------------------------------------------------------------------- */
/* Transaction preset                                                         */
/* -------------------------------------------------------------------------- */

/** One account-transaction row for the `TransactionSheet` preset. */
export interface TransactionRow {
  /** Stable GL transaction id (half of the comment row key). */
  transactionId: string;
  /** Stable per-line id (other half of the comment row key). */
  lineId: string;
  /** YYYY-MM-DD posting date. */
  date: string;
  /** Human-facing transaction number/id. */
  tranId: string;
  /** Transaction type (e.g. "Journal", "Invoice"). */
  type: string;
  /** Related entity / name. */
  entity: string;
  debit: number;
  credit: number;
  /** Signed net amount. */
  amount: number;
  memo: string;
  /** Editable comment; empty by default. */
  comment?: string;
}

/** Props for the `TransactionSheet` preset. */
export interface TransactionSheetProps {
  rows: TransactionRow[];
  /** Currency symbol for money columns. Defaults to "$". */
  currencySymbol?: string | null;
  /** Identifies the (subsidiary, account) bucket for localStorage persistence. */
  subsidiaryId?: string;
  accountId?: string;
  /** Pre-loaded comments (`${txnId}:${lineId}` → text). Overrides storage. */
  comments?: Record<string, string>;
  /** Pre-loaded column widths (columnIndex → px). Overrides storage. */
  columnWidths?: Record<number, number>;
  /** Optional className for the host container. */
  className?: string;
}

/** Imperative handle exposed via ref. */
export interface LevichSheetHandle {
  /** Export the LIVE sheet to a full-fidelity .xlsx. Resolves with rows written (0 if not ready). */
  exportXlsx: (fileName?: string) => Promise<number>;
}

/* -------------------------------------------------------------------------- */
/* Internal (compile pipeline)                                                */
/* -------------------------------------------------------------------------- */

/** Univer inline cell style object (short keys: bl, cl, bg, n.pattern, ht). */
export type CellStyle = Record<string, unknown>;

/** A Univer cell: literal value `v`, formula `f`, and/or inline style `s`. */
export type Cell = { v?: string | number; f?: string; s?: CellStyle } | null;

/** Loose alias for the Univer IWorkbookData snapshot we build. */
export type WorkbookData = Record<string, unknown>;

/** Result of compiling data + columns into a Univer workbook snapshot. */
export interface BuiltWorkbook {
  workbookData: WorkbookData;
  /** Number of data rows (excludes header + any footer). */
  rowCount: number;
  /** 0-based header row index (always 0). */
  headerRowIndex: number;
  /** 0-based footer/totals row index, when present. */
  footerRowIndex?: number;
  /** Number of columns occupied by the table. */
  columnCount: number;
}
