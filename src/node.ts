/**
 * `@levichco/finsheets/node` — the Node-safe (headless) entry point.
 *
 * Re-exports ONLY the pure-data converter / assembly functions — **no React, no
 * CSS, no Univer UI presets, no icons** — so a headless Node / tsx process (e.g.
 * the `finsheets-service` backend) can reuse the SAME `.xlsx` ↔ snapshot converter
 * the browser uses, without ever loading the front-end bundle.
 *
 * This is the "one isomorphic converter, wrapper-not-fork" design made real: the
 * main entry (`@levichco/finsheets`) is the React UI; this subpath is the Node
 * data core. Every module below has only type-only top-level imports (or `exceljs`,
 * which is Node-native + externalized), so `dist/node.js` loads cleanly in Node.
 *
 *   import { parseXlsxToSnapshot } from "@levichco/finsheets/node";
 */

// .xlsx → Univer snapshot (exceljs is dynamically imported; no DOM at load).
export { parseXlsxToSnapshot, sheetIndexToName, inlineSheetStyles, firstVisibleSheet } from "./core/xlsx-to-snapshot";
export type { ImportImage, ParsedSnapshot } from "./core/xlsx-to-snapshot";

// Interactive-pivot reconstruction from imported .xlsx (jszip/saxes are dynamically
// imported — both Node-native, transitive via exceljs — so this stays headless).
export { parsePivotsFromXlsx, mapSubtotal } from "./core/pivot-import";
export type { ImportedPivot } from "./core/pivot-import";

// Durable persistence of reconstructed pivots via a Univer `resources` entry that
// round-trips through save/load — so a host can re-open imported pivots on ANY
// open, not just the import session. Pure data; safe headless.
export { PIVOTS_IMPORT_RESOURCE, attachPivotsResource, readPivotsResource } from "./core/pivots-resource";

// Univer snapshot → ExcelJS workbook (Node-safe primitive). NOTE: `exportToXlsx`
// is intentionally NOT re-exported — it triggers a browser download (DOM). Use
// `buildExcelWorkbook(...).workbook.xlsx.writeBuffer()` to emit bytes headlessly.
export { buildExcelWorkbook, toArgb } from "./core/export-xlsx";
export type { WorkbookSnapshot, SnapshotSource } from "./core/export-xlsx";

// Lazy multi-sheet shell-workbook assembly (pure data).
export { buildShellWorkbook } from "./core/shell-workbook";
export type { SheetManifestEntry, SingleSheetSnapshot, BuildShellWorkbookParams } from "./core/shell-workbook";

// Version diff — "Highlight changes" engine (pure data; its version-store import is type-only).
export { diffSheet, highlightSnapshot } from "./features/version-diff";
export type { SheetDiff } from "./features/version-diff";

// Core snapshot types.
export type { WorkbookData, Cell, CellStyle } from "./core/types";
