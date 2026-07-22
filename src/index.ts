// @levich/univer-sheets — public API surface (the package contract).
// Stylesheet: import "@levich/univer-sheets/styles.css" once in the host.

export { LevichSheet, default as default } from "./LevichSheet";
export { TransactionSheet } from "./presets/transaction-sheet";
export { exportToXlsx } from "./core/export-xlsx";
export { parseXlsxToSnapshot } from "./core/xlsx-to-snapshot";
export type { ParsedSnapshot } from "./core/xlsx-to-snapshot";
export { parsePivotsFromXlsx, mapSubtotal } from "./core/pivot-import";
export type { ImportedPivot } from "./core/pivot-import";
export { buildShellWorkbook } from "./core/shell-workbook";
export { GOOGLE_FONTS, ensureGoogleFont, ensureGoogleFonts, ensureFontsForSnapshot, fontsInSnapshot } from "./theme/google-fonts";
export type { SheetManifestEntry, SingleSheetSnapshot, BuildShellWorkbookParams } from "./core/shell-workbook";
export { parseFileToGrid, stashImportPayload, takeImportPayload, stashSnapshotPayload, takeSnapshotPayload } from "./core/import-data";
export { LevichMenuBar } from "./menu/levich-menu-bar";
export { LevichToolbar } from "./toolbar/levich-toolbar";
export { Modal, ConfirmModal, Button } from "./components/modal";
export { SheetTabBar } from "./components/sheet-tab-bar";
export type { SheetTabInfo, SheetTabBarProps } from "./components/sheet-tab-bar";
export { VersionHistoryDrawer } from "./components/version-history-drawer";
export type { VersionHistoryDrawerProps } from "./components/version-history-drawer";
export { createVersionStore } from "./features/version-store";
export type { Version, VersionKind, VersionStore, DocumentSnapshot } from "./features/version-store";
export { diffSheet, highlightSnapshot } from "./features/version-diff";
export type { SheetDiff } from "./features/version-diff";
export { RenameModal } from "./components/rename-modal";
export type { RenameModalProps } from "./components/rename-modal";
export { FindReplaceModal } from "./components/find-replace-modal";
export { PivotPanel } from "./features/pivot-panel";
export type { PivotPanelProps, PivotArea } from "./features/pivot-panel";
export { computePivotModel, renderPivotModel, valueLabel } from "./features/pivot-model";
export { LEVICH_BRAND } from "./theme/brand";

export type {
  SheetData,
  ColumnDef,
  ColumnFormat,
  FreezeConfig,
  FooterConfig,
  PivotConfig,
  PivotAggregate,
  PivotSource,
  PivotSpec,
  PivotValueField,
  PivotFilterField,
  PivotModel,
  PivotNode,
  CellEditEvent,
  ImportLocation,
  LevichSheetProps,
  LevichSheetHandle,
  TransactionRow,
  TransactionSheetProps,
} from "./core/types";
