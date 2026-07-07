# Changelog

All notable changes to `@levichco/finsheets` (formerly `@levich/univer-sheets`) are
documented here. The package
follows semantic versioning. Each release records the **exact Univer engine
version** it bundles, so a given package version is reproducible and any release
can be rolled back to (see constitution Principle VIII).

## [0.1.5] - 2026-07-07

Headless-safe converter entry so the backend can reuse the FE converter.

### Added
- **`@levichco/finsheets/node`** subpath — a Node-safe barrel that re-exports only
  the pure-data converter/assembly functions (`parseXlsxToSnapshot`, `buildExcelWorkbook`,
  `buildShellWorkbook`, `diffSheet`/`highlightSnapshot`, types) with **no React, CSS,
  Univer UI presets, or icons**. Fixes headless import of the package in Node/tsx
  (the main `.` barrel pulls the whole FE stack, which can't load server-side).
  Backend: `import { parseXlsxToSnapshot } from "@levichco/finsheets/node"`.
- Second tsup entry (`src/node.ts`) → `dist/node.js` (+ `.d.ts`), exposed via the
  `"./node"` export condition. The React UI barrel (`.`) is unchanged.

### Hardening (production-readiness audit)
- **Version-history retention** — auto-checkpoints are now capped (named / original /
  restore kept forever) so long sessions on large workbooks don't grow IndexedDB / heap
  without bound; added `VersionStore.deleteVersion`.
- **Converter tests** — added round-trip tests for `parseXlsxToSnapshot` (values, formulas,
  bold, merges), `buildShellWorkbook`, `diffSheet`, and the `./node` entry (14 → 19 tests).
- **Demo build isolation** — `demo:build` now emits to `dist-demo/` so it never overwrites
  the published `dist/` library output.
- **Leak fix** — the product app now disposes its Univer event subscriptions + activation
  timers on remount/unmount (it remounts on every tab switch).

## [0.1.4] - 2026-07-07

Release-hygiene fix so consumers can pin a reproducible tag under the current name.

### Fixed
- **Version/name tag mismatch** — `v0.1.3` was tagged as the old name `@levich/univer-sheets`
  while `main` (same 0.1.3) had already been renamed to `@levichco/finsheets`, so no tag
  matched the current name (consumers had to install from `#main`). This cuts a fresh
  **`v0.1.4`** from `main` under `@levichco/finsheets` — a pinnable, reproducible tag.

### Added
- **`publishConfig`** (GitHub Packages registry) so the package can be **published**
  (`npm publish`) instead of installed via `github:` — a published install pulls only
  `dist` + runtime deps (per the `files` allow-list), avoiding the dev build-toolchain bloat
  a git install drags in.

## [0.1.3] - 2026-07-07

The **lazy multi-sheet product** + **Version History** + **Google Fonts** release.

### Added
- **Lazy shell-workbook** (`buildShellWorkbook`) — one workbook where every sheet is a
  tab but only the active sheet is hydrated; opens a 33 MB / 69-sheet workbook instantly.
- **`<SheetTabBar>`** — Google/Excel-style bottom tabs (rename · duplicate · colour ·
  hide/unhide · move · delete-with-confirm · searchable all-sheets list · zoom · one-tab
  scroll), the sole controller of the active sheet (native footer hidden via `sheetBar={false}`).
- **Version History** (browser-only, IndexedDB): `createVersionStore`, `<VersionHistoryDrawer>`
  (clock trigger, day-grouped list, Current-version badge, View original, ⋮ Name / Make a copy,
  Restore — non-destructive), `diffSheet`/`highlightSnapshot` ("Highlight changes" /
  "Show unmodified rows" + "No changes to this sheet in this revision"), and `readOnly` preview
  mode on `<LevichSheet>`.
- **Google Fonts library** — on-demand loader (`ensureGoogleFont(s)`, `ensureFontsForSnapshot`,
  `fontsInSnapshot`, `GOOGLE_FONTS`); toolbar Font dropdown gains a search box + ~50 fonts;
  **sticky current font** (pick once → applies to everything you type, first char included).
- **Office-font aliasing** (`office-fonts.css`) — Calibri→Carlito etc. so imported text isn't serif.
- **Shared `<ColorPanel>`** — full Google palette + Standard + custom HSV/hex picker, used by the
  toolbar and the sheet-tab colour menu.
- Host hooks for manifest-driven sheet visibility (`onHideActiveSheet` / `onShowSheet` /
  `hiddenSheetList` / `canHideActiveSheet`).

### Changed / Fixed
- Fonts render correctly on Univer's canvas — apply after load + `refreshCanvas()` re-measure.
- Colour picker submenus positioned correctly (`position: absolute`).

### Engine
- Univer (`@univerjs/presets`, `@univerjs/preset-sheets-core`): **0.25.0** (exact pin)

## [0.1.2] - 2026-07-06

Excel-fidelity pass on `.xlsx` import (fidelity target is **Excel**, not Google Sheets).

### Added
- **Filters** — `autoFilter` → `SHEET_FILTER_PLUGIN` (funnel dropdowns).
- **Conditional formatting** — `conditionalFormattings` → `SHEET_CONDITIONAL_FORMATTING_PLUGIN`
  (highlight-cell / color-scale / data-bar).
- **Hyperlinks** — stored cell links (`mailto:`/URLs) → `SHEET_HYPER_LINK_PLUGIN`, plus
  **auto-linkify** of plain-text single-email/URL cells.
- **Hidden rows/columns** (`col.hidden`/`row.hidden` → `hd`), tab colours, gridline toggle.
- **Large-font & wrapped-newline row-height** estimation; source default row-height/col-width.

### Changed / Fixed
- **Preserve source colours exactly** — removed the near-white-font guard (white-on-no-fill
  stays invisible, matching Excel).
- **Hidden fill-only rows** now captured (`includeEmpty:true`) — no more black-bar leak.
- **Merged-box sizing** — skip the height bump for multi-row merges.
- **Image anchors** treat hidden rows/cols as 0 (correct placement).
- Excel-standard column widths / row heights; `$`-prefixed range refs parsed.

### Engine
- Univer (`@univerjs/presets`, `@univerjs/preset-sheets-core`): **0.25.0** (exact pin)
- ExcelJS: **4.4.0**

---

## [0.1.1] - 2026-07-02

### Added
- **Save** — `⌘S`/`Ctrl+S` and a File ▸ Save menu item, with a new `onSave(snapshot)`
  prop for host persistence. Falls back to `localStorage["levich:save:<id>"]` and a
  "Saved" toast when no host hook is given.
- **Rename spreadsheet** — File ▸ Rename now renames the *document* (workbook) via a
  styled modal (replacing the native `prompt`), driving `FWorkbook.setName` + the
  `onRename(name)` host hook.
- **Strikethrough shortcut** — `Alt+Shift+5` toggles strikethrough.

### Changed
- **Toolbar state now reflects shortcuts.** `⌘B`/`⌘I`/`⌘U` (and strikethrough) are
  routed through the Facade style setters so the toolbar B/I/U/S̶ buttons light up
  when toggled by keyboard (previously only mouse clicks did). Bails while a cell
  editor is focused so in-cell rich-text formatting still works.
- **Sheet-tab Rename** uses the same styled modal (with empty/duplicate-name
  validation) instead of the native `prompt`/`alert`.

### Engine
- Univer (`@univerjs/presets`, `@univerjs/preset-sheets-core`): **0.25.0** (exact pin)
- ExcelJS: **4.4.0**

---

## [0.1.0]

### Added
- Initial package scaffold and Spec Kit feature `001-univer-levich-sheets`.
- Core: `LevichSheet` component, workbook compiler, Univer lifecycle, full-fidelity
  ExcelJS export, configurable freeze/locked-columns, comments + width hooks,
  free pivot engine, Levich theme, and the `TransactionSheet` preset.
- Rich `.xlsx` import (styles, colours, merges, number formats, images, multi-sheet),
  Google-style sheet-tab management, cross-platform keyboard shortcuts, and the full
  ~526-function catalog on the free Univer formula engine.

### Engine
- Univer (`@univerjs/presets`, `@univerjs/preset-sheets-core`): **0.25.0** (exact pin)
- ExcelJS: **4.4.0**

---

## Release / rollback procedure

1. To adopt a new Univer version: change the exact pin in `package.json`, run
   `npm run build` and `npm run test`, record the new engine version here, then
   tag the release.
2. To roll back: pin the previous package version (or check out the previous tag
   while locally linked). Because the engine version is pinned exactly, the prior
   behavior is restored exactly.
