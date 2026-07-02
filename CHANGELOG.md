# Changelog

All notable changes to `@levich/univer-sheets` are documented here. The package
follows semantic versioning. Each release records the **exact Univer engine
version** it bundles, so a given package version is reproducible and any release
can be rolled back to (see constitution Principle VIII).

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
