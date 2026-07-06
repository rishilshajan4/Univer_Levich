# FinSheets — Rendering Engine: Detailed Plan

> Written 2026-07-04, from the validated `?poc` (BE-convert + FE lazy per-sheet
> render). This turns the PoC into a **production capability**: open arbitrarily
> large multi-sheet workbooks instantly, render each sheet perfectly, load on
> demand. Parent: `ROADMAP.md` (Phases A–B). Rendering stays on the FE (canvas is
> browser-bound); parsing/conversion/storage move to the BE.

## 0. Goal & non-goals
**Goal:** any `.xlsx` (5 MB, 69 sheets, 170k cells and beyond) opens in <300 ms,
each sheet renders in <100 ms, tabs load on demand — as a first-class package
feature, not a demo.

**Non-goals (later phases):** editing persistence to BE (C), version history (D),
collaboration (F), charts (G). This plan is **view + render fidelity + scale**.

## 1. What the PoC proved (baseline)
- Converter `parseXlsxToSnapshot` runs **unchanged in Node** — 3.0 s for the whole
  workbook. → the BE Processing Unit.
- FE loads a **7 KB manifest + one sheet** and renders in **~10–40 ms**; caches on
  tab-click. → the lazy loader.
- Bug fixed: unstable `data`/`columns` refs caused an infinite Univer
  dispose/recreate loop → module-level stable refs.

## 2. The two hard problems (why this is more than the PoC)

### 2.1 Cross-sheet formulas — THE critical one
Financial models cross-reference sheets constantly (`='Balance sheet'!D5`,
`=SUM(Model!A:A)`). If we lazy-load **one** sheet, formulas that reference an
**unloaded** sheet cannot recompute → they'd show `#REF!`/`0`/stale.

**Design — two modes:**
- **View mode (Phase 1, default):** render the **cached values** Excel last
  computed. ExcelJS exposes them and our converter already carries a cell's
  cached result alongside its formula. So a cross-sheet formula **displays its
  correct last value without recompute** — exactly how Excel shows a freshly
  opened file. We must tell Univer **not to force a recalculation** of formulas
  whose dependencies aren't loaded (else it overwrites the cached value with an
  error). Options: (a) load formulas as values in view mode, keeping `f` in a
  side-channel for round-trip; or (b) suppress the initial recalc pass.
- **Edit mode (Phase 2):** **dependency-aware loading** — when the user edits or a
  formula needs another sheet, load the referenced sheet(s) on demand (parse the
  formula's sheet refs, fetch them, then recalc). Precompute a lightweight
  cross-sheet dependency map on the BE so we know what to pull.

> Decision for v1: **ship View mode with cached values.** It's correct for the
> "open a huge model and read it" use case and sidesteps the dependency graph.
> **Known limitation:** a cross-sheet formula whose result was NOT cached in the
> `.xlsx` (e.g. full-column `=SUMIFS('Model - IS'!AI:AI,…)`) renders **blank** —
> there's no cached value to show and it can't recompute without the other sheet.
> Google avoids this by computing server-side against the full model.
>
> **DECIDED (2026-07-04): the full fix is Option C — a server-side calc engine —
> to be built LATER** (edit/collab phase). Run Univer's formula engine (isomorphic;
> the same one that passed 526 functions, runs in Node) on the BE holding the
> whole workbook + dependency graph; recompute on edit; push computed *values* to
> the FE, which lazy-loads and renders. That is exactly Google's architecture.

### 2.2 Smooth sheet-swap without losing edits
The PoC **remounts** Univer per tab (`key`) — simple, but it **disposes the engine
each swap** (flash, and any in-memory edit is lost). Two strategies:

| | Option 1 — Remount (PoC) | Option 2 — In-place fill (target) |
|---|---|---|
| Mechanism | `key={sheetId}` → dispose+recreate | One workbook; all sheets as **empty stubs**; fill a sheet's data on activation |
| Tabs | Our custom tab bar | Univer's **native** tab bar (all sheets present) |
| Swap cost | ~40 ms recreate | near-zero (data already-in-engine after first load) |
| Edits | Lost on swap unless snapshotted back | **Preserved** |
| Complexity | Low | Higher — needs style-registry merge + a "set worksheet data" path |

**Plan:** ship **Option 1** for view-only quickly; migrate to **Option 2** for the
real feature (preserves edits, native tabs, smoother). For Option 1 with editing,
snapshot the active sheet back into the cache before swapping.

## 3. Data contracts
```jsonc
// manifest (tiny — loads first)
{
  "documentId": "…",
  "activeSheetId": "sheet_2",
  "workbookResources": { /* defined names, shared workbook-level state */ },
  "sheets": [ { "sheetId":"sheet_2", "name":"Executive Dashboard",
                "order":1, "hidden":0, "tabColor":"#…" }, … ]
}
// per-sheet snapshot (fetched on activation) — single-sheet IWorkbookData
{ "id":"sheet_2", "name":"…", "sheetOrder":["sheet_2"],
  "styles": { /* PRUNED to this sheet */ },
  "sheets": { "sheet_2": { /* cellData, merges, rowData, columnData, freeze */ } },
  "resources": [ /* this sheet's images only */ ] }
```
**Workbook-level state** (defined names / named ranges) must load **upfront** in
the manifest, because cross-sheet references and some formulas depend on them.

## 4. Converter optimizations (cut JSON size)
- **Collapse empty-but-styled cells → row/column default styles.** The 69-sheet
  file is **33 MB** of JSON, mostly empty-but-styled cells across full grids.
  Emitting a row/col default instead of per-cell `s` cuts heavy sheets ~3–5×.
- **Images as URLs**, not base64 — store once on the BE, serve lazily, browser-cached.
- **Prune per-sheet styles** (already done in the PoC splitter).
- Target: biggest sheet <500 KB (from 1.8 MB), total corpus <10 MB (from 33 MB).

## 5. Package API (the FE feature)
Add a **lazy/manifest mode** alongside today's single-`snapshot` mode:
```ts
<LevichSheet
  manifest={manifest}                              // tabs + workbook resources
  loadSheet={(sheetId) => Promise<SheetSnapshot>}  // host wires to BE or static
  // internal: tab bar from manifest, cache, active-sheet render, prefetch
/>
```
- The package renders tabs from the manifest, calls `loadSheet` on activation,
  caches results (memory + optional IndexedDB), and can **prefetch** neighbors.
- Small files keep using the existing `snapshot` prop unchanged.

## 6. BE service (`finsheets-service`) — rendering endpoints
- `POST /documents/import` (multipart `.xlsx`) → ExcelJS convert → per-sheet
  snapshots + manifest → Postgres + Redis → returns `documentId` + manifest.
- `GET /documents/:id/manifest`
- `GET /documents/:id/sheets/:sheetId`  (Redis → Postgres)
- `GET /documents/:id/images/:imageId`
- Storage: Postgres per-sheet JSON (`jsonb` or gzip'd `bytea`); Redis hot cache
  (manifest + recently opened sheets).

## 7. Edge cases to cover
- Hidden sheets & separators (`>>>`), tab colors, **per-sheet frozen panes**,
  merges, **per-sheet images**, number formats, RTL.
- **Defined names / named ranges** (workbook-level) → ship in manifest.
- Cross-sheet formulas (§2.1). Very wide/tall sheets → Univer viewport-virtualizes
  the draw; the model load is the only cost (already <40 ms/sheet).

## 8. Performance targets + benchmark harness
- Time-to-first-interactive **< 300 ms** (manifest + first sheet).
- Per-sheet render **< 100 ms**; manifest **< 50 KB**.
- Build a benchmark script measuring conversion time, per-sheet JSON size, and
  in-browser render time across sample files (extend `scripts/xlsx-poc.mjs`).

## 9. Milestones / sequencing
- **M1 — Converter optimization + benchmarks.** Empty-styled-cell collapse; size
  report. (FE-only, no BE.)
- **M2 — Package lazy API.** `manifest` + `loadSheet` + internal cache/tab bar
  (Option 1 remount, view-only). Migrate `poc-app.tsx` onto it.
- **M3 — Cross-sheet cached-value rendering.** Suppress recalc / values-mode so
  huge models read correctly. (Correctness milestone.)
- **M4 — finsheets-service skeleton.** Endpoints + Postgres + Redis; reuse the
  converter. Images as URLs.
- **M5 — Wire FE → BE.** Replace static fetches; end-to-end at scale.
- **M6 — Option 2 in-place fill.** Native tabs, smooth swaps, edit preservation.

## 10. Risks
- **Cross-sheet recompute** — mitigated by cached-value view mode (M3), dependency
  loading later.
- **In-place fill (M6)** — needs a style-registry merge + set-worksheet-data path;
  prototype against the Facade early.
- **Defined names** — must be workbook-level in the manifest, not per-sheet.
- **Converter size** — addressed in M1; verify against real files.
