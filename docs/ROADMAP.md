# FinSheets (@levich/univer-sheets) — Forward Roadmap

> Written 2026-07-04, from v0.1.1. This is the **next-level** plan (Google-Sheets/
> Excel-class): large-file handling, backend, persistence, version history,
> collaboration, and fidelity parity. The original package build plan (`PLAN.md`,
> Stages 1–9) is essentially delivered. Design docs: `VERSION-HISTORY.md`.

## 0. Vision
A reusable, embeddable spreadsheet that behaves like Google Sheets / Excel —
rich rendering, arbitrarily large files, autosave, version history, and (later)
real-time collaboration — built entirely on the **free Univer engine**
(Apache-2.0) + **ExcelJS** (MIT), distributed as a **versioned git package** so
consumers `npm update` to adopt features.

## 1. Where we are (v0.1.1) — validated
- **FE package**: `<LevichSheet>`, rich `.xlsx` import (custom ExcelJS→`IWorkbookData`
  converter — styles/merges/formats/formulas/images/multi-sheet), full-fidelity
  export, ~526-function engine, cross-platform keyboard shortcuts, Google-style
  sheet management, Save (`onSave` + localStorage fallback), spreadsheet Rename,
  git-dependency distribution.
- **Proven PoC** (the large-file fix): the converter **runs unchanged in Node**
  (3.0 s for a 5.5 MB / 69-sheet / 170k-cell workbook); the FE loads a
  **manifest + one sheet at a time** and renders each in **~10–40 ms**. The file
  that used to freeze the tab now opens instantly. → `scripts/xlsx-poc.mjs`,
  `demo/poc-app.tsx`.

## 2. Target architecture
Rendering stays on the FE (canvas is browser-bound); **parsing/conversion,
storage, and caching move to the backend.**

```
┌─────────────────────────── FE (browser) ───────────────────────────┐
│  @levich/univer-sheets                                              │
│   • Univer render (canvas, viewport-virtualized)                   │
│   • Lazy sheet loader (manifest → active sheet → fetch-on-tab)     │
│   • IndexedDB cache (offline + instant re-open)                    │
│   • thin API client                                                │
└───────────────▲───────────────────────────────┬────────────────────┘
                │ manifest + per-sheet JSON       │ save / autosave
┌───────────────┴───────────────────────────────▼────────────────────┐
│  finsheets-service (BE)                                             │
│   • Processing Unit — ExcelJS parse → IWorkbookData (per sheet)    │
│   • DB Layer — PostgreSQL (documents, sheets, versions)           │
│   • Cache Layer — Redis (hot snapshots, manifests)                │
│   • [later] Collaboration — Yjs + Hocuspocus + Redis fan-out      │
└─────────────────────────────────────────────────────────────────────┘
```

## 3. Phased plan

### Phase A — FE hardening (no backend yet)
Promote the PoC into first-class package capability.
- **Lazy-sheet loading as a package feature** — a `sheetProvider`/manifest prop so
  the host renders a tab bar and fills sheets on demand (today it's demo-only).
- **Converter optimization** — collapse **empty-but-styled cells** into row/column
  default styles. (The 69-sheet file is 33 MB of JSON, mostly styled-empty cells;
  this cuts heavy sheets ~3–5×.)
- **Deferred-dispose on sheet swap** — avoid the "unmount root during render" warning.
- Exit: a host can open a huge multi-sheet workbook from pre-converted JSON with
  instant tab switching, shipped in the package (v0.2.0).

### Phase B — Backend: import & render at scale (`finsheets-service`)
- **Processing Unit** — `POST /documents/import` (multipart `.xlsx`) → ExcelJS →
  per-sheet `IWorkbookData` + manifest. (Literally `scripts/xlsx-poc.mjs` behind HTTP.)
- **DB Layer** — PostgreSQL: `finsheets_document`, `finsheets_sheet` (per-sheet
  JSON, `jsonb` or gzip'd `bytea`), manifest.
- **Cache Layer** — Redis: hot manifests + recently opened sheets.
- **Images** — store once, serve as **URLs** (lazy, browser-cacheable) instead of
  base64 in the JSON.
- **FE** — point the lazy loader at `GET /manifest` + `GET /sheets/:id`.
- Exit: import & render arbitrarily large workbooks from the real service.

### Phase C — Persistence & edit round-trip
- **Save → BE** — `onSave(snapshot)` persists per-sheet snapshots (reuses `getSnapshot()`).
- **Autosave** — IndexedDB local (instant) + debounced BE sync.
- **Server-side export** — Univer JSON → `.xlsx` on the BE for large files.
- Exit: edits survive reload; export works at scale.

### Phase D — Version history (single-user) — design done
Per `VERSION-HISTORY.md` **v1 stack = Postgres + Redis + IndexedDB** (Yjs/Hocuspocus
deferred).
- Checkpoint snapshots (JSON) on triggers; Redis debounce; IndexedDB offline.
- UI: right-drawer timeline, preview (read-only), non-destructive restore, named versions.
- Exit: Google-Docs-style history on single-user documents.

### Phase E — The FinSheets app (Google-Sheets hub)
- Home/hub: new sheet, recents, **folders/groups**, **templates** (Flux drill = a
  data-backed template).
- **Per-sheet sharing & permissions**.
- Exit: FinSheets is a standalone app, not just an embedded component.

### Phase F — Real-time collaboration
- **Yjs + Hocuspocus + Redis fan-out**; presence/cursors; conflict-free merge.
- Version history **swaps snapshot payload to Yjs state** (tables/UI unchanged).
- Main risk: the **Univer↔Yjs binding** (mutation subscribe/apply) — correctness.
- Exit: multiple people edit the same sheet live.

### Phase G — Fidelity parity & mega-sheets
- **Charts** (parse chart XML + a chart renderer), **conditional-formatting import**,
  **data validation**, **pivot tables**, **comments** round-trip.
- **Row/column windowing** for single sheets >100k cells (serve by row-range).
- Exit: near-Excel import fidelity + no single-sheet ceiling.

## 4. Cross-cutting workstreams
- **Distribution** — semver git tags; consumers `npm update` (`#semver:^0.1.0`).
- **Testing** — unit + fidelity round-trip + **perf benchmarks** (time-to-render,
  conversion time, JSON size per sheet).
- **Security** — two auth planes, tenant isolation (workspace+document keys), fail-closed access.
- **Observability** — metrics on conversion time, render time, cache hit rate.

## 5. Immediate next steps (concrete)
1. **Promote lazy-sheet loading** to a package feature (prop/provider) + deferred dispose.
2. **Optimize the converter** (row/col default styles) to cut JSON size.
3. **Stand up `finsheets-service` skeleton** — `POST /import`, `GET /manifest`,
   `GET /sheets/:id`, reusing `scripts/xlsx-poc.mjs`; Postgres + Redis.
4. **Wire `poc-app.tsx`** to the API (replace static fetches) → end-to-end at scale.

## 6. Risks & open items
- **Univer↔Yjs binding** (Phase F) — the main correctness risk; prototype early.
- **Empty-styled-cell bloat** — addressed in Phase A converter optimization.
- **Charts** — free-tier gap (ExcelJS can't read charts; no free Univer chart renderer).
- **Mega single-sheet** (>100k cells) — needs row-windowing (Phase G).
- **Offline timestamping** & **Flux data-refresh** — open decisions in `VERSION-HISTORY.md`.
