# FinSheets — PoC → Product Plan

> Written 2026-07-06. Turns the validated PoC into **FinSheets**, a multi-tenant,
> bring-your-own-infrastructure spreadsheet product where **only rendering runs on
> the FE and everything else (conversion, calculation, storage, persistence) runs
> on the server** for accuracy. Builds on `ROADMAP.md`, `RENDERING-ENGINE.md`,
> `VERSION-HISTORY.md`. Fidelity target: **Excel**.

## 0. Product definition
FinSheets is:
1. A **frontend package** (`@levich/univer-sheets`) that *renders* spreadsheets and sends edits — nothing else.
2. A **backend service** (`finsheets-service`) that does **all the heavy lifting**: `.xlsx`/`.csv` conversion, **server-side calculation** (the real fix for cross-sheet formulas), storage, persistence, version history, collaboration.
3. **Bring-your-own-infrastructure & multi-tenant**: a consumer plugs in *their own* **database URL** and *their own* **object-storage URL** (DigitalOcean Spaces / S3 / GCS / Azure / R2 / MinIO). The URLs are held in the BE per tenant; the FE never sees them.

**Guiding principle:** *Render on the client, compute & store on the server.*
Rendering is browser-bound (canvas); everything that affects **correctness** —
formula calc, number formats, cross-sheet references — happens **server-side** so
results are exact and consistent for every viewer.

## 0.5 Deployment model — SELF-HOSTED (DECIDED 2026-07-06)
The host platform (e.g. Finopz) **runs its own `finsheets-service`** (a Docker
image it deploys), and points it at **its own** Postgres and object storage via
**environment config** — no FinSheets-operated SaaS, no cross-customer secret
registry.
```
# finsheets-service .env (the host's infra)
DATABASE_URL=postgres://…               # the host's DB
STORAGE_ENDPOINT=https://nyc3.digitaloceanspaces.com
STORAGE_BUCKET=finopz-files             # the host's bucket
STORAGE_KEY=…  STORAGE_SECRET=…
INTERNAL_API_KEY=…                      # service↔service auth
```
- **Distribution:** ship `finsheets-service` as a **Docker image** (+ Compose /
  DO App Platform / k8s manifest). The FE package points at the host's service URL.
- **What this SIMPLIFIES vs the SaaS design:**
  - **No per-tenant DB/storage registry, no BYOD connection routing, no encrypted
    secret vault** — there's exactly **one** DB URL + **one** storage config, from env.
  - **Multi-tenancy becomes *intra-install workspace scoping*** — the single service
    serves the host's many workspaces/users, isolated by `workspace_id` (+ Postgres
    RLS), not by separate databases. The host's own auth (JWT with a workspace
    claim) drives it. (Cross-*customer* isolation = the host runs separate installs.)
  - Auth = trust the **host platform's JWT**; FinSheets validates the workspace claim.
- The §3.1/§3.2 "bring your own" pillars still hold, but "bring your own" now means
  **the host's env config**, resolved **once at boot**, not a per-request tenant lookup.

## 1. The killer insight — one engine, both sides
Univer ships an **official headless Node.js runtime**: `@univerjs/preset-sheets-node-core`
+ `@univerjs/engine-formula` (Node ≥ 18.17). So the **exact same Univer engine** we
verified (526 functions, Apache-2.0) that **renders** on the FE also **calculates**
on the BE — no second calc library, no GPL (rules out HyperFormula, which is
GPLv3/commercial). Our `xlsx-to-snapshot.ts` converter is already isomorphic
(proven running in Node). **The whole stack is one codebase, two runtimes.**

→ **Server-side calc (Option C) is now cheap**: load the full workbook snapshot into
headless Univer on the BE, let it compute every formula (direct + cross-sheet
SUMIFS/INDEX-MATCH), read back computed **values**, and ship those to the FE. The FE
renders values; the cross-sheet `#VALUE!`/blank problem disappears.

## 2. Architecture
```
┌────────────── FE (browser) — RENDER ONLY ──────────────┐
│  @levich/univer-sheets                                 │
│   • Univer canvas render (viewport-virtualized)        │
│   • Lazy per-sheet loader + IndexedDB cache            │
│   • Edit capture → send ops to BE                      │
│   • thin, tenant-scoped API client (JWT)               │
└───────────────▲──────────────────────┬─────────────────┘
     manifest + per-sheet VALUES        │ edits / import
┌───────────────┴──────────────────────▼─────────────────┐
│  finsheets-service (BE) — EVERYTHING ELSE              │
│  ┌──────────────────────────────────────────────────┐ │
│  │ API gateway · auth · TENANT RESOLVER             │ │
│  └──────────────────────────────────────────────────┘ │
│  Processing Unit   ExcelJS → IWorkbookData (converter) │
│  Calc Engine       headless Univer + engine-formula    │
│                    (full workbook, recompute on edit)  │
│  DB Layer          per-tenant Postgres URL (BYOD)      │
│  Storage Layer     per-tenant bucket URL (S3-compat)   │
│  Cache Layer       Redis (hot snapshots, calc results) │
│  [Collab later]    Yjs + Hocuspocus + Redis fan-out    │
└─────────────────────────────────────────────────────────┘
```

## 3. The three "bring your own" pillars

### 3.1 Bring-your-own Database (BYOD) — multi-tenancy
Consumers pass **their own Postgres connection URL**; the BE stores it in a
**tenant registry** and **routes** each request to the right database. Research
(2026) says the pragmatic default is a **hybrid** model:
- **Pooled tier** — shared Postgres with **Row-Level Security (RLS)** keyed by
  `tenant_id` (cheapest, safe default; a missing filter can't leak across tenants).
- **Isolated tier (BYOD)** — enterprise/compliance tenants get a **dedicated
  database** via their own connection string; the app keeps a **connection
  registry** and routes queries per tenant.
- Serverless options (**Neon** branch-per-tenant, **Turso** DB-per-tenant) sit
  between the two if we want cheap isolation.

**Flow:** request → auth (JWT) → **tenant resolver** looks up the tenant's DB URL in
the registry → connect (pooled) → execute. The DB URL lives **only in the BE**,
encrypted at rest.

### 3.2 Bring-your-own Storage — the file pool
`.xlsx` uploads and generated snapshots live in **object storage the tenant owns**.
Everything S3-compatible works (**DigitalOcean Spaces, AWS S3, GCS, Azure Blob, R2,
MinIO**). We use a **storage-abstraction driver** (FlyDrive / `@tweedegolf/storage-
abstraction` / AWS SDK) so the same code targets any provider — the tenant just
supplies **bucket URL + credentials** (or per-bucket access keys, which DO Spaces
now supports for isolation). Local FS driver for dev.

**Flow:** upload → BE streams to the **tenant's bucket** → Processing Unit converts →
snapshots cached/stored → FE fetches per-sheet values.

### 3.3 Tenant isolation
- Every row/object keyed by `workspace + document`; never readable across tenants.
- Pooled tier: **Postgres RLS**. Isolated tier: separate DB. Storage: per-tenant
  bucket/prefix + scoped keys.
- Two auth planes: service↔service (`x-internal-api-key`, constant-time) and
  user (cookie-JWT). No anonymous access; fail-closed.

## 4. Server-side calculation (Option C) — the accuracy engine
- On import/open, load the full `IWorkbookData` into **headless Univer** on the BE.
- Univer's `engine-formula` computes **every** formula against the full dependency
  graph — direct refs *and* computed cross-sheet (`=SUMIFS('Model - IS'!…)`),
  regardless of what the file cached.
- Read back **computed values** via the Facade; store them (+ formulas) in Postgres,
  cache hot results in Redis.
- The FE lazy-loads **values** per sheet → renders instantly, always correct.
- **On edit:** FE sends the op → BE applies it to the headless workbook →
  **incremental recalc** of affected cells → push changed values back to the FE
  (exactly the Google model). Interim static direct-ref resolution (already
  prototyped) bridges until this is live.

## 5. Component breakdown / workstreams
- **WS-1 FE package** — lazy-sheet loading as a real feature (`manifest` + `loadSheet`),
  edit capture, IndexedDB cache, tenant-scoped client.
- **WS-2 Processing Unit** — productionize `xlsx-to-snapshot.ts` behind `POST /import`;
  converter optimizations (empty-styled-cell collapse; images → URLs).
- **WS-3 Calc Engine** — headless Univer service; full recompute + incremental
  recalc-on-edit; computed-value extraction.
- **WS-4 DB Layer** — tenant registry, connection routing, RLS (pooled) + BYOD
  (isolated); schema (documents, sheets, versions, tenants).
- **WS-5 Storage Layer** — storage-abstraction driver; per-tenant bucket config;
  image serving.
- **WS-6 Multi-tenancy & auth** — tenant resolver, JWT, encryption of tenant secrets,
  isolation tests.
- **WS-7 Version history** — per `VERSION-HISTORY.md` (Postgres snapshots + Redis + IndexedDB).
- **WS-8 App shell** (later) — hub, folders, templates, sharing.
- **WS-9 Collaboration** (later) — Yjs + Hocuspocus + Redis fan-out.

## 6. Data model (per-workspace, pooled tier)
- `tenant` — id, name, **db_url (encrypted, nullable = use pooled)**, **storage_config
  (encrypted: provider, bucket, region, keys)**, tier.
- `finsheets_document` — id, tenant_id, title, type (blank|flux_drill|imported),
  head_snapshot, updated_at. (RLS on tenant_id.)
- `finsheets_sheet` — id, document_id, sheet_id, name, order, hidden, **snapshot
  (jsonb / gzip bytea: values + formulas + styles)**.
- `finsheets_version` — id, document_id, seq, state_blob, label, kind, authors, created_at.
- `finsheets_manifest` — document_id → ordered sheet metadata.

## 7. API surface (tenant-scoped, JWT)
```
POST /documents/import            multipart .xlsx → convert+calc → doc + manifest
GET  /documents/:id/manifest      tab list + workbook resources
GET  /documents/:id/sheets/:sid   one sheet's COMPUTED values (+ formulas for edit)
POST /documents/:id/edit          apply an op → incremental recalc → changed cells
GET  /documents/:id/sheets/:sid/images/:img
POST /documents/:id/versions      · POST .../restore · GET .../versions
POST /admin/tenants               register tenant: db_url + storage_config (BYO)
```

## 8. Phased roadmap (PoC → product)
- **P0 — Rendering engine (in progress).** Lazy loading, import fidelity to Excel,
  static direct-ref resolution. *(Mostly done in PoC.)*
- **P1 — `finsheets-service` skeleton.** `POST /import` + `GET /manifest|sheets` reusing
  the converter; Postgres + Redis; wire the FE off static files onto the API.
- **P2 — Server-side calc (Option C).** Headless Univer; full recompute; computed-value
  serving. Kills cross-sheet gaps.
- **P3 — Self-host packaging + workspace scoping.** Docker image + env config (one DB
  URL, one storage config); `workspace_id` + Postgres RLS; validate the host
  platform's JWT (workspace claim); storage-abstraction driver. *(No cross-customer
  tenant registry — self-hosted.)*
- **P4 — Persistence & edit round-trip.** Edit ops → incremental recalc → push;
  autosave; server-side `.xlsx` export.
- **P5 — Version history.** Per the design doc.
- **P6 — App shell & sharing** · **P7 — Real-time collaboration** (Yjs/Hocuspocus).

## 9. Locked tech decisions (from research)
- **Calc engine = headless Univer** (`preset-sheets-node-core` + `engine-formula`,
  Apache-2.0, isomorphic — same engine as render). **Not** HyperFormula (GPLv3).
- **Converter = ExcelJS** (already isomorphic; runs in Node).
- **DB = PostgreSQL**, hybrid tenancy: **shared + RLS** default, **BYOD** for isolation.
- **Storage = S3-compatible via an abstraction driver** (FlyDrive / storage-abstraction);
  tenant brings bucket URL + keys (DO Spaces / S3 / GCS / Azure / R2 / MinIO).
- **Cache = Redis.** **Collab (later) = Yjs + Hocuspocus.**
- **FE = render-only**; secrets (DB/storage URLs) live only in the BE, encrypted.

## 10. Open decisions
- **Tenancy default**: pooled-RLS-first vs BYOD-first (cost vs isolation; research
  leans pooled+RLS default, BYOD for enterprise).
- **Storage per-tenant scoping**: per-bucket keys vs prefix-scoped (DO Spaces lacks
  S3-style IAM per prefix — may push toward bucket-per-tenant for hard isolation).
- **Calc trigger granularity**: full recompute on open vs incremental only.
- **Offline timestamping & Flux data-refresh** (from VERSION-HISTORY.md).

---
### Sources
- Univer headless Node.js: https://docs.univer.ai/guides/sheets/getting-started/node · https://github.com/dream-num/univer · https://docs.univer.ai/reference/packages/plugins/univerjs/engine-formula
- Multi-tenancy patterns (2026): https://dev.to/young_gao/multi-tenant-architecture-database-per-tenant-vs-shared-schema-1n2e · https://www.bytebase.com/blog/multi-tenant-database-architecture-patterns-explained/ · https://northflank.com/blog/multi-tenant-saas-platform-deployment
- HyperFormula (alt engine, GPLv3): https://github.com/handsontable/hyperformula · https://hyperformula.handsontable.com/
- Storage: https://www.digitalocean.com/products/spaces · https://docs.digitalocean.com/products/spaces/how-to/use-aws-sdks/ · https://github.com/slynova-org/flydrive · https://github.com/tweedegolf/storage-abstraction
