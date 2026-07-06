# finsheets-service — Detailed Implementation Plan (Finopz `fiab` conventions)

> Written 2026-07-06 after auditing the `fiab` backend. `finsheets-service` is a
> **new service in the `fiab` Turborepo**, built exactly like `integration-service`
> / `core-service` (plain Express + TS, `@fiab/*` shared packages, `x-internal-api-key`,
> schema-per-tenant Postgres, nginx subdomain, GHCR Docker image). Supersedes the
> generic §3 of `PRODUCT-PLAN.md` where Finopz already provides the primitive.
> Fidelity target: **Excel**. Calc: **headless Univer**.

## 0. It fits the existing pattern — confirmed
Finopz backend = **`fiab/` Turborepo**, npm workspaces (`apps/*`, `packages/*`), Node ≥22,
**plain Express + TypeScript** (no NestJS). Existing services: `auth-service`:3000,
`core-service`:3001, `upload-service`:3002, `integration-service`:3003,
`cron-service`:3004. **`finsheets-service` = `apps/finsheets-service`, npm name
`fiab-finsheets-service`, port 3006, subdomain `finsheets.finopz.ai`.**

**What Finopz already gives us (so we DON'T rebuild it):**
| Our earlier "product" concern | Finopz already has it |
| --- | --- |
| Auth / JWT | **`@fiab/auth`** — `createAuthMiddleware` (local JWT verify + auth-service liveness check); cookie `access_token`. |
| Service-to-service auth | **`x-internal-api-key`** (timing-safe), per-service key, `/internal/*` routes. |
| Multi-tenancy | **Schema-per-tenant** — every workspace has its own PG schema `workspace_<uuid>`; `workspaces.schema_name`; `connection-pool-manager` routes via `search_path`. **NOT RLS, not BYOD.** |
| Storage / file pool | **`upload-service`** (S3 / volume). We fetch files from it over its internal API. |
| DB config | One shared Postgres via `DATABASE_URL` (managed DO in prod). |
| Deploy | Multi-stage Dockerfile (`node:22.22.3`) → GHCR `ghcr.io/levichco/fiab-<svc>` → nginx per-service subdomain. |

→ **The "self-hosted, bring-your-own DB/storage" model IS how `fiab` already runs.**
Finopz deploys `fiab` with its `DATABASE_URL` + upload-service storage; `finsheets-service`
inherits both. Multi-tenancy = **schema-per-tenant** (existing), *not* the RLS/BYOD
design from PRODUCT-PLAN §3 — **use Finopz's pattern.**

## 1. Framework & baseline (mirror `integration-service`/`core-service`)
- **Express 4.18** + TS; dev `nodemon --exec tsx src/index.ts`; build `tsc --build`; start `node dist/index.js`.
- **Entry** (`src/index.ts`, copy core's): `dotenv` FIRST → `express()` → `helmet()` →
  `HealthService.initialize({serviceName:'finsheets-service', ...})` → `cors({origin:config.frontend.url||true, credentials:true})` →
  `express.json({limit:'25mb'})` (bigger, for imports) → `cookieParser()` → `requestLogger` →
  `GET /health` → `app.use(routes)` → 404 → `errorLogger` → graceful SIGINT/SIGTERM.
- **Depend on** (`file:../../packages/*`): `@fiab/auth`, `@fiab/health`, `@fiab/logger`,
  `@fiab/response`, `@fiab/config`, `@fiab/redis`; plus `@prisma/client`, `express@^4.18.2`,
  `cors`, `helmet`, `cookie-parser`, `zod`, `axios`, `dotenv`, `uuid`.
- **Our own runtime deps**: `exceljs` (convert), `@univerjs/presets`
  `@univerjs/preset-sheets-node-core` `@univerjs/engine-formula` (**headless calc**),
  and the converter — reuse `@levich/univer-sheets`'s `parseXlsxToSnapshot` (import the
  git package, or vendor `xlsx-to-snapshot.ts` server-side).
- **tsconfig** extends `@fiab/typescript-config/node.json`; project `reference` to `../../packages/auth`.
- **Gotchas to mirror**: pin `node:22.22.3` in the Dockerfile; generate Prisma to local
  `generated/`; load `dotenv` before env-reading imports; wrap `createAuthMiddleware` once
  with `requireSecretInProduction('JWT_SECRET', …)`; copy `internalAuth.middleware.ts` (timing-safe).

## 2. Auth (copy core's wrappers)
- **Public routes** `/api/*` → `authMiddleware = createAuthMiddleware({ jwtSecret, authServiceUrl })`.
  `req.user` carries `{ userId, email, role, workspace{...} }` → drives workspace scoping.
- **Internal routes** `/internal/*` → `internalAuthMiddleware` with **`FINSHEETS_SERVICE_INTERNAL_API_KEY`**.
- To call other services (upload-service, core), send `x-internal-api-key: <THAT_SERVICE>_INTERNAL_API_KEY`.

## 3. Multi-tenancy & DB (schema-per-tenant — Finopz's pattern)
- **No new tenant registry, no RLS.** finsheets tables live in the **workspace schema**
  (`workspace_<uuid>`). Resolve the workspace from `req.user.workspace`, look up
  `schema_name`, get a schema-scoped Prisma client via the **`connection-pool-manager`**
  pattern (copy core's `WorkspaceService.getWorkspacePrismaClient` + pool manager).
- **`prisma/schema.prisma`** (workspace-scoped, generated to `generated/finsheets-prisma`):
  - `finsheets_document(id, title, type, head_snapshot_id, created_by, updated_at)`
  - `finsheets_sheet(id, document_id, sheet_id, name, ord, hidden, snapshot jsonb|bytea)` — values+formulas+styles per sheet
  - `finsheets_version(id, document_id, seq, state_blob, label, kind, authors jsonb, created_at)`
  - `finsheets_manifest(document_id, sheets jsonb)` — tab metadata
- **Workspace-schema migrations** follow Finopz's custom `vNNN` versioned system
  (registry `fiab-shared/.specify/registry/migrations.json` + `db:workspace:migrate:direct`).

## 4. Storage — reuse `upload-service`
`.xlsx` uploads and generated images go through **upload-service** (the existing file pool /
S3 / volume). finsheets-service **does not** implement its own storage abstraction:
- Import: client uploads to upload-service → gets a file id/URL → calls
  `finsheets POST /api/documents/import { fileId }` → we fetch the bytes from upload-service
  (internal API, `x-internal-api-key`) → convert.
- Images extracted from `.xlsx` → store via upload-service → serve as URLs.
- (If a leaner path is wanted, finsheets can accept the multipart directly and forward to
  upload-service — but reuse it either way.)

## 5. Server-side calc (headless Univer) — the accuracy engine
- **`services/calc.service.ts`**: boot a **headless Univer** instance
  (`createUniver` + `UniverSheetsNodeCorePreset` + `@univerjs/engine-formula`), load the
  full `IWorkbookData`, let it compute **all** formulas (direct + cross-sheet), read back
  **computed values** via the Facade.
- Persist computed values (+ formulas) per sheet; cache hot results in **`@fiab/redis`**.
- Serve **values** to the FE per sheet (lazy). On edit (P4): apply op → incremental recalc →
  push changed cells. Kills cross-sheet `#VALUE!`/blank permanently.

## 6. Directory structure (`fiab/apps/finsheets-service/`)
```
src/
  index.ts                      # Express bootstrap (copy core)
  config/index.ts               # @fiab/config loader
  middlewares/
    auth.middleware.ts          # wraps createAuthMiddleware
    internalAuth.middleware.ts  # copy verbatim
  routes/
    documents.routes.ts         # /api/documents/* (authMiddleware)
    internal.routes.ts          # /internal/* (internalAuthMiddleware)
  services/
    import.service.ts           # ExcelJS → IWorkbookData (converter) + split
    calc.service.ts             # headless Univer recompute → values
    workspace-db.service.ts     # schema-per-tenant Prisma (copy core pattern)
    upload-client.service.ts    # fetch/put files via upload-service (internal)
  schemas/                      # zod request validation
prisma/
  schema.prisma                 # workspace-scoped finsheets tables
Dockerfile                      # node:22.22.3, multi-stage (copy integration's)
package.json  tsconfig.json  .env.example
```

## 7. API surface
```
# public (authMiddleware — JWT, workspace from req.user)
POST /api/documents/import        { fileId }  → convert + calc → { documentId, manifest }
GET  /api/documents/:id/manifest              → tab list + workbook resources
GET  /api/documents/:id/sheets/:sid           → one sheet's COMPUTED values (+ formulas)
POST /api/documents/:id/edit                  → apply op → incremental recalc → changed cells   (P4)
POST /api/documents/:id/versions  · /restore · GET /versions                                      (P5)
# internal (internalAuthMiddleware — x-internal-api-key)
POST /internal/documents/reindex  etc.
GET  /health
```

## 8. Deployment (mirror the other services)
- **Dockerfile** — copy `integration-service/Dockerfile` (simplest, single-Prisma), pin
  `node:22.22.3`, multi-stage, `prisma generate`, non-root `nodejs:1001`, `EXPOSE 3006`,
  `HEALTHCHECK /health`, `CMD migrate-and-start.sh`.
- **Compose** — add `finsheets-service` to all three:
  - `docker-compose.yml` (local): `3006:3006`, `fiab-network`.
  - `deploy/ssl/docker-compose.staging.yml` + `deploy/prod/docker-compose.prod.yml`:
    `container_name: finopz-finsheets-service`, no exposed ports,
    `image: ghcr.io/levichco/fiab-finsheets-service:${IMAGE_TAG:-prod}`, healthcheck, network.
- **nginx** — new `server` block (prod `finopz.ai.conf`) + `upstream` (staging) →
  `finopz-finsheets-service:3006`, subdomain `finsheets.finopz.ai` / `finsheets.stage.finopz.ai`.
- **Root `package.json`** — add `dev:finsheets` / `build:finsheets` / `start:finsheets`
  (`--filter=finsheets-service`). Turbo picks it up via the workspace glob automatically.
- **CI** — register the GHCR image build for `fiab-finsheets-service`.

## 9. Phased build (concrete)
- **P1 — Skeleton service.** Scaffold `apps/finsheets-service` (Express + health + auth
  wrappers + `/api/documents/import` returning a converted manifest, reusing the converter;
  no DB yet, in-memory). Wire the FE PoC (`poc-app.tsx`) to hit it instead of static files.
- **P2 — Calc.** `calc.service.ts` headless Univer full recompute → serve computed values.
- **P3 — Persistence (schema-per-tenant).** Prisma workspace tables; store snapshots +
  manifest; Redis cache; upload-service for files/images.
- **P4 — Edit round-trip** (incremental recalc) · **P5 — Version history** · **P6 — app
  shell / sharing** · **P7 — collaboration** (Yjs + Hocuspocus; note core already hosts
  Socket.IO via `@fiab/websocket`, but collab uses its own Hocuspocus socket).

## 10. Corrections to PRODUCT-PLAN.md (Finopz reality)
- **Multi-tenancy** = **schema-per-tenant** (Finopz existing), not shared-RLS or BYOD registry.
- **Storage** = **reuse upload-service**, not a new storage-abstraction driver.
- **Auth** = **`@fiab/auth`** (existing), not a new identity system.
- **DB URL / storage** = the **`fiab` env** Finopz already sets; finsheets inherits it.
- Everything else (headless-Univer calc, FE render-only, phases) stands.
