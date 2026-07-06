# FinSheets — Version History (design)

> Status: **design / not yet built.** FinSheets today is the **frontend-only**
> React package (`@levich/univer-sheets`). Version history is a future capability
> that requires a backend (`finsheets-service`). This document is the full design
> so the FE package can expose the right hooks/props and the service can be built
> against a fixed spec.
>
> **License note:** built on OSS only — we do **not** use
> `@univerjs-pro/edit-history-*` (unlicensed). (Yjs, when added in the collab
> phase, is MIT.)

---

## 0. Scope — v1 (single-user) vs later (collaborative)  · DECIDED 2026-07-04

Version history (**storage** problem) is decoupled from real-time collaboration
(**transport** problem), so we ship a lean single-user v1 first and add
collaboration later **without reworking history**.

**v1 — single-user "my own version history" · stack = PostgreSQL + Redis + IndexedDB.**
**Yjs + Hocuspocus are DEFERRED to the collaborative phase.** There are no
concurrent editors to merge, so no CRDT and no collab socket are needed.

| Tool | Job in v1 |
| ---- | --------- |
| **PostgreSQL** | Durable **version store** = source of truth. Rows hold a **Univer workbook-JSON snapshot** (from `getActiveWorkbook().getSnapshot()` — the *same* payload the Save feature already produces) + metadata (`label`, `kind`, `author_id`, `created_at`). **No Yjs changeset log in v1** → coarse "restore to checkpoint N," Google-Docs-versions style. |
| **Redis** | **Checkpoint debounce/queue** (buffer edits, cut a version after N min / M edits) + **hot cache** of the head snapshot and the recent-versions list for fast open. *Not* live fan-out yet (that's the collab phase). |
| **IndexedDB** | **Instant local autosave + offline buffer** on the client — keeps recent snapshots locally so history/undo works instantly and offline, syncing to Postgres on reconnect. |

**v1 pipeline:**

```
edit → debounced local autosave (IndexedDB)
     → sync to finsheets-service
     → Redis debounce/queue → checkpoint
     → Postgres version row  (+ Redis caches head snapshot)
```

**What differs from the full design below (which is the collab-phase target):**
- Snapshot = **Univer workbook JSON**, not encoded Yjs state.
- **No Hocuspocus** — plain HTTP save/sync, not a Yjs WebSocket room.
- **No CRDT merge, no Redis fan-out.**
- Everything else — timeline, non-destructive restore, preview, naming,
  retention, permissions, the API surface (§10) — is **unchanged**.

**Upgrade path (collab phase):** swap the snapshot payload to Yjs state, add
Hocuspocus + Redis fan-out (and optionally a fine-grained changeset log for
per-edit scrubbing). The version tables and the whole history UI stay as-is.

> The sections below (§1–§13) describe the **full/target** design. For v1, read
> them with the §0 substitutions above (JSON snapshots, no Yjs/Hocuspocus).

---

## 1. What the user gets (behaviour)

A **"Version history"** panel (right-side drawer), Google-Docs style:

- A **timeline** of versions, newest first, **grouped by day**.
- Each entry shows: **time**, **who edited** (avatars of contributors), and an
  **auto vs named** badge.
- Click a version → the sheet shows that past state **read-only**, with a banner:
  _"Viewing version from Jul 3, 2:14 PM — [Restore this version]"_.
- **Name a version** ("Before Q2 close adjustments") to bookmark an important point.
- **Restore** any version — non-destructive (see §5).
- **Attribution** on every change ("edited by Asha, Ben").

---

## 2. How it's captured (the mechanism)

Yjs represents a sheet as a stream of tiny **binary updates** (one per edit).
Version history is built from two things stored in **PostgreSQL**:

1. **Changeset log** — every Yjs update, appended in order. Fine-grained; can
   reconstruct any point in time.
2. **Snapshots (checkpoints)** — the full encoded Yjs state at a moment. These are
   the entries the user sees as "versions," and they make loading/restoring fast
   (no need to replay thousands of updates).

**Capture pipeline:**

```
edit → Yjs update
   │
   ▼
Hocuspocus persistence hook (debounced)
   ├─► append update  → finsheets_changeset   (continuous, fine-grained)
   └─► checkpoint job → finsheets_snapshot     (periodic "version")
```

**Checkpoint triggers** (a new version is cut when any fires):
- **time-based** — after N minutes of activity (e.g. 10),
- **edit-based** — after M updates (e.g. 200),
- **on idle / last editor leaves** the sheet,
- **on manual "name this version."**

---

## 3. What a "version" records

| Field         | Meaning                                              |
| ------------- | ---------------------------------------------------- |
| `id`          | version id                                           |
| `document_id` | which sheet                                          |
| `seq`         | monotonic order                                      |
| `state_blob`  | full encoded Yjs state (the restorable snapshot)     |
| `label`       | auto (timestamp) or user-given name                  |
| `kind`        | `auto` \| `named` \| `restore` \| `data_refresh`     |
| `authors`     | user ids who contributed since the previous version  |
| `created_at`  | when cut                                             |

---

## 4. Data model (PostgreSQL, per-workspace schema)

- **`finsheets_changeset`** — `id, document_id, seq, update_blob (bytea),
  author_id, created_at`. **Append-only.** The fine-grained truth.
- **`finsheets_snapshot`** *(= a version)* — `id, document_id, seq, state_blob
  (bytea), label, kind, authors (jsonb), created_at`. The checkpoints the user
  sees.
- **`finsheets_document`** — points to the **current head snapshot** for fast open.

> **Redis and IndexedDB are not part of history storage** — Redis = live
> fan-out/speed, IndexedDB = offline buffer. **History lives entirely in
> PostgreSQL.**

---

## 5. Restore semantics (non-destructive)

Restoring version **V** does **not** delete anything:

1. Load **V**'s `state_blob`.
2. Set it as the sheet's new **current state** and broadcast to everyone open
   (Yjs update).
3. Write a **new** version entry `kind: 'restore'`, labelled _"Restored from
   Jul 3, 2:14 PM."_

So the timeline only ever **grows** — you can always undo a restore by restoring
the prior version. Matches Google Docs exactly.

---

## 6. Preview & diff

- **Preview** — click a version → render its snapshot **read-only** in the sheet
  (no live edits while previewing).
- **Diff (Phase 2)** — Yjs snapshots support diffing, so we can highlight
  **changed cells** between two versions ("what changed since yesterday").
  Nice-to-have, not required for v1.

---

## 7. Retention & storage control

To avoid unbounded growth on a busy sheet:

- **Named versions** → kept forever.
- **Auto versions** → pruned on a Google-Docs-like schedule (e.g. keep every
  checkpoint for 24h, then hourly for a week, then daily). Configurable per
  workspace.
- **Changeset compaction** — once a snapshot covers a range, the raw fine-grained
  changesets before it can be compacted/pruned (keep snapshots + a recent tail for
  fine restore). Keeps the table bounded while preserving the version timeline.

Defaults are conservative; all tunable in config.

---

## 8. Attribution & offline

- Each changeset carries `author_id` (from the authenticated Hocuspocus
  connection), so versions correctly show contributors.
- **Offline edits** (buffered in IndexedDB) merge on reconnect and appear in
  history attributed to their author.

> **Open decision (O — offline timestamping):** timestamp offline edits at their
> **original edit time** (needs a trusted client clock) or at **reconnect time**
> (server clock, simpler). **Recommendation: reconnect time for v1** — no
> client-clock trust issues, noting the edit may show slightly later than it
> actually happened.

---

## 9. Permissions & guardrails

- **View history:** anyone who can **view** the document.
  - Flux drill sheets → follow the module rules (`FLUX-ANALYSIS-MODULE.md`):
    module access + account in data scope. Unassigned users can view history
    (they can already view the data + comment).
- **Restore:** requires **edit** rights.
  - Flux drill → assigned Preparer/Reviewer or Owner only. Unassigned users
    **cannot restore** (fail-closed).
- **Name a version:** edit rights (same as restore).
- **Tenant isolation:** versions keyed by **workspace + document**; never
  readable/restorable across workspaces.
- **Audit:** every **restore** and **named version** is audit-logged (who, when,
  which version).
- **License:** pure Yjs (MIT) snapshot/diff — **no** `@univerjs-pro/edit-history-*`.

---

## 10. API surface (`finsheets-service`)

| Method | Path                                             | Does                                              | Requires |
| ------ | ------------------------------------------------ | ------------------------------------------------- | -------- |
| GET    | `/documents/:id/versions`                        | list versions (id, label, kind, authors, time)    | view     |
| GET    | `/documents/:id/versions/:vid`                   | fetch a version's snapshot (preview)              | view     |
| POST   | `/documents/:id/versions`                        | create a named version now                        | edit     |
| POST   | `/documents/:id/versions/:vid/restore`           | non-destructive restore                           | edit     |
| GET    | `/documents/:id/versions/:a/diff/:b` *(Phase 2)* | cell-level diff                                   | view     |

---

## 11. The one real product decision — Flux drill + data refresh (O-5)

A **Flux drill sheet** is generated from NetSuite data. If a user has edited it
collaboratively and then the underlying data refreshes, what should happen?

- **A — New version tagged `data_refresh`:** regenerate the sheet from fresh data
  as a **new version on the same timeline**; user edits before it stay in history.
  _(Simple, one timeline.)_
- **B — Branch:** keep the user's edited sheet as-is; the refresh creates a
  **separate** regenerated sheet. _(Cleaner separation, more entities to manage.)_
- **C — Lock:** once edited, the drill sheet **detaches** from live data and no
  longer auto-refreshes. _(Simplest, but the sheet can go stale.)_

> **Recommendation: A for v1** — one timeline, nothing lost, clearly labelled.
> Genuinely a product decision; reconcile with the drill's read-from-synced-DB
> model.

---

## 12. Open decisions (to finalize)

- **O — Offline timestamping** (§8): original-edit-time vs reconnect-time.
  _Leaning **reconnect time**._
- **O-5 — Flux drill data-refresh behaviour** (§11): A / B / C. _Leaning **A**._

---

## 13. What the FE package (`@levich/univer-sheets`) must expose

So the service can drive this, the FE package will need (future work):

- A **collab binding** — subscribe to Univer mutations (via `onReady` /
  `onCommandExecuted`) → emit Yjs updates; apply remote updates via
  `executeCommand`. (Main correctness risk — see service doc O-1.)
- **Version-history UI hooks** — a right-drawer panel component + props/callbacks
  to `listVersions` / `previewVersion(vid)` / `nameVersion(label)` /
  `restoreVersion(vid)`, so the host wires them to the `finsheets-service` API in
  §10.
- **Read-only preview mode** — render a supplied snapshot without accepting edits
  (for the preview banner state).

These are additive to the current FE package; nothing here ships until the
backend exists.
