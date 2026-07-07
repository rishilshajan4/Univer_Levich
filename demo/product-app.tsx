/**
 * FinSheets product view (localhost:9100 — the default route).
 *
 * The real product, not a PoC: a large workbook opens instantly and its sheets
 * load on demand. It hides Univer's native footer tabs and renders our own
 * Google-Sheets-style <SheetTabBar>, which is the SOLE controller of the active
 * sheet — so Univer only ever holds one hydrated sheet and can never blank the
 * grid by switching to an empty shell.
 *
 * Multi-sheet STRUCTURE (order, names, colours, hidden, add/duplicate/delete) is
 * owned by the manifest here — Univer just renders the one active sheet. Cell
 * edits are captured back into the cache on every switch so nothing is lost.
 *
 * "Backend" = static JSON in /public (scripts/xlsx-poc.mjs). In production `client`
 * is swapped for finsheets-service (manifest / sheets / edit endpoints).
 */
import { useCallback, useEffect, useMemo, useRef, useState } from "react";
import { ClockRewind } from "@untitledui/icons";
import {
  LevichSheet, buildShellWorkbook, SheetTabBar, stashSnapshotPayload, VersionHistoryDrawer, RenameModal, createVersionStore,
  diffSheet, highlightSnapshot,
  type ColumnDef, type LevichSheetHandle, type SheetData, type SheetManifestEntry, type SingleSheetSnapshot, type SheetTabInfo,
  type Version, type VersionKind, type DocumentSnapshot,
} from "../src";
import "../src/theme/office-fonts.css"; // alias Calibri→Carlito etc. so text isn't serif

const NO_DATA: SheetData = [];
const NO_COLUMNS: ColumnDef[] = [];

const client = {
  getManifest: () => fetch("/poc-manifest.json").then((r) => r.json() as Promise<SheetManifestEntry[]>),
  getSheet: (sheetId: string) => fetch(`/poc-sheets/${sheetId}.json`).then((r) => r.json() as Promise<SingleSheetSnapshot>),
};

// Default: a BLANK spreadsheet (one empty sheet) — a clean slate to add/import into.
// `?sample` loads the 69-sheet imported workbook for the lazy-load demo.
// `?copy` hydrates from a version stashed by "Make a copy" (see COPY_KEY below).
const params = typeof location !== "undefined" ? new URLSearchParams(location.search) : new URLSearchParams();
const SAMPLE = params.has("sample");
const COPY = params.has("copy");
// Version history is keyed per-document so blank / sample / copy don't share a timeline.
const DOCUMENT_ID = COPY ? `poc-doc-copy-${Date.now().toString(36)}` : SAMPLE ? "poc-doc-sample" : "poc-doc-blank";
const COPY_KEY = "finsheets:version-copy";
// Retention cap for auto-checkpoints (named / original / restore are never pruned).
const MAX_AUTO_VERSIONS = 25;
const store = createVersionStore();
const BLANK_SHEET_ID = "sheet1";
let versionSeqId = 0;
function newVersionId() { return `v_${Date.now().toString(36)}_${versionSeqId++}`; }
/** Deep clone via JSON (snapshots are JSON-serializable). */
function clone<T>(x: T): T { return JSON.parse(JSON.stringify(x)) as T; }
function blankSnapshot(id: string, name: string): SingleSheetSnapshot {
  return { sheets: { [id]: { id, name, rowCount: 200, columnCount: 26, cellData: {} } }, styles: {}, resources: [] };
}

interface LiveSheet { name?: string; hidden?: number; tabColor?: string; cellData?: Record<string, unknown> }
interface LiveSnapshot { sheets?: Record<string, LiveSheet>; styles?: Record<string, unknown>; resources?: Array<{ name: string; data: string }> }
interface WorkbookApi {
  getActiveWorkbook?: () => {
    getSnapshot?: () => LiveSnapshot;
    getSheetBySheetId?: (id: string) => { activate?: () => void } | null;
    getActiveSheet?: () => { zoom?: (ratio: number) => void } | null;
  } | null;
  Event?: Record<string, string>;
  addEvent?: (event: string, cb: (p: { zoom?: number }) => void) => unknown;
}

function resourcesForSheet(resources: Array<{ name: string; data: string }> = [], sheetId: string) {
  return resources.filter((r) => { try { return !!JSON.parse(r.data)?.[sheetId]; } catch { return false; } });
}

/** FinSheets brand logo (FinOpz gold mark + FinSheets wordmark). */
function FinOpzLogo() {
  return (
    <div style={{ display: "inline-flex", alignItems: "center", gap: 6 }} aria-label="FinSheets" role="img">
      <svg width="22" height="22" viewBox="0 0 21 21" fill="none" xmlns="http://www.w3.org/2000/svg" style={{ flexShrink: 0 }}>
        <path d="M20.1172 19.7521C20.1172 19.9541 19.9534 20.1179 19.7514 20.1179H13.7772C13.5752 20.1179 13.4114 19.9541 13.4114 19.7521V13.4121H20.1172V19.7521Z" fill="#EFC71D" />
        <path d="M20.1172 6.70576H13.4114V4.76837e-07H19.7514C19.9534 4.76837e-07 20.1172 0.163761 20.1172 0.365769V6.70576Z" fill="#EFC71D" />
        <rect x="13.4121" y="6.70576" width="6.70576" height="6.70576" transform="rotate(180 13.4121 6.70576)" fill="#EFC71D" />
        <path d="M6.70508 6.70576H0.365085C0.163076 6.70576 -0.000684261 6.542 -0.000684261 6.33999V0.365769C-0.000684261 0.163761 0.163076 4.76837e-07 0.365085 4.76837e-07H6.70508V6.70576Z" fill="#EFC71D" />
        <rect x="20.1172" y="13.4108" width="6.70576" height="6.70576" transform="rotate(180 20.1172 13.4108)" fill="#EFC71D" />
        <path d="M13.4121 13.4108H7.07212C6.87011 13.4108 6.70635 13.2471 6.70635 13.0451V6.70508H13.4121V13.4108Z" fill="#EFC71D" />
        <rect x="6.70508" y="20.1179" width="6.70576" height="6.70576" rx="0.365769" transform="rotate(180 6.70508 20.1179)" fill="#EFC71D" />
      </svg>
      <span style={{ fontFamily: "'Work Sans', ui-sans-serif, system-ui, sans-serif", fontSize: 20, lineHeight: 1, fontWeight: 400, letterSpacing: "-0.6px", color: "#313131" }}>FinSheets</span>
    </div>
  );
}

let idSeq = 0;
function newSheetId() { return `sheet_new_${Date.now().toString(36)}_${idSeq++}`; }

/** Deep-copy a sheet's snapshot under a new id/name, remapping its resources. */
function cloneSnapshot(snap: SingleSheetSnapshot, srcId: string, newId: string, newName: string): SingleSheetSnapshot {
  const copy = JSON.parse(JSON.stringify(snap.sheets[srcId] ?? {}));
  copy.id = newId; copy.name = newName;
  const resources = (snap.resources ?? []).map((r) => {
    try { const d = JSON.parse(r.data); if (d[srcId]) { d[newId] = d[srcId]; delete d[srcId]; } return { name: r.name, data: JSON.stringify(d) }; }
    catch { return r; }
  });
  return { sheets: { [newId]: copy }, styles: snap.styles ?? {}, resources };
}

/** Icon button with an immediate styled tooltip below it (native `title` is slow). */
function IconButtonWithTip({ label, active, onClick, children }: { label: string; active?: boolean; onClick: () => void; children: React.ReactNode }) {
  const [hover, setHover] = useState(false);
  return (
    <div style={{ position: "relative", display: "inline-flex" }} onMouseEnter={() => setHover(true)} onMouseLeave={() => setHover(false)}>
      <button
        type="button" onClick={onClick} aria-label={label}
        style={{ display: "inline-flex", alignItems: "center", justifyContent: "center", width: 36, height: 36, borderRadius: 8, border: "1px solid #d0d5dd", background: active ? "#f2f4f7" : "#fff", color: "#344054", cursor: "pointer" }}
      >
        {children}
      </button>
      {hover && (
        <span
          role="tooltip"
          style={{
            position: "absolute", top: "calc(100% + 6px)", right: 0, whiteSpace: "nowrap", zIndex: 5000,
            background: "#101828", color: "#fff", fontSize: 12, fontWeight: 500, lineHeight: 1, padding: "6px 8px",
            borderRadius: 6, boxShadow: "0 4px 12px rgba(16,24,40,0.18)", pointerEvents: "none",
          }}
        >
          {label}
        </span>
      )}
    </div>
  );
}

export function ProductApp() {
  const ref = useRef<LevichSheetHandle>(null);
  const apiRef = useRef<WorkbookApi | null>(null);
  const cache = useRef(new Map<string, SingleSheetSnapshot>());
  const [manifest, setManifest] = useState<SheetManifestEntry[]>([]);
  const manifestRef = useRef<SheetManifestEntry[]>([]);
  manifestRef.current = manifest;
  const [activeId, setActiveId] = useState<string | null>(null);
  const activeIdRef = useRef<string | null>(null);
  activeIdRef.current = activeId;
  const [activeSnapshot, setActiveSnapshot] = useState<SingleSheetSnapshot | null>(null);
  const [title, setTitle] = useState("Untitled spreadsheet");
  const [, setNote] = useState("Loading…"); // status kept for future use; not shown in the product header
  const [zoomPct, setZoomPct] = useState(100);
  const zoomRef = useRef(100);
  zoomRef.current = zoomPct;
  const renderT0 = useRef(0);

  /* ------------------------------ Version history ------------------------------ */
  const [versions, setVersions] = useState<Version[]>([]);
  const [drawerOpen, setDrawerOpen] = useState(false);
  const [currentVersionId, setCurrentVersionId] = useState<string | null>(null);
  const [preview, setPreview] = useState<Version | null>(null);
  const [previewActiveId, setPreviewActiveId] = useState<string | null>(null);
  const previewRef = useRef<Version | null>(null);
  previewRef.current = preview;
  const seqRef = useRef(0);
  const originalCaptured = useRef(false);
  const autoTimer = useRef<ReturnType<typeof setTimeout> | undefined>(undefined);
  const [nameReq, setNameReq] = useState<Version | null>(null);
  // VH-2 diff ("Highlight changes" / "Show unmodified rows").
  const [highlightChanges, setHighlightChanges] = useState(true);
  const [showUnmodifiedRows, setShowUnmodifiedRows] = useState(true);
  const previewing = !!preview;

  const [fontsReady, setFontsReady] = useState(false);
  useEffect(() => {
    const wanted = ["16px Calibri", "bold 16px Calibri", "16px Arial", '16px "Times New Roman"', "16px Cambria"];
    const done = setTimeout(() => setFontsReady(true), 1500);
    Promise.all(wanted.map((f) => (document.fonts?.load(f) ?? Promise.resolve()).catch(() => {}))).then(() => { clearTimeout(done); setFontsReady(true); });
    return () => clearTimeout(done);
  }, []);

  // Default = blank spreadsheet (one empty sheet). `?sample` loads the 69-sheet
  // workbook. `?copy` hydrates a document stashed by "Make a copy".
  useEffect(() => {
    if (COPY) {
      try {
        const raw = sessionStorage.getItem(COPY_KEY);
        if (raw) {
          const doc = JSON.parse(raw) as DocumentSnapshot;
          for (const [id, snap] of Object.entries(doc.sheets)) cache.current.set(id, snap);
          setManifest(doc.manifest.map((m) => ({ ...m })));
          setTitle(`${doc.title} (copy)`);
          const first = doc.manifest.find((s) => !s.hidden) ?? doc.manifest[0];
          setActiveId(first ? first.sheetId : null);
          setNote("Copied version");
          return;
        }
      } catch { /* fall through to blank */ }
    }
    if (!SAMPLE) {
      cache.current.set(BLANK_SHEET_ID, blankSnapshot(BLANK_SHEET_ID, "Sheet1"));
      setManifest([{ order: 0, sheetId: BLANK_SHEET_ID, name: "Sheet1", hidden: 0 }]);
      setActiveId(BLANK_SHEET_ID);
      setNote("Blank spreadsheet");
      return;
    }
    void (async () => {
      const t0 = performance.now();
      const m = await client.getManifest();
      setManifest(m);
      const first = m.find((s) => !s.hidden && !s.name.includes(">>>")) ?? m.find((s) => !s.hidden) ?? m[0];
      setNote(`${m.length} sheets in ${Math.round(performance.now() - t0)}ms`);
      if (first) setActiveId(first.sheetId);
    })();
  }, []);

  // Hydrate the active sheet (cache-aware).
  useEffect(() => {
    if (!activeId) return;
    const cached = cache.current.get(activeId);
    if (cached) { renderT0.current = performance.now(); setActiveSnapshot(cached); return; }
    const t0 = performance.now();
    void (async () => {
      const snap = await client.getSheet(activeId);
      cache.current.set(activeId, snap);
      renderT0.current = performance.now();
      setNote(`Fetched in ${Math.round(performance.now() - t0)}ms`);
      setActiveSnapshot(snap);
    })();
  }, [activeId]);

  // Persist the active sheet's live edits into the cache (before any switch).
  const captureLive = useCallback(() => {
    const live = apiRef.current?.getActiveWorkbook?.()?.getSnapshot?.();
    const id = activeIdRef.current;
    if (!live?.sheets || !id || !live.sheets[id]) return;
    cache.current.set(id, { sheets: { [id]: live.sheets[id] as Record<string, unknown> }, styles: live.styles ?? {}, resources: resourcesForSheet(live.resources, id) });
  }, []);

  const ensureSnapshot = useCallback(async (id: string): Promise<SingleSheetSnapshot> => {
    const c = cache.current.get(id);
    if (c) return c;
    const snap = await client.getSheet(id);
    cache.current.set(id, snap);
    return snap;
  }, []);

  /* ---- Version-history capture / preview / restore ------------------------- */

  // Assemble a full DocumentSnapshot from the live manifest + every sheet's
  // cached snapshot (fetching any not-yet-visited sheet first, so a checkpoint is
  // the WHOLE document, not just the active sheet). Deep-cloned so later edits
  // never mutate the stored version.
  const assembleDocumentSnapshot = useCallback(async (): Promise<DocumentSnapshot> => {
    captureLive();
    const m = manifestRef.current;
    const sheets: Record<string, SingleSheetSnapshot> = {};
    for (const entry of m) sheets[entry.sheetId] = await ensureSnapshot(entry.sheetId);
    return clone({ manifest: m, sheets, title });
  }, [captureLive, ensureSnapshot, title]);

  const refreshVersions = useCallback(async () => {
    setVersions(await store.listVersions(DOCUMENT_ID));
  }, []);

  const captureVersion = useCallback(async (kind: VersionKind, label?: string, doc?: DocumentSnapshot): Promise<Version> => {
    const document = doc ?? (await assembleDocumentSnapshot());
    const version: Version = {
      id: newVersionId(), seq: seqRef.current++, kind, label,
      author: "You", createdAt: Date.now(), document,
    };
    await store.putVersion(DOCUMENT_ID, version);
    // Retention (H-1): named / original (import|blank) / restore versions are kept
    // forever; auto-checkpoints are capped so a long editing session on a large
    // workbook doesn't grow IndexedDB / heap without bound.
    if (kind === "auto") {
      const autos = (await store.listVersions(DOCUMENT_ID)).filter((v) => v.kind === "auto").sort((a, b) => a.seq - b.seq);
      for (let i = 0; i < autos.length - MAX_AUTO_VERSIONS; i++) await store.deleteVersion(DOCUMENT_ID, autos[i].id);
    }
    await refreshVersions();
    setCurrentVersionId(version.id);
    return version;
  }, [assembleDocumentSnapshot, refreshVersions]);

  // Restore any persisted timeline for this document on mount; else the first
  // load will cut the seq-0 original (see the effect below). `restoreDone` gates
  // that cut so it can't race the (async) restore and duplicate the "blank"
  // original on every page reload.
  const [restoreDone, setRestoreDone] = useState(false);
  useEffect(() => {
    void (async () => {
      const list = await store.listVersions(DOCUMENT_ID);
      if (list.length) {
        setVersions(list);
        setCurrentVersionId(list[list.length - 1].id);
        seqRef.current = Math.max(...list.map((v) => v.seq)) + 1;
        originalCaptured.current = true;
      }
      setRestoreDone(true);
    })();
  }, []);

  // Cut the immutable original (seq 0) once the document has first loaded — but
  // only after the restore check ran, so an existing timeline is never re-cut.
  useEffect(() => {
    if (!restoreDone || originalCaptured.current || !activeId || manifest.length === 0) return;
    originalCaptured.current = true;
    void captureVersion(SAMPLE ? "import" : "blank");
  }, [restoreDone, activeId, manifest, captureVersion]);

  // Debounced auto-checkpoint after a burst of edits (~10s idle). Skipped while
  // previewing a past version (its edits are vetoed anyway).
  const scheduleAutoCheckpoint = useCallback(() => {
    if (previewRef.current) return;
    if (autoTimer.current) clearTimeout(autoTimer.current);
    autoTimer.current = setTimeout(() => { void captureVersion("auto"); }, 10_000);
  }, [captureVersion]);

  const enterPreview = useCallback((v: Version) => {
    const first = v.document.manifest.find((s) => !s.hidden) ?? v.document.manifest[0];
    setPreview(v);
    setPreviewActiveId(first ? first.sheetId : null);
    setDrawerOpen(true);
  }, []);

  const exitPreview = useCallback(() => { setPreview(null); setPreviewActiveId(null); }, []);

  const restoreVersion = useCallback((v: Version) => {
    const doc = clone(v.document);
    cache.current.clear();
    for (const [id, snap] of Object.entries(doc.sheets)) cache.current.set(id, snap);
    setManifest(doc.manifest.map((m) => ({ ...m })));
    setTitle(doc.title);
    const first = doc.manifest.find((s) => !s.hidden) ?? doc.manifest[0];
    exitPreview();
    setActiveSnapshot(null);
    setActiveId(first ? first.sheetId : null);
    // Non-destructive: append a new 'restore' checkpoint from the restored doc.
    void captureVersion("restore", `Restored from ${new Date(v.createdAt).toLocaleString()}`, doc);
  }, [captureVersion, exitPreview]);

  const nameVersion = useCallback(async (v: Version, label: string) => {
    const updated: Version = { ...v, label };
    await store.putVersion(DOCUMENT_ID, updated);
    await refreshVersions();
  }, [refreshVersions]);

  const makeCopy = useCallback((v: Version) => {
    try {
      sessionStorage.setItem(COPY_KEY, JSON.stringify(v.document));
      window.open(`${window.location.pathname}?copy=1`, "_blank");
    } catch { console.warn("[finsheets] Make a copy failed (storage)"); }
  }, []);

  const viewOriginal = useCallback(() => {
    const original = versions.find((v) => v.seq === 0) ?? versions[0];
    if (original) enterPreview(original);
  }, [versions, enterPreview]);

  // Build the combined workbook ONCE per active sheet — Univer renders only the
  // active sheet, so manifest metadata changes (rename/colour/move) must NOT
  // rebuild it. sheetOrder is read from the manifest ref at build time.
  const shellWorkbook = useMemo(() => {
    if (!activeId || !activeSnapshot) return null;
    return buildShellWorkbook({ documentId: DOCUMENT_ID, title, manifest: manifestRef.current, activeSheetId: activeId, activeSnapshot });
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [activeId, activeSnapshot]);

  // Diff the previewed sheet against the PREVIOUS version (Highlight changes).
  const previewDiff = useMemo(() => {
    if (!preview || !previewActiveId) return null;
    const prev = versions.find((v) => v.seq === preview.seq - 1) ?? null;
    return diffSheet(prev?.document ?? null, preview.document, previewActiveId);
  }, [preview, previewActiveId, versions]);

  // The previewed version's shell workbook (read-only), with changed cells
  // highlighted / unmodified rows hidden per the diff toggles.
  const previewWorkbook = useMemo(() => {
    if (!preview || !previewActiveId) return null;
    const doc = preview.document;
    let snap = doc.sheets[previewActiveId];
    if (!snap) return null;
    if (previewDiff && (highlightChanges || !showUnmodifiedRows)) {
      snap = highlightSnapshot(snap, previewActiveId, previewDiff.changed, { highlight: highlightChanges, hideUnchangedRows: !showUnmodifiedRows });
    }
    return buildShellWorkbook({ documentId: DOCUMENT_ID, title: doc.title, manifest: doc.manifest, activeSheetId: previewActiveId, activeSnapshot: snap });
  }, [preview, previewActiveId, previewDiff, highlightChanges, showUnmodifiedRows]);

  // Whichever sheet id is currently on screen (normal OR previewed) — used to
  // activate the right tab after each (re)mount.
  const renderActiveId = previewing ? previewActiveId : activeId;
  const renderActiveRef = useRef<string | null>(null);
  renderActiveRef.current = renderActiveId;

  // Cleanup for the CURRENT mount's event subscriptions + activation timers.
  // LevichSheet remounts on every tab switch / fontsReady toggle, so we dispose
  // the previous mount's before wiring the new one (M-5: no listener/timer leak).
  const mountCleanup = useRef<() => void>(() => {});
  const activateActive = useCallback((api: WorkbookApi) => {
    mountCleanup.current(); // tear down the previous mount's subscriptions/timers
    const timers: ReturnType<typeof setTimeout>[] = [];
    const disposers: Array<() => void> = [];
    const track = (d: unknown) => {
      if (typeof d === "function") disposers.push(d as () => void);
      else if (d && typeof (d as { dispose?: unknown }).dispose === "function") disposers.push(() => (d as { dispose: () => void }).dispose());
    };
    const go = () => {
      try {
        api.getActiveWorkbook?.()?.getSheetBySheetId?.(renderActiveRef.current ?? "")?.activate?.();
        if (zoomRef.current !== 100) api.getActiveWorkbook?.()?.getActiveSheet?.()?.zoom?.(zoomRef.current / 100);
      } catch { /* best-effort */ }
    };
    go(); timers.push(setTimeout(go, 60), setTimeout(go, 200));
    // Keep the bottom-right zoom control in sync when zoom changes elsewhere.
    try { track(api.addEvent?.(api.Event?.SheetZoomChanged ?? "SheetZoomChanged", (p) => { if (typeof p?.zoom === "number") setZoomPct(Math.round(p.zoom * 100)); })); } catch { /* */ }
    // Debounced auto-checkpoint on edit (VH auto versions).
    try { track(api.addEvent?.(api.Event?.SheetEditEnded ?? "SheetEditEnded", () => scheduleAutoCheckpoint())); } catch { /* */ }
    mountCleanup.current = () => {
      timers.forEach(clearTimeout);
      disposers.forEach((d) => { try { d(); } catch { /* */ } });
      mountCleanup.current = () => {};
    };
  }, [scheduleAutoCheckpoint]);

  // Final cleanup on unmount: drop the last mount's subscriptions + the auto timer.
  useEffect(() => () => { mountCleanup.current(); if (autoTimer.current) clearTimeout(autoTimer.current); }, []);

  const handleZoom = useCallback((percent: number) => {
    setZoomPct(percent);
    try { apiRef.current?.getActiveWorkbook?.()?.getActiveSheet?.()?.zoom?.(percent / 100); } catch { /* best-effort */ }
  }, []);

  /* ---- SheetTabBar handlers (manifest + cache; no Univer sheet management) --- */
  const select = useCallback((id: string) => { if (id === activeIdRef.current) return; captureLive(); setActiveId(id); }, [captureLive]);

  // Tab select that is preview-aware: while previewing a past version, switching
  // tabs shows that sheet AT the previewed revision instead of leaving preview.
  const handleTabSelect = useCallback((id: string) => {
    if (previewRef.current) { setPreviewActiveId(id); return; }
    select(id);
  }, [select]);

  const patchManifest = useCallback((fn: (m: SheetManifestEntry[]) => SheetManifestEntry[]) => setManifest((prev) => fn([...prev])), []);

  const addSheet = useCallback(() => {
    const id = newSheetId();
    const name = `Sheet${manifestRef.current.length + 1}`;
    cache.current.set(id, { sheets: { [id]: { id, name, rowCount: 200, columnCount: 26, cellData: {} } }, styles: {}, resources: [] });
    patchManifest((m) => [...m, { order: m.length, sheetId: id, name, hidden: 0 }]);
    captureLive(); setActiveId(id);
  }, [captureLive, patchManifest]);

  const renameSheet = useCallback((id: string, name: string) => {
    patchManifest((m) => m.map((s) => (s.sheetId === id ? { ...s, name } : s)));
    const c = cache.current.get(id); if (c?.sheets[id]) (c.sheets[id] as any).name = name;
  }, [patchManifest]);

  const duplicateSheet = useCallback((id: string) => {
    void (async () => {
      captureLive();
      const src = await ensureSnapshot(id);
      const srcName = manifestRef.current.find((s) => s.sheetId === id)?.name ?? "Sheet";
      const newId = newSheetId();
      const newName = `Copy of ${srcName}`;
      cache.current.set(newId, cloneSnapshot(src, id, newId, newName));
      patchManifest((m) => {
        const i = m.findIndex((s) => s.sheetId === id);
        const entry: SheetManifestEntry = { order: 0, sheetId: newId, name: newName, hidden: 0, tabColor: m[i]?.tabColor };
        const next = [...m]; next.splice(i + 1, 0, entry); return next;
      });
      setActiveId(newId);
    })();
  }, [captureLive, ensureSnapshot, patchManifest]);

  const deleteSheet = useCallback((id: string) => {
    const visible = manifestRef.current.filter((s) => !s.hidden);
    if (visible.length <= 1) return;
    const wasActive = id === activeIdRef.current;
    const idx = visible.findIndex((s) => s.sheetId === id);
    const neighbor = visible[idx + 1] ?? visible[idx - 1];
    cache.current.delete(id);
    patchManifest((m) => m.filter((s) => s.sheetId !== id));
    if (wasActive && neighbor) { captureLive(); setActiveId(neighbor.sheetId); }
  }, [captureLive, patchManifest]);

  const changeColor = useCallback((id: string, color: string) => {
    patchManifest((m) => m.map((s) => (s.sheetId === id ? { ...s, tabColor: color || undefined } : s)));
    const c = cache.current.get(id); if (c?.sheets[id]) (c.sheets[id] as any).tabColor = color || undefined;
  }, [patchManifest]);

  const hideSheet = useCallback((id: string) => {
    const visible = manifestRef.current.filter((s) => !s.hidden);
    if (visible.length <= 1) return;
    const wasActive = id === activeIdRef.current;
    const neighbor = visible.find((s) => s.sheetId !== id);
    patchManifest((m) => m.map((s) => (s.sheetId === id ? { ...s, hidden: 1 } : s)));
    if (wasActive && neighbor) { captureLive(); setActiveId(neighbor.sheetId); }
  }, [captureLive, patchManifest]);

  const unhideSheet = useCallback((id: string) => patchManifest((m) => m.map((s) => (s.sheetId === id ? { ...s, hidden: 0 } : s))), [patchManifest]);

  // Move one position among VISIBLE sheets (swap with the adjacent visible sheet).
  const moveSheet = useCallback((id: string, dir: -1 | 1) => {
    patchManifest((m) => {
      const visIdx = m.map((s, i) => ({ s, i })).filter((x) => !x.s.hidden);
      const pos = visIdx.findIndex((x) => x.s.sheetId === id);
      const target = visIdx[pos + dir];
      if (!target) return m;
      const a = visIdx[pos].i, b = target.i;
      const next = [...m]; [next[a], next[b]] = [next[b], next[a]]; return next;
    });
  }, [patchManifest]);

  const copyToNew = useCallback((id: string) => {
    void (async () => {
      const snap = await ensureSnapshot(id);
      if (stashSnapshotPayload(snap as Record<string, unknown>)) window.open(window.location.href, "_blank");
    })();
  }, [ensureSnapshot]);

  const copyToExisting = useCallback((id: string) => {
    // Production: POST the sheet snapshot to the chosen target document (BE-owned).
    console.info(`[finsheets] Copy "${id}" to an existing spreadsheet → routes to FinOpz BE`);
  }, []);

  const download = async () => {
    captureLive();
    const n = (await ref.current?.exportXlsx(`${title || "finsheets"}.xlsx`)) ?? 0;
    setNote(n ? `Exported ${n} rows` : "Nothing to export");
  };

  const tabs: SheetTabInfo[] = useMemo(
    () => manifest.map((m) => ({ sheetId: m.sheetId, name: m.name, tabColor: m.tabColor, hidden: m.hidden })),
    [manifest],
  );
  const activeName = manifest.find((s) => s.sheetId === activeId)?.name ?? "";
  const hiddenSheetList = useMemo(() => manifest.filter((m) => m.hidden).map((m) => ({ sheetId: m.sheetId, name: m.name })), [manifest]);
  const visibleCount = useMemo(() => manifest.filter((m) => !m.hidden).length, [manifest]);

  // Effective (preview-aware) render inputs: while previewing a past version we
  // render THAT version's tabs/sheets read-only; otherwise the live document.
  const previewTabs: SheetTabInfo[] = useMemo(
    () => (preview ? preview.document.manifest.map((m) => ({ sheetId: m.sheetId, name: m.name, tabColor: m.tabColor, hidden: m.hidden })) : []),
    [preview],
  );
  const effWorkbook = previewing ? previewWorkbook : shellWorkbook;
  const effTabs = previewing ? previewTabs : tabs;
  const effActiveTabId = previewing ? previewActiveId : activeId;

  const toggleDrawer = () => setDrawerOpen((o) => { if (o) exitPreview(); return !o; });

  return (
    <div style={{ display: "flex", flexDirection: "column", height: "100vh", fontFamily: "'Work Sans', system-ui, sans-serif" }}>
      <header style={{ display: "flex", alignItems: "center", gap: 14, padding: "10px 20px", borderBottom: "1px solid #e4e7ec", background: "#fff" }}>
        <FinOpzLogo />
        <div style={{ width: 1, height: 24, background: "#e4e7ec" }} />
        <span style={{ fontSize: 14, fontWeight: 600, color: "#101828" }}>{title}</span>
        {activeName && <span style={{ fontSize: 13, color: "#98a2b3" }}>· {activeName}</span>}
        <div style={{ marginLeft: "auto", display: "flex", alignItems: "center", gap: 10 }}>
          <IconButtonWithTip label="Version history" active={drawerOpen} onClick={toggleDrawer}>
            <ClockRewind size={20} />
          </IconButtonWithTip>
          <button onClick={download} style={{ padding: "8px 16px", borderRadius: 8, border: "none", background: "#101828", color: "#fff", fontSize: 14, fontWeight: 600, cursor: "pointer" }}>Download .xlsx</button>
        </div>
      </header>

      <div style={{ flex: 1, minHeight: 0, display: "flex", overflow: "hidden" }}>
        {/* min-width:0 lets the sheet area shrink so the 320px drawer fits instead of being pushed off-screen. */}
        <div style={{ flex: 1, minWidth: 0, minHeight: 0, display: "flex", flexDirection: "column" }}>
          {previewing && preview && (
            <div style={{ display: "flex", alignItems: "center", gap: 12, padding: "8px 16px", background: "#eff4ff", borderBottom: "1px solid #b2ccff", fontSize: 13, color: "#175cd3" }}>
              <span style={{ fontWeight: 600 }}>{new Date(preview.createdAt).toLocaleString()}</span>
              <span style={{ color: "#3538cd" }}>· read-only preview</span>
              <div style={{ marginLeft: "auto", display: "flex", gap: 8 }}>
                {preview.id !== currentVersionId && (
                  <button onClick={() => restoreVersion(preview)} style={{ padding: "6px 14px", borderRadius: 8, border: "none", background: "#101828", color: "#fff", fontSize: 13, fontWeight: 600, cursor: "pointer" }}>Restore this version</button>
                )}
                <button onClick={exitPreview} style={{ padding: "6px 14px", borderRadius: 8, border: "1px solid #b2ccff", background: "#fff", color: "#175cd3", fontSize: 13, fontWeight: 600, cursor: "pointer" }}>Exit preview</button>
              </div>
            </div>
          )}
          <div style={{ flex: 1, minHeight: 0, position: "relative" }}>
            {previewing && previewDiff?.noChanges && (
              <div style={{ position: "absolute", inset: 0, zIndex: 5, display: "flex", alignItems: "center", justifyContent: "center", pointerEvents: "none" }}>
                <div style={{ background: "rgba(255,255,255,0.92)", border: "1px solid #eaecf0", borderRadius: 12, padding: "18px 26px", color: "#667085", fontSize: 14, fontWeight: 500, boxShadow: "0 8px 24px rgba(16,24,40,0.08)" }}>
                  No changes to this sheet in this revision
                </div>
              </div>
            )}
            {effWorkbook ? (
              <LevichSheet
                key={`${previewing ? `preview:${preview?.id}:${previewActiveId}:${highlightChanges ? 1 : 0}:${showUnmodifiedRows ? 1 : 0}` : activeId}:${fontsReady}`}
                ref={ref}
                data={NO_DATA}
                columns={NO_COLUMNS}
                snapshot={effWorkbook}
                readOnly={previewing}
                sheetBar={false} // hide Univer's native tabs — we render <SheetTabBar>
                onReady={(api) => {
                  apiRef.current = api as unknown as WorkbookApi;
                  activateActive(api as unknown as WorkbookApi);
                  const ms = Math.round(performance.now() - renderT0.current);
                  if (renderT0.current) setNote(`${activeName} rendered in ${ms}ms`);
                }}
                onSave={() => { if (!previewRef.current) void captureVersion("auto"); return true; }}
                onRename={(name) => { setTitle(name); document.title = `${name} - FinOpz Sheets`; }}
                // Sheet-visibility routed through the manifest (same as the tab bar), so
                // View ▸ Hide sheet / Show sheets stay in sync with the tabs.
                onHideActiveSheet={() => { if (activeIdRef.current) hideSheet(activeIdRef.current); }}
                onShowSheet={(id) => { unhideSheet(id); select(id); }}
                hiddenSheetList={hiddenSheetList}
                canHideActiveSheet={visibleCount > 1}
              />
            ) : (
              <div style={{ padding: 40, color: "#667085" }}>Loading {activeName || "workbook"}…</div>
            )}
          </div>

          <SheetTabBar
            sheets={effTabs}
            activeSheetId={effActiveTabId}
            onSelect={handleTabSelect}
            onAdd={previewing ? undefined : addSheet}
            onRename={previewing ? undefined : renameSheet}
            onDuplicate={previewing ? undefined : duplicateSheet}
            onDelete={previewing ? undefined : deleteSheet}
            onChangeColor={previewing ? undefined : changeColor}
            onHide={previewing ? undefined : hideSheet}
            onUnhide={previewing ? undefined : unhideSheet}
            onMove={previewing ? undefined : moveSheet}
            onCopyToNew={previewing ? undefined : copyToNew}
            onCopyToExisting={previewing ? undefined : copyToExisting}
            zoom={zoomPct}
            onZoom={handleZoom}
          />
        </div>

        <VersionHistoryDrawer
          open={drawerOpen}
          versions={versions}
          currentVersionId={currentVersionId}
          previewingId={preview?.id ?? null}
          onClose={() => { setDrawerOpen(false); exitPreview(); }}
          onPreview={enterPreview}
          onRestore={restoreVersion}
          onName={(v) => setNameReq(v)}
          onMakeCopy={makeCopy}
          onViewOriginal={viewOriginal}
          highlightChanges={highlightChanges}
          showUnmodifiedRows={showUnmodifiedRows}
          onToggleHighlight={setHighlightChanges}
          onToggleUnmodified={setShowUnmodifiedRows}
        />
      </div>

      {nameReq && (
        <RenameModal
          open
          title="Name this version"
          current={nameReq.label ?? ""}
          onCancel={() => setNameReq(null)}
          onRename={(name) => { const v = nameReq; setNameReq(null); void nameVersion(v, name); }}
        />
      )}
    </div>
  );
}

export default ProductApp;
