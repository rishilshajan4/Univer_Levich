/**
 * Google-Sheets-style right-click menu for the sheet tabs (Delete · Duplicate ·
 * Copy to ▸ · Rename · Hide sheet). Univer's native footer tab menu can't host
 * confirmations or cross-spreadsheet copy, so we intercept the contextmenu on
 * its tabs ([data-u-comp=slide-tab-item], identified by their label text) and
 * render our own menu that drives the same Facade.
 */
import { useEffect, useRef, useState, type CSSProperties } from "react";
import { createPortal } from "react-dom";
import { stashSnapshotPayload } from "../core/import-data";
import { RenameModal } from "../components/rename-modal";

/* ---- Loose Facade views --------------------------------------------------- */
interface FSheet {
  getSheetName?: () => string;
  getSheetId?: () => string;
  setName?: (name: string) => unknown;
  hideSheet?: () => unknown;
  isSheetHidden?: () => boolean;
  activate?: () => unknown;
  setTabColor?: (color: string) => unknown;
}
interface FWorkbook {
  getSheetByName?: (name: string) => FSheet | null;
  getSheetBySheetId?: (id: string) => FSheet | null;
  getActiveSheet?: () => FSheet | null;
  getSheets?: () => FSheet[];
  deleteSheet?: (sheet: FSheet | string) => unknown;
  duplicateSheet?: (sheet: FSheet) => FSheet | null;
  moveSheet?: (sheet: FSheet, index: number) => unknown;
  getSnapshot?: () => { sheetOrder?: string[]; sheets?: Record<string, unknown> } & Record<string, unknown>;
}

// Google-Sheets tab colour palette (compact).
const TAB_COLORS = ["#000000", "#ea4335", "#ff9900", "#fbbc04", "#34a853", "#4285f4", "#a142f4", "#f439a0", "#674ea7", "#999999"];
interface TabMenuApi {
  getActiveWorkbook?: () => FWorkbook | null;
}

export interface SheetTabMenuProps {
  api: unknown;
  /**
   * File ▸ tab ▸ Copy to ▸ Existing spreadsheet. Cross-spreadsheet copy is owned
   * by the host/backend (it must pick a target doc). Receives the sheet's raw
   * IWorksheetData-ish snapshot + name. Localhost has no backend, so the default
   * just logs; production wires this to FinOpz BE.
   */
  onCopyToExisting?: (sheetName: string, sheetSnapshot: unknown) => void;
}

const panel: CSSProperties = {
  position: "fixed",
  minWidth: 190,
  background: "#fff",
  border: "1px solid #eaecf0",
  borderRadius: 10,
  boxShadow: "0 12px 32px rgba(16,24,40,0.16)",
  padding: 6,
  zIndex: 4000,
  fontFamily: "'Work Sans', ui-sans-serif, system-ui, sans-serif",
};
const item: CSSProperties = {
  display: "flex",
  alignItems: "center",
  gap: 8,
  width: "100%",
  padding: "8px 12px",
  borderRadius: 6,
  border: "none",
  background: "transparent",
  color: "#344054",
  fontSize: 13,
  textAlign: "left",
  cursor: "pointer",
  whiteSpace: "nowrap",
};
const sepLine: CSSProperties = { height: 1, background: "#eef0f3", margin: "6px 4px" };

function Row({ label, onClick, danger, disabled, children }: { label: string; onClick?: () => void; danger?: boolean; disabled?: boolean; children?: React.ReactNode }) {
  const color = disabled ? "#c3c8d0" : danger ? "#d92d20" : (item.color as string);
  return (
    <button
      type="button"
      disabled={disabled}
      onClick={disabled ? undefined : onClick}
      style={{ ...item, color, cursor: disabled ? "default" : "pointer", position: "relative" }}
      onMouseEnter={(e) => { if (!disabled) e.currentTarget.style.background = danger ? "#fef3f2" : "#f9fafb"; }}
      onMouseLeave={(e) => (e.currentTarget.style.background = "transparent")}
    >
      {label}
      {children}
    </button>
  );
}

export function SheetTabMenu({ api, onCopyToExisting }: SheetTabMenuProps) {
  const [menu, setMenu] = useState<{ x: number; y: number; name: string; sheetId: string } | null>(null);
  const [copyOpen, setCopyOpen] = useState(false);
  const [colorOpen, setColorOpen] = useState(false);
  // Rename modal request — held OUTSIDE the menu so it survives the menu closing
  // (the menu unmounts on `menu = null`). Carries the target sheet id/name and
  // the other sheets' names to block duplicates.
  const [renameReq, setRenameReq] = useState<{ sheetId: string; current: string; taken: Set<string> } | null>(null);
  const ref = useRef<HTMLDivElement>(null);
  // Keep the latest api for the document-level delegated handlers (which are
  // registered once and would otherwise close over the initial null api).
  const apiRef = useRef(api);
  apiRef.current = api;

  // The sheet name shown on a tab (excluding our injected caret glyph).
  const tabName = (tab: HTMLElement): string => {
    const span = (tab.querySelector("span:not([data-levich-caret])") as HTMLElement | null) ?? (tab.querySelector("span") as HTMLElement | null);
    return (span?.textContent ?? tab.textContent ?? "").replace(/[▾▴]/g, "").trim();
  };

  // Intercept right-clicks on Univer's sheet tabs (capture phase, so we win over
  // Univer's own contextmenu). The sheet is identified by the tab's label text.
  useEffect(() => {
    const onCtx = (e: MouseEvent) => {
      const el = e.target as HTMLElement | null;
      const tab = el?.closest?.("[data-u-comp=slide-tab-item]") as HTMLElement | null;
      if (!tab) return;
      const sheetId = tab.getAttribute("data-id") ?? "";
      const w = (apiRef.current as TabMenuApi | null)?.getActiveWorkbook?.() ?? null;
      const name = tabName(tab) || w?.getSheetBySheetId?.(sheetId)?.getSheetName?.() || w?.getActiveSheet?.()?.getSheetName?.() || "";
      if (!name) return;
      e.preventDefault();
      e.stopPropagation();
      setCopyOpen(false);
      setMenu({ x: e.clientX, y: e.clientY, name, sheetId });
    };
    document.addEventListener("contextmenu", onCtx, true);
    return () => document.removeEventListener("contextmenu", onCtx, true);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  // Open the menu for a given tab (resolve sheet id + name, position above the
  // caret, toggle). Held in a ref so the injected per-caret listeners always
  // call the latest version without re-binding.
  const openForTabRef = useRef<(tab: HTMLElement, caret: HTMLElement) => void>(() => {});
  openForTabRef.current = (tab, caret) => {
    const w = (apiRef.current as TabMenuApi | null)?.getActiveWorkbook?.() ?? null;
    const sheetId = tab.getAttribute("data-id") ?? "";
    let name = tabName(tab);
    if (!name || !w?.getSheetByName?.(name)) {
      name = w?.getSheetBySheetId?.(sheetId)?.getSheetName?.() || w?.getActiveSheet?.()?.getSheetName?.() || name;
    }
    if (!name) {
      console.warn("[levich] sheet-tab menu: couldn't resolve a sheet name");
      return;
    }
    try {
      (w?.getSheetBySheetId?.(sheetId) ?? w?.getSheetByName?.(name))?.activate?.();
    } catch {
      /* ignore */
    }
    const r = caret.getBoundingClientRect();
    setCopyOpen(false);
    setColorOpen(false);
    setMenu((m) => (m && m.sheetId === sheetId ? null : { x: r.left, y: r.top, name, sheetId })); // toggle
  };

  // Inject a ▴ caret into every sheet tab, each with its OWN click listener that
  // opens the dropdown. Univer re-renders the tab DOM, so a MutationObserver
  // re-adds carets (with fresh listeners) as tabs are added / renamed / moved.
  // Direct per-element listeners proved far more reliable here than document
  // delegation.
  useEffect(() => {
    const decorate = () => {
      document.querySelectorAll("[data-u-comp=slide-tab-item]").forEach((node) => {
        const tab = node as HTMLElement;
        if (tab.querySelector("[data-levich-caret]")) return;
        if (getComputedStyle(tab).position === "static") tab.style.position = "relative";
        const caret = document.createElement("span");
        caret.setAttribute("data-levich-caret", "1");
        caret.setAttribute("role", "button");
        caret.setAttribute("aria-label", "Sheet options");
        caret.textContent = "▴"; // up arrow — the menu opens upward
        Object.assign(caret.style, {
          display: "inline-flex",
          alignItems: "center",
          marginLeft: "4px",
          padding: "0 3px",
          color: "#667085",
          fontSize: "11px",
          lineHeight: "1",
          cursor: "pointer",
          borderRadius: "4px",
        } as CSSStyleDeclaration);
        // Stop the mousedown from activating/dragging the tab; open on click.
        caret.addEventListener("mousedown", (ev) => { ev.preventDefault(); ev.stopPropagation(); });
        caret.addEventListener("click", (ev) => {
          ev.preventDefault();
          ev.stopPropagation();
          openForTabRef.current(tab, caret);
        });
        tab.appendChild(caret);
      });
    };
    decorate();
    const container = document.querySelector("[data-u-comp=sheet-bar-tabs]") ?? document.body;
    const obs = new MutationObserver(decorate);
    obs.observe(container, { childList: true, subtree: true });
    return () => obs.disconnect();
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  // Dismiss on outside click / Escape / scroll.
  useEffect(() => {
    if (!menu) return;
    const close = () => setMenu(null);
    const onDown = (e: MouseEvent) => {
      const t = e.target as HTMLElement | null;
      if (ref.current && ref.current.contains(t)) return; // click inside the menu
      if (t?.closest?.("[data-levich-caret]")) return; // caret handled by delegation (toggle)
      close();
    };
    const onKey = (e: KeyboardEvent) => e.key === "Escape" && close();
    document.addEventListener("mousedown", onDown, true);
    document.addEventListener("keydown", onKey);
    return () => {
      document.removeEventListener("mousedown", onDown, true);
      document.removeEventListener("keydown", onKey);
    };
  }, [menu]);

  // When the context menu is closed we still render the Rename modal if one was
  // requested from it (the modal outlives the menu).
  if (!menu) {
    if (!renameReq) return null;
    const applyTabRename = (name: string) => {
      const req = renameReq;
      setRenameReq(null);
      try {
        const w = (api as TabMenuApi | null)?.getActiveWorkbook?.() ?? null;
        const t = w?.getSheetBySheetId?.(req.sheetId) ?? w?.getSheetByName?.(req.current) ?? w?.getActiveSheet?.() ?? null;
        t?.setName?.(name);
      } catch (e) {
        console.warn("[levich] rename failed", e);
      }
    };
    return createPortal(
      <RenameModal
        open
        title="Rename sheet"
        current={renameReq.current}
        taken={renameReq.taken}
        onCancel={() => setRenameReq(null)}
        onRename={applyTabRename}
      />,
      document.body,
    );
  }

  const wb = () => (api as TabMenuApi | null)?.getActiveWorkbook?.() ?? null;
  // Resolve by the STABLE sheet id first (from the tab's data-id), then by name,
  // then the active sheet — so operations hit the right sheet even when the tab
  // label couldn't be read.
  const target = () => {
    const w = wb();
    return (menu.sheetId ? w?.getSheetBySheetId?.(menu.sheetId) : null) ?? w?.getSheetByName?.(menu.name) ?? w?.getActiveSheet?.() ?? null;
  };
  const close = () => setMenu(null);

  const del = () => {
    const w = wb();
    const visible = (w?.getSheets?.() ?? []).filter((s) => !s.isSheetHidden?.());
    if (visible.length <= 1) {
      window.alert("You can't delete the only visible sheet.");
      return close();
    }
    if (!window.confirm(`Delete "${menu.name}"? This can't be undone.`)) return close();
    try {
      const t = target();
      if (t) w?.deleteSheet?.(t);
    } catch {
      /* ignore */
    }
    close();
  };
  const duplicate = () => {
    if (!window.confirm(`Make a duplicate of "${menu.name}"?`)) return close();
    try {
      const t = target();
      console.info("[levich] duplicate →", { sheetId: menu.sheetId, resolved: !!t });
      if (t) wb()?.duplicateSheet?.(t);
    } catch (e) {
      console.warn("[levich] duplicate failed", e);
    }
    close();
  };
  const rename = () => {
    const t = target();
    const current = t?.getSheetName?.() ?? menu.name;
    const myId = t?.getSheetId?.() ?? menu.sheetId;
    // Names of the OTHER sheets (lower-cased) so the modal can block duplicates.
    const taken = new Set(
      (wb()?.getSheets?.() ?? [])
        .filter((s) => s.getSheetId?.() !== myId)
        .map((s) => (s.getSheetName?.() ?? "").toLowerCase())
        .filter(Boolean),
    );
    setRenameReq({ sheetId: myId, current, taken });
    close(); // close the context menu; the modal is independent state
  };
  const hide = () => {
    const visible = (wb()?.getSheets?.() ?? []).filter((s) => !s.isSheetHidden?.());
    if (visible.length <= 1) {
      window.alert("You can't hide the only visible sheet.");
      return close();
    }
    try {
      target()?.hideSheet?.();
    } catch {
      /* ignore */
    }
    close();
  };
  // Build a single-sheet snapshot of the target sheet.
  const singleSheetSnapshot = (): { snap: Record<string, unknown> | null } => {
    try {
      const full = wb()?.getSnapshot?.();
      const sid = target()?.getSheetId?.();
      if (!full || !sid || !full.sheets?.[sid]) return { snap: null };
      return { snap: { ...full, sheetOrder: [sid], sheets: { [sid]: full.sheets[sid] } } };
    } catch {
      return { snap: null };
    }
  };
  const copyToNew = () => {
    const { snap } = singleSheetSnapshot();
    if (snap && stashSnapshotPayload(snap) && typeof window !== "undefined") {
      window.open(window.location.href, "_blank");
    }
    close();
  };
  const copyToExisting = () => {
    const { snap } = singleSheetSnapshot();
    if (onCopyToExisting) onCopyToExisting(menu.name, snap);
    else console.info(`[levich] Copy "${menu.name}" to an existing spreadsheet → routes to FinOpz BE (no-op on localhost)`);
    close();
  };
  const changeColor = (c: string) => {
    try {
      target()?.setTabColor?.(c);
    } catch {
      /* ignore */
    }
    close();
  };
  const moveBy = (delta: number) => {
    try {
      const w = wb();
      const sheets = w?.getSheets?.() ?? [];
      const idx = sheets.findIndex((s) => (menu.sheetId ? s.getSheetId?.() === menu.sheetId : s.getSheetName?.() === menu.name));
      const t = target();
      if (w && t && idx >= 0) {
        const next = Math.max(0, Math.min(sheets.length - 1, idx + delta));
        if (next !== idx) w.moveSheet?.(t, next);
      }
    } catch {
      /* ignore */
    }
    close();
  };

  // The caret sits at the very bottom, so the menu ALWAYS opens upward: anchor
  // its BOTTOM just above the caret's top (menu.y).
  const left = Math.min(menu.x, window.innerWidth - 220);
  const bottomPx = Math.max(8, window.innerHeight - menu.y + 6);

  // Google-style enabled/disabled state: with a single sheet, Delete / Hide /
  // Move are shown but GREYED (a workbook must keep ≥1 visible sheet); Move is
  // also greyed at the ends. Duplicate / Copy to / Rename / Change colour are
  // always available.
  const allSheets = wb()?.getSheets?.() ?? [];
  const visibleCount = allSheets.filter((s) => !s.isSheetHidden?.()).length;
  const idx = allSheets.findIndex((s) => (menu.sheetId ? s.getSheetId?.() === menu.sheetId : s.getSheetName?.() === menu.name));
  const canDelete = visibleCount > 1;
  const canHide = visibleCount > 1;
  const canMoveLeft = idx > 0;
  const canMoveRight = idx >= 0 && idx < allSheets.length - 1;

  return createPortal(
    <div ref={ref} style={{ ...panel, left, bottom: bottomPx }} onContextMenu={(e) => e.preventDefault()}>
      <Row label="Delete" danger disabled={!canDelete} onClick={del} />
      <Row label="Duplicate" onClick={duplicate} />
      <div style={{ position: "relative" }} onMouseEnter={() => { setCopyOpen(true); setColorOpen(false); }} onMouseLeave={() => setCopyOpen(false)}>
        <Row label="Copy to">
          <span style={{ marginLeft: "auto", color: "#98a2b3" }}>▸</span>
        </Row>
        {copyOpen && (
          <div style={{ ...panel, left: "100%", bottom: 0, minWidth: 200 }}>
            <Row label="New spreadsheet" onClick={copyToNew} />
            <Row label="Existing spreadsheet" onClick={copyToExisting} />
          </div>
        )}
      </div>
      <Row label="Rename" onClick={rename} />
      <div style={{ position: "relative" }} onMouseEnter={() => { setColorOpen(true); setCopyOpen(false); }} onMouseLeave={() => setColorOpen(false)}>
        <Row label="Change colour">
          <span style={{ marginLeft: "auto", color: "#98a2b3" }}>▸</span>
        </Row>
        {colorOpen && (
          <div style={{ ...panel, left: "100%", bottom: 0, minWidth: 168, padding: 10 }}>
            <div style={{ display: "grid", gridTemplateColumns: "repeat(5, 1fr)", gap: 7 }}>
              <button type="button" title="None" onClick={() => changeColor("")} style={{ width: 20, height: 20, borderRadius: "50%", border: "1px solid #d0d5dd", background: "#fff", cursor: "pointer", color: "#98a2b3", fontSize: 12, lineHeight: 1 }}>⦸</button>
              {TAB_COLORS.map((c) => (
                <button key={c} type="button" title={c} onClick={() => changeColor(c)} style={{ width: 20, height: 20, borderRadius: "50%", border: "1px solid rgba(0,0,0,.12)", background: c, cursor: "pointer", padding: 0 }} />
              ))}
            </div>
          </div>
        )}
      </div>
      <div style={sepLine} />
      <Row label="Hide sheet" disabled={!canHide} onClick={hide} />
      <div style={sepLine} />
      <Row label="Move right" disabled={!canMoveRight} onClick={() => moveBy(1)} />
      <Row label="Move left" disabled={!canMoveLeft} onClick={() => moveBy(-1)} />
    </div>,
    document.body,
  );
}
