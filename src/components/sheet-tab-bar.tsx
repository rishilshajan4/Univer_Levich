/**
 * <SheetTabBar> — a Google-Sheets / Excel-style bottom sheet-tab bar.
 *
 * Replaces Univer's native footer tabs (hidden via `sheetBar={false}`) so the
 * host — not Univer — is the single source of truth for the active sheet. In the
 * lazy shell-workbook product this is essential: Univer only ever holds ONE
 * hydrated sheet, so letting its native tabs switch to an (empty) shell blanks
 * the grid. Here every switch flows through `onSelect`, which re-hydrates.
 *
 * Fully controlled — it owns no sheet state, only menu/rename UI. Behaviours
 * mirror Google Sheets: click to switch, a ▾ dropdown on EVERY visible tab (and
 * right-click) opens the menu (Delete · Duplicate · Copy to ▸ · Rename · Change
 * colour ▸ · Unhide ▸ · Hide · Move), double-click to rename, `+` to add, ☰ for
 * a searchable all-sheets list. Icons are Untitled UI.
 */
import { useEffect, useRef, useState, type CSSProperties } from "react";
import { createPortal } from "react-dom";
import { Z_BASE } from "../core/z-index";
import { ChevronDown, ChevronLeft, ChevronRight, Copy01, Edit01, EyeOff, Grid01, Menu01, Palette, Plus, SearchMd, Trash01, ZoomIn, ZoomOut } from "@untitledui/icons";
import { RenameModal } from "./rename-modal";
import { ConfirmModal } from "./modal";
import { ColorPanel } from "./color-panel";

export interface SheetTabInfo {
  sheetId: string;
  name: string;
  /** #RRGGBB tab colour, or empty/undefined for none. */
  tabColor?: string;
  /** 1 = hidden (no tab; appears in the ☰ list and the Unhide submenu). */
  hidden?: number;
}

export interface SheetTabBarProps {
  sheets: SheetTabInfo[];
  activeSheetId: string | null;
  onSelect: (sheetId: string) => void;
  onAdd?: () => void;
  onRename?: (sheetId: string, name: string) => void;
  onDuplicate?: (sheetId: string) => void;
  onDelete?: (sheetId: string) => void;
  /** Empty string = clear the colour. */
  onChangeColor?: (sheetId: string, color: string) => void;
  onHide?: (sheetId: string) => void;
  onUnhide?: (sheetId: string) => void;
  onMove?: (sheetId: string, direction: -1 | 1) => void;
  onCopyToNew?: (sheetId: string) => void;
  onCopyToExisting?: (sheetId: string) => void;
  /** Current zoom percentage (e.g. 100). When provided with `onZoom`, a bottom-right zoom control is shown. */
  zoom?: number;
  /** Zoom change handler (receives a clamped percentage 50–200). */
  onZoom?: (percent: number) => void;
  /** Toggle worksheet gridlines. When provided, a gridlines button shows next to zoom. */
  onToggleGridlines?: () => void;
  /** Current gridlines visibility (drives the toggle button's on/off look). */
  gridlinesVisible?: boolean;
}

const ZOOM_MIN = 50;
const ZOOM_MAX = 200;
const clampZoom = (p: number) => Math.max(ZOOM_MIN, Math.min(ZOOM_MAX, Math.round(p)));

// FinOpz accent (brand yellow) — the active-tab indicator over the black-and-white base.
const YELLOW = "#EFC71D";
/**
 * Readable text colour for a filled tab: dark text on light fills, light text on
 * dark fills — chosen by whichever of black/white has the higher WCAG contrast
 * ratio against the fill (gamma-correct relative luminance).
 */
function textOn(hex?: string): string {
  const h = (hex ?? "").replace("#", "");
  if (h.length < 6) return "#111827";
  const lin = (i: number) => {
    const c = parseInt(h.slice(i, i + 2), 16) / 255;
    return c <= 0.03928 ? c / 12.92 : Math.pow((c + 0.055) / 1.055, 2.4);
  };
  const L = 0.2126 * lin(0) + 0.7152 * lin(2) + 0.0722 * lin(4);
  const contrastBlack = (L + 0.05) / 0.05; // contrast of black text on this fill
  const contrastWhite = 1.05 / (L + 0.05); // contrast of white text on this fill
  return contrastBlack >= contrastWhite ? "#111827" : "#ffffff";
}

// Anchor for the all-sheets popover (off React state — avoids a re-render loop).
const allAnchor = { x: 8, y: 0 };

const barStyle: CSSProperties = {
  display: "flex", alignItems: "stretch", gap: 2, height: 40, flex: "0 0 40px",
  padding: "0 8px", background: "#f8f9fa", borderTop: "1px solid #e4e7ec",
  fontFamily: "'Work Sans', ui-sans-serif, system-ui, sans-serif", userSelect: "none",
};
const iconBtn: CSSProperties = {
  display: "inline-flex", alignItems: "center", justifyContent: "center", width: 34,
  border: "none", background: "transparent", color: "#5f6368", cursor: "pointer", borderRadius: 6,
};
const panel: CSSProperties = {
  position: "fixed", minWidth: 200, background: "#fff", border: "1px solid #eaecf0", borderRadius: 10,
  boxShadow: "0 12px 32px rgba(16,24,40,0.16)", padding: 6, zIndex: Z_BASE + 4000,
  fontFamily: "'Work Sans', ui-sans-serif, system-ui, sans-serif",
};
const itemStyle: CSSProperties = {
  display: "flex", alignItems: "center", gap: 10, width: "100%", padding: "8px 12px", borderRadius: 6,
  border: "none", background: "transparent", color: "#344054", fontSize: 13, textAlign: "left", cursor: "pointer", whiteSpace: "nowrap",
};
const sepLine: CSSProperties = { height: 1, background: "#eef0f3", margin: "6px 4px" };

function Row({ label, icon, onClick, danger, disabled, children }: { label: string; icon?: React.ReactNode; onClick?: () => void; danger?: boolean; disabled?: boolean; children?: React.ReactNode }) {
  const color = disabled ? "#c3c8d0" : danger ? "#d92d20" : "#344054";
  return (
    <button
      type="button" disabled={disabled} onClick={disabled ? undefined : onClick}
      style={{ ...itemStyle, color, cursor: disabled ? "default" : "pointer", position: "relative" }}
      onMouseEnter={(e) => { if (!disabled) e.currentTarget.style.background = danger ? "#fef3f2" : "#f9fafb"; }}
      onMouseLeave={(e) => (e.currentTarget.style.background = "transparent")}
    >
      {icon && <span style={{ display: "inline-flex", color: disabled ? "#c3c8d0" : danger ? "#d92d20" : "#667085" }}>{icon}</span>}
      <span style={{ flex: 1 }}>{label}</span>
      {children}
    </button>
  );
}

export function SheetTabBar(props: SheetTabBarProps) {
  const { sheets, activeSheetId, onSelect, onAdd, onRename, onDuplicate, onDelete, onChangeColor, onHide, onUnhide, onMove, onCopyToNew, onCopyToExisting, zoom, onZoom, onToggleGridlines, gridlinesVisible } = props;
  const zoomPct = zoom ?? 100;
  const [menu, setMenu] = useState<{ x: number; y: number; sheetId: string } | null>(null);
  const [sub, setSub] = useState<null | "copy" | "color" | "unhide">(null);
  const [allOpen, setAllOpen] = useState(false);
  const [allQuery, setAllQuery] = useState("");
  const [renameReq, setRenameReq] = useState<{ sheetId: string; current: string; taken: Set<string> } | null>(null);
  const [confirmDelete, setConfirmDelete] = useState<{ sheetId: string; name: string } | null>(null);
  const menuRef = useRef<HTMLDivElement>(null);
  const scrollRef = useRef<HTMLDivElement>(null);

  const visible = sheets.filter((s) => !s.hidden);
  const hidden = sheets.filter((s) => s.hidden);
  const visibleCount = visible.length;

  const openMenuForTab = (el: HTMLElement, sheetId: string) => {
    const r = el.getBoundingClientRect();
    setSub(null);
    setMenu({ x: r.left, y: r.top, sheetId });
  };

  useEffect(() => {
    if (!menu && !allOpen) return;
    const onDown = (e: MouseEvent) => {
      const t = e.target as HTMLElement | null;
      if (menuRef.current?.contains(t)) return;
      if (t?.closest?.("[data-tabbar-keep]")) return;
      setMenu(null); setAllOpen(false);
    };
    const onKey = (e: KeyboardEvent) => { if (e.key === "Escape") { setMenu(null); setAllOpen(false); } };
    document.addEventListener("mousedown", onDown, true);
    document.addEventListener("keydown", onKey);
    return () => { document.removeEventListener("mousedown", onDown, true); document.removeEventListener("keydown", onKey); };
  }, [menu, allOpen]);

  const requestRename = (sheetId: string) => {
    const me = sheets.find((s) => s.sheetId === sheetId);
    const taken = new Set(sheets.filter((s) => s.sheetId !== sheetId).map((s) => s.name.toLowerCase()));
    setRenameReq({ sheetId, current: me?.name ?? "", taken });
    setMenu(null);
  };

  // Scroll ONE tab at a time (never jump the strip or move the active tab).
  const scrollOneTab = (dir: -1 | 1) => {
    const el = scrollRef.current;
    if (!el) return;
    const kids = Array.from(el.children) as HTMLElement[];
    const left = el.scrollLeft;
    if (dir > 0) {
      const viewRight = left + el.clientWidth;
      for (const k of kids) if (k.offsetLeft + k.offsetWidth > viewRight + 1) { el.scrollTo({ left: k.offsetLeft, behavior: "smooth" }); return; }
      el.scrollTo({ left: el.scrollWidth, behavior: "smooth" });
    } else {
      for (let i = kids.length - 1; i >= 0; i--) if (kids[i].offsetLeft < left - 1) { el.scrollTo({ left: kids[i].offsetLeft, behavior: "smooth" }); return; }
      el.scrollTo({ left: 0, behavior: "smooth" });
    }
  };

  const menuSheet = menu ? sheets.find((s) => s.sheetId === menu.sheetId) : null;
  const menuIdx = menu ? visible.findIndex((s) => s.sheetId === menu.sheetId) : -1;
  const left = menu ? Math.min(menu.x, window.innerWidth - 230) : 0;
  const bottomPx = menu ? Math.max(8, window.innerHeight - menu.y + 6) : 0;
  const q = allQuery.trim().toLowerCase();
  const allFiltered = sheets.filter((s) => s.name.toLowerCase().includes(q));

  return (
    <>
      <div style={barStyle}>
        <button type="button" title="Add sheet" style={iconBtn} onClick={onAdd}
          onMouseEnter={(e) => (e.currentTarget.style.background = "#eceef1")} onMouseLeave={(e) => (e.currentTarget.style.background = "transparent")}><Plus size={18} /></button>
        <button type="button" title="All sheets" data-tabbar-keep style={iconBtn}
          onClick={(e) => { const r = e.currentTarget.getBoundingClientRect(); setAllOpen((o) => !o); setAllQuery(""); setMenu(null); allAnchor.x = r.left; allAnchor.y = r.top; }}
          onMouseEnter={(e) => (e.currentTarget.style.background = "#eceef1")} onMouseLeave={(e) => (e.currentTarget.style.background = "transparent")}><Menu01 size={18} /></button>
        <div style={{ width: 1, background: "#e4e7ec", margin: "8px 4px" }} />

        <div ref={scrollRef} style={{ display: "flex", alignItems: "stretch", gap: 2, overflowX: "auto", scrollbarWidth: "none", flex: 1 }}>
          {visible.map((s) => {
            const active = s.sheetId === activeSheetId;
            // Excel-style: a coloured sheet fills the ENTIRE tab with its colour. Base is
            // black & white — active tab is a white card, inactive is transparent.
            const colored = !!s.tabColor;
            const bg = colored ? s.tabColor! : active ? "#ffffff" : "transparent";
            const fg = colored ? textOn(s.tabColor) : active ? "#111827" : "#3c4043";
            const chevron = colored ? fg : active ? "#5f6368" : "#98a2b3";
            return (
              <div
                key={s.sheetId}
                data-tabbar-keep
                onClick={() => { if (!active) onSelect(s.sheetId); }}
                onDoubleClick={() => requestRename(s.sheetId)}
                // Right-click opens the menu for THIS tab WITHOUT switching to it —
                // you can act on sheet C while staying on sheet A.
                onContextMenu={(e) => { e.preventDefault(); openMenuForTab(e.currentTarget as HTMLElement, s.sheetId); }}
                title={s.name}
                style={{
                  display: "flex", alignItems: "center", gap: 4, padding: "0 8px 0 12px", cursor: "pointer", position: "relative",
                  maxWidth: 220, borderTopLeftRadius: 8, borderTopRightRadius: 8,
                  background: bg,
                  color: fg,
                  fontSize: 13, fontWeight: active ? 600 : 500,
                  boxShadow: active ? "1px 0 0 #e4e7ec, -1px 0 0 #e4e7ec" : "none",
                }}
              >
                <span style={{ overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{s.name}</span>
                {/* ▾ dropdown trigger on EVERY visible tab */}
                <span
                  role="button" aria-label="Sheet options" title="Sheet options"
                  // Opens THIS tab's menu without switching to it (stay on the current sheet).
                  onClick={(e) => { e.stopPropagation(); openMenuForTab(e.currentTarget.parentElement as HTMLElement, s.sheetId); }}
                  style={{ display: "inline-flex", alignItems: "center", padding: "3px 1px", color: chevron, borderRadius: 4 }}
                >
                  <ChevronDown size={14} />
                </span>
                {/* Active tab indicator — a FinOpz-yellow bar along the bottom of the tab. */}
                {active && <span aria-hidden style={{ position: "absolute", left: 0, right: 0, bottom: 0, height: 3, background: YELLOW, borderRadius: "3px 3px 0 0" }} />}
              </div>
            );
          })}
        </div>

        <div style={{ display: "flex", alignItems: "center" }}>
          <button type="button" title="Previous tab" style={iconBtn} onClick={() => scrollOneTab(-1)}><ChevronLeft size={18} /></button>
          <button type="button" title="Next tab" style={iconBtn} onClick={() => scrollOneTab(1)}><ChevronRight size={18} /></button>
        </div>

        {/* Gridlines on/off — next to zoom. */}
        {onToggleGridlines && (
          <div style={{ display: "flex", alignItems: "center", paddingLeft: 8, marginLeft: 4, borderLeft: "1px solid #e4e7ec" }}>
            <button
              type="button"
              title={gridlinesVisible === false ? "Show gridlines" : "Hide gridlines"}
              aria-pressed={gridlinesVisible !== false}
              style={{ ...iconBtn, width: 30, color: gridlinesVisible === false ? "#98a2b3" : "#344054" }}
              onClick={onToggleGridlines}
              onMouseEnter={(e) => (e.currentTarget.style.background = "#eceef1")}
              onMouseLeave={(e) => (e.currentTarget.style.background = "transparent")}
            >
              <Grid01 size={16} />
            </button>
          </div>
        )}

        {/* Zoom control (bottom-right, Google/Excel style) */}
        {onZoom && (
          <div style={{ display: "flex", alignItems: "center", gap: 6, paddingLeft: 8, marginLeft: 4, borderLeft: "1px solid #e4e7ec" }}>
            <button type="button" title="Zoom out" style={{ ...iconBtn, width: 26 }} onClick={() => onZoom(clampZoom(zoomPct - 10))}><ZoomOut size={16} /></button>
            <input
              type="range" min={ZOOM_MIN} max={ZOOM_MAX} step={5} value={zoomPct}
              onChange={(e) => onZoom(clampZoom(Number(e.target.value)))}
              title={`${zoomPct}%`} style={{ width: 96, accentColor: "#0a0a0a", cursor: "pointer" }}
            />
            <button type="button" title="Zoom in" style={{ ...iconBtn, width: 26 }} onClick={() => onZoom(clampZoom(zoomPct + 10))}><ZoomIn size={16} /></button>
            <button type="button" title="Reset to 100%" onClick={() => onZoom(100)}
              style={{ minWidth: 46, height: 26, border: "1px solid #e4e7ec", borderRadius: 6, background: "#fff", color: "#344054", fontSize: 12, fontWeight: 600, cursor: "pointer" }}>{zoomPct}%</button>
          </div>
        )}
      </div>

      {/* Tab context menu */}
      {menu && menuSheet && createPortal(
        <div ref={menuRef} data-tabbar-keep style={{ ...panel, left, bottom: bottomPx }} onContextMenu={(e) => e.preventDefault()}>
          <Row label="Delete" icon={<Trash01 size={16} />} danger disabled={visibleCount <= 1} onClick={() => { const s = menuSheet; setMenu(null); setConfirmDelete({ sheetId: s.sheetId, name: s.name }); }} />
          <Row label="Duplicate" icon={<Copy01 size={16} />} onClick={() => { setMenu(null); onDuplicate?.(menu.sheetId); }} />
          <div style={{ position: "relative" }} onMouseEnter={() => setSub("copy")} onMouseLeave={() => setSub((v) => (v === "copy" ? null : v))}>
            <Row label="Copy to" icon={<Copy01 size={16} />}><ChevronRight size={15} color="#98a2b3" /></Row>
            {sub === "copy" && (
              <div style={{ ...panel, position: "absolute", left: "100%", bottom: 0, minWidth: 200 }}>
                <Row label="New spreadsheet" onClick={() => { setMenu(null); onCopyToNew?.(menu.sheetId); }} />
                <Row label="Existing spreadsheet" onClick={() => { setMenu(null); onCopyToExisting?.(menu.sheetId); }} />
              </div>
            )}
          </div>
          <Row label="Rename" icon={<Edit01 size={16} />} onClick={() => requestRename(menu.sheetId)} />
          <div style={{ position: "relative" }} onMouseEnter={() => setSub("color")} onMouseLeave={() => setSub((v) => (v === "color" ? null : v))}>
            <Row label="Change colour" icon={<Palette size={16} />}><ChevronRight size={15} color="#98a2b3" /></Row>
            {sub === "color" && (
              <div style={{ ...panel, position: "absolute", left: "100%", bottom: 0, width: 220, padding: 12 }}>
                <button
                  type="button"
                  onClick={() => { setMenu(null); onChangeColor?.(menu.sheetId, ""); }}
                  style={{ display: "flex", alignItems: "center", gap: 8, width: "100%", border: "none", background: "transparent", cursor: "pointer", fontSize: 13, color: "#475467", padding: "2px 2px 8px" }}
                >
                  <span style={{ width: 18, height: 18, borderRadius: "50%", border: "1px solid #d0d5dd", display: "inline-flex", alignItems: "center", justifyContent: "center", color: "#98a2b3", fontSize: 12 }}>⦸</span>
                  Reset
                </button>
                <ColorPanel
                  current={menuSheet.tabColor}
                  onPick={(c) => { setMenu(null); onChangeColor?.(menu.sheetId, c); }}
                  onApply={(c) => onChangeColor?.(menu.sheetId, c)}
                />
              </div>
            )}
          </div>
          <div style={sepLine} />
          {hidden.length > 0 && (
            <div style={{ position: "relative" }} onMouseEnter={() => setSub("unhide")} onMouseLeave={() => setSub((v) => (v === "unhide" ? null : v))}>
              <Row label="Unhide sheet" icon={<EyeOff size={16} />}><ChevronRight size={15} color="#98a2b3" /></Row>
              {sub === "unhide" && (
                <div style={{ ...panel, position: "absolute", left: "100%", bottom: 0, minWidth: 200, maxHeight: 6 * 38, overflowY: "auto" }}>
                  {hidden.map((s) => <Row key={s.sheetId} label={s.name} onClick={() => { setMenu(null); onUnhide?.(s.sheetId); onSelect(s.sheetId); }} />)}
                </div>
              )}
            </div>
          )}
          <Row label="Hide sheet" icon={<EyeOff size={16} />} disabled={visibleCount <= 1} onClick={() => { setMenu(null); onHide?.(menu.sheetId); }} />
          <div style={sepLine} />
          <Row label="Move right" disabled={menuIdx < 0 || menuIdx >= visibleCount - 1} onClick={() => { setMenu(null); onMove?.(menu.sheetId, 1); }} />
          <Row label="Move left" disabled={menuIdx <= 0} onClick={() => { setMenu(null); onMove?.(menu.sheetId, -1); }} />
        </div>,
        document.body,
      )}

      {/* All-sheets list (☰) — search + ~6 rows then scroll; hidden sheets show EyeOff + unhide */}
      {allOpen && createPortal(
        <div ref={menuRef} data-tabbar-keep style={{ ...panel, left: allAnchor.x, bottom: Math.max(8, window.innerHeight - allAnchor.y + 6), width: 264, padding: 8 }}>
          <div style={{ display: "flex", alignItems: "center", gap: 8, padding: "6px 10px", margin: "2px 2px 8px", border: "1px solid #e4e7ec", borderRadius: 8, color: "#667085" }}>
            <SearchMd size={16} />
            <input autoFocus value={allQuery} onChange={(e) => setAllQuery(e.target.value)} placeholder="Search sheets"
              style={{ border: "none", outline: "none", flex: 1, fontSize: 13, color: "#101828", background: "transparent", fontFamily: "inherit" }} />
          </div>
          <div style={{ maxHeight: 6 * 38, overflowY: "auto" }}>
            {allFiltered.map((s) => (
              <Row
                key={s.sheetId}
                label={s.name}
                // Leading tab-colour dot (Google-style), or a transparent 8px slot so names stay aligned.
                icon={<span style={{ width: 8, height: 8, borderRadius: "50%", background: s.tabColor || "transparent", display: "inline-block", flexShrink: 0 }} />}
                onClick={() => { setAllOpen(false); setAllQuery(""); if (s.hidden) onUnhide?.(s.sheetId); onSelect(s.sheetId); }}
              >
                {s.sheetId === activeSheetId
                  ? <span style={{ color: "#12b76a", fontWeight: 700 }}>✓</span>
                  : s.hidden ? <span style={{ color: "#98a2b3", display: "inline-flex" }} title="Hidden (click to unhide)"><EyeOff size={15} /></span> : null}
              </Row>
            ))}
            {allFiltered.length === 0 && <div style={{ padding: "10px 12px", color: "#98a2b3", fontSize: 13 }}>No sheets found</div>}
          </div>
        </div>,
        document.body,
      )}

      {renameReq && createPortal(
        <RenameModal open title="Rename sheet" current={renameReq.current} taken={renameReq.taken}
          onCancel={() => setRenameReq(null)}
          onRename={(name) => { const req = renameReq; setRenameReq(null); onRename?.(req.sheetId, name); }} />,
        document.body,
      )}

      <ConfirmModal
        open={!!confirmDelete}
        title="Delete sheet?"
        message={confirmDelete ? `"${confirmDelete.name}" will be permanently deleted. This can't be undone.` : ""}
        confirmLabel="Delete"
        cancelLabel="Cancel"
        onConfirm={() => { if (confirmDelete) onDelete?.(confirmDelete.sheetId); }}
        onClose={() => setConfirmDelete(null)}
      />
    </>
  );
}

export default SheetTabBar;
