/**
 * FunctionsMenu — the Σ "functions" dropdown (rendered inside the toolbar's
 * Functions Dropdown portal). Mirrors Google's Σ menu:
 *
 *   [ search any function … ]
 *   SUM · AVERAGE · COUNT · MAX · MIN     (quick aggregates, auto-ranged)
 *   ───────────────
 *   Math ▸ · Statistical ▸ · Financial ▸ · Logical ▸ · Lookup ▸ · Text ▸ · …
 *
 * The category list + every function come from Univer's real formula catalog
 * (function-catalog.generated.ts — 500+ functions, all evaluated by the engine).
 * Typing in the search box flattens to matching functions across all categories.
 *
 * The category fly-outs are PORTALED to <body> (and positioned with fixed
 * coords) so the scrollable category list can't clip them, with a short
 * hover-bridge timer so moving the pointer onto the fly-out doesn't dismiss it.
 */
import { useMemo, useRef, useState, type CSSProperties, type ReactNode } from "react";
import { createPortal } from "react-dom";
import { Z_BASE } from "../core/z-index";
import { FULL_FUNCTION_CATEGORIES, type CatalogFn } from "../features/function-catalog.generated";
import type { Aggregate } from "../features/functions";

const INK = "#101828";
const TXT = "#344054";
const HOVER = "#f9fafb";
const FLYOUT_W = 252;

const itemStyle: CSSProperties = {
  display: "flex",
  alignItems: "center",
  gap: 10,
  width: "100%",
  padding: "8px 10px",
  borderRadius: 6,
  border: "none",
  background: "transparent",
  color: TXT,
  fontSize: 13,
  textAlign: "left",
  cursor: "pointer",
  fontFamily: "inherit",
  whiteSpace: "nowrap",
};

const QUICK: Aggregate[] = ["SUM", "AVERAGE", "COUNT", "MAX", "MIN"];
const ALL_FNS: CatalogFn[] = FULL_FUNCTION_CATEGORIES.flatMap((c) => c.fns);

function ChevR() {
  return (
    <svg width="12" height="12" viewBox="0 0 24 24" style={{ marginLeft: "auto" }}>
      <path d="M9 6l6 6-6 6" fill="none" stroke="#98a2b3" strokeWidth="2.2" strokeLinecap="round" strokeLinejoin="round" />
    </svg>
  );
}

function Row({ children, onClick, active }: { children: ReactNode; onClick?: (e: React.MouseEvent) => void; active?: boolean }) {
  return (
    <button
      type="button"
      onMouseDown={(e) => e.preventDefault()}
      onClick={(e) => onClick?.(e)}
      onMouseEnter={(e) => (e.currentTarget.style.background = HOVER)}
      onMouseLeave={(e) => (e.currentTarget.style.background = active ? HOVER : "transparent")}
      style={{ ...itemStyle, background: active ? HOVER : "transparent" }}
    >
      {children}
    </button>
  );
}

function FnRow({ fn, onClick }: { fn: CatalogFn; onClick: () => void }) {
  // Name on top, description stacked beneath it (Google-style), so the hint
  // wraps to full text instead of being clipped on the right.
  return (
    <button
      type="button"
      onMouseDown={(e) => e.preventDefault()}
      onClick={onClick}
      onMouseEnter={(e) => (e.currentTarget.style.background = HOVER)}
      onMouseLeave={(e) => (e.currentTarget.style.background = "transparent")}
      style={{
        display: "flex",
        flexDirection: "column",
        alignItems: "flex-start",
        gap: 2,
        width: "100%",
        padding: "8px 10px",
        borderRadius: 6,
        border: "none",
        background: "transparent",
        cursor: "pointer",
        fontFamily: "inherit",
        textAlign: "left",
      }}
    >
      <span style={{ fontWeight: 600, color: INK, fontSize: 13 }}>{fn.name}</span>
      <span style={{ color: "#98a2b3", fontSize: 12, lineHeight: 1.35, whiteSpace: "normal" }}>{fn.hint}</span>
    </button>
  );
}

export interface FunctionsMenuProps {
  onQuick: (fn: Aggregate) => void;
  onInsert: (name: string) => void;
  close: () => void;
}

export function FunctionsMenu({ onQuick, onInsert, close }: FunctionsMenuProps) {
  const [open, setOpen] = useState<{ cat: string; top: number; left: number } | null>(null);
  const [query, setQuery] = useState("");
  const catListRef = useRef<HTMLDivElement>(null);

  // Click a category to open its fly-out (toggle if it's already open). Every
  // fly-out opens at the SAME top — anchored to the top of the category list
  // (i.e. level with "Math") — regardless of which category was clicked.
  const toggleCat = (cat: string, e: React.MouseEvent) => {
    if (open?.cat === cat) {
      setOpen(null);
      return;
    }
    const r = (e.currentTarget as HTMLElement).getBoundingClientRect();
    const listTop = catListRef.current?.getBoundingClientRect().top ?? r.top;
    const left = Math.max(8, r.left - FLYOUT_W - 2);
    const top = Math.max(8, Math.min(listTop, window.innerHeight - 392));
    setOpen({ cat, top, left });
  };

  const results = useMemo(() => {
    const q = query.trim().toUpperCase();
    if (!q) return null;
    const starts = ALL_FNS.filter((f) => f.name.startsWith(q));
    const contains = ALL_FNS.filter((f) => !f.name.startsWith(q) && (f.name.includes(q) || f.hint.toUpperCase().includes(q)));
    return [...starts, ...contains].slice(0, 60);
  }, [query]);

  const activeCat = open && FULL_FUNCTION_CATEGORIES.find((c) => c.category === open.cat);

  return (
    <div>
      {/* Search */}
      <div style={{ position: "relative", margin: "2px 2px 6px" }}>
        <input
          value={query}
          onChange={(e) => setQuery(e.target.value)}
          onMouseDown={(e) => e.stopPropagation()}
          placeholder="Search 500+ functions"
          style={{ width: "100%", padding: "8px 10px 8px 30px", border: "1px solid #d0d5dd", borderRadius: 8, fontSize: 13, fontFamily: "inherit", outline: "none", boxSizing: "border-box" }}
        />
        <svg width="15" height="15" viewBox="0 0 24 24" style={{ position: "absolute", left: 9, top: 9 }}>
          <circle cx="11" cy="11" r="7" fill="none" stroke="#98a2b3" strokeWidth="2" />
          <path d="M20 20l-3.5-3.5" stroke="#98a2b3" strokeWidth="2" strokeLinecap="round" />
        </svg>
      </div>

      {results ? (
        <div style={{ maxHeight: 320, overflowY: "auto" }}>
          {results.length === 0 && <div style={{ padding: "8px 10px", fontSize: 13, color: "#98a2b3" }}>No matching function</div>}
          {results.map((f) => (
            <FnRow key={f.name} fn={f} onClick={() => { close(); onInsert(f.name); }} />
          ))}
        </div>
      ) : (
        <>
          {QUICK.map((fn) => (
            <Row key={fn} onClick={() => { close(); onQuick(fn); }}>
              <span style={{ fontWeight: 600, color: INK }}>{fn}</span>
            </Row>
          ))}
          <div style={{ height: 1, background: "#eaecf0", margin: "6px 4px" }} />
          <div ref={catListRef} style={{ maxHeight: 300, overflowY: "auto" }}>
            {FULL_FUNCTION_CATEGORIES.map((cat) => (
              <Row key={cat.category} active={open?.cat === cat.category} onClick={(e) => toggleCat(cat.category, e)}>
                {cat.category} <span style={{ color: "#98a2b3", fontSize: 12 }}>{cat.fns.length}</span> <ChevR />
              </Row>
            ))}
          </div>

          {activeCat &&
            createPortal(
              <div
                data-levich-dd
                style={{
                  position: "fixed",
                  top: open.top,
                  left: open.left,
                  width: FLYOUT_W,
                  maxHeight: 384,
                  overflowY: "auto",
                  background: "#fff",
                  border: "1px solid #eaecf0",
                  borderRadius: 10,
                  boxShadow: "0 8px 24px rgba(16,24,40,0.16)",
                  padding: 6,
                  zIndex: Z_BASE + 1200,
                  fontFamily: "'Work Sans', ui-sans-serif, system-ui, sans-serif",
                }}
              >
                {activeCat.fns.map((f) => (
                  <FnRow key={f.name} fn={f} onClick={() => { close(); onInsert(f.name); }} />
                ))}
              </div>,
              document.body,
            )}
        </>
      )}
    </div>
  );
}
