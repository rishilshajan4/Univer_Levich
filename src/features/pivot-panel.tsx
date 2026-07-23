/**
 * <PivotPanel> — an Excel / Google-Sheets-style "PivotTable Fields" drawer.
 *
 * A right-hand pane listing the source fields (searchable, draggable chips) and
 * four drop areas — Filters · Columns · Rows · Values — into which fields are
 * arranged. Dragging uses native HTML5 drag-and-drop (no new dependency). A chip
 * in the Values area carries an aggregation dropdown (Sum/Count/…). Every edit
 * fires `onChange(spec)`.
 *
 * The panel is a thin, controlled view over a `PivotSpec`: all spec transforms
 * live in the exported pure helpers below (unit-testable without a DOM), and the
 * component just wires drag events → those helpers → `onChange`.
 *
 * Styling uses the host's Untitled UI design tokens (CSS custom properties from
 * theme.css, resolved because the panel renders inside the host DOM) via the `T`
 * map below — so buttons/borders/text are on-brand (fiab-yellow) and theme-aware.
 */
import { useEffect, useRef, useState, type CSSProperties, type DragEvent, type ReactNode } from "react";
import { ChevronDown, Plus, X } from "@untitledui/icons";
import type { PivotAggregate, PivotDimSetting, PivotShowAs, PivotSpec, PivotValueField } from "../core/types";

/* ─── Spec model (pure, testable) ─────────────────────────────────────────── */

/** The four drop areas a field can live in. */
export type PivotArea = "filters" | "columns" | "rows" | "values";

export const PIVOT_AREAS: Array<{ key: PivotArea; label: string }> = [
  { key: "filters", label: "Filters" },
  { key: "columns", label: "Columns" },
  { key: "rows", label: "Rows" },
  { key: "values", label: "Values" },
];

/** The 13 Google-Sheets "Summarise by" functions (internal aggregate ← Google label). */
export const PIVOT_AGGREGATES: Array<{ value: PivotAggregate; label: string }> = [
  { value: "sum", label: "SUM" },
  { value: "count", label: "COUNTA" },
  { value: "countNumbers", label: "COUNT" },
  { value: "countunique", label: "COUNTUNIQUE" },
  { value: "average", label: "AVERAGE" },
  { value: "max", label: "MAX" },
  { value: "min", label: "MIN" },
  { value: "median", label: "MEDIAN" },
  { value: "product", label: "PRODUCT" },
  { value: "stdev", label: "STDEV" },
  { value: "stdevp", label: "STDEVP" },
  { value: "var", label: "VAR" },
  { value: "varp", label: "VARP" },
];

/** "Show as" options (Google Sheets). */
export const PIVOT_SHOW_AS: Array<{ value: PivotShowAs; label: string }> = [
  { value: "default", label: "Default" },
  { value: "pctOfRow", label: "% of row" },
  { value: "pctOfCol", label: "% of column" },
  { value: "pctOfGrand", label: "% of grand total" },
];

/** Order options for Rows/Columns. */
const ORDER_OPTS: Array<{ value: "asc" | "desc"; label: string }> = [
  { value: "asc", label: "Ascending" },
  { value: "desc", label: "Descending" },
];

/** The ordered field names in a given area (values → their `field`). */
export function fieldsInArea(spec: PivotSpec, area: PivotArea): string[] {
  switch (area) {
    case "filters":
      return (spec.filters ?? []).map((f) => f.field);
    case "columns":
      return spec.columns;
    case "rows":
      return spec.rows;
    case "values":
      return spec.values.map((v) => v.field);
  }
}

/** Which area a field currently belongs to (first match), or null if unplaced. */
export function areaOfField(spec: PivotSpec, field: string): PivotArea | null {
  for (const { key } of PIVOT_AREAS) {
    if (fieldsInArea(spec, key).includes(field)) return key;
  }
  return null;
}

/** Remove a field from every area, returning a new spec. */
export function removeField(spec: PivotSpec, field: string): PivotSpec {
  return {
    ...spec,
    rows: spec.rows.filter((f) => f !== field),
    columns: spec.columns.filter((f) => f !== field),
    values: spec.values.filter((v) => v.field !== field),
    filters: (spec.filters ?? []).filter((f) => f.field !== field),
  };
}

/**
 * Insert `field` into `area` at `index` (append when index is undefined / out of
 * range). The field is first removed from wherever it was, so moving between
 * areas and reordering within an area are the same operation. When dropping into
 * Values, the field's existing `PivotValueField` (aggregate) is preserved; a new
 * one defaults to "sum".
 */
export function placeField(spec: PivotSpec, field: string, area: PivotArea, index?: number): PivotSpec {
  // Preserve an existing value-field config (its aggregate) across the move.
  const existingValue = spec.values.find((v) => v.field === field);
  const cleared = removeField(spec, field);
  const at = (arr: unknown[]): number => (index === undefined || index < 0 || index > arr.length ? arr.length : index);

  switch (area) {
    case "rows": {
      const rows = [...cleared.rows];
      rows.splice(at(rows), 0, field);
      return { ...cleared, rows };
    }
    case "columns": {
      const columns = [...cleared.columns];
      columns.splice(at(columns), 0, field);
      return { ...cleared, columns };
    }
    case "filters": {
      const filters = [...(cleared.filters ?? [])];
      filters.splice(at(filters), 0, { field });
      return { ...cleared, filters };
    }
    case "values": {
      const values = [...cleared.values];
      const vf: PivotValueField = existingValue ?? { field, aggregate: "sum" };
      values.splice(at(values), 0, vf);
      return { ...cleared, values };
    }
  }
}

/** Change the aggregation of a value field. */
export function setValueAggregate(spec: PivotSpec, field: string, aggregate: PivotAggregate): PivotSpec {
  return {
    ...spec,
    values: spec.values.map((v) => (v.field === field ? { ...v, aggregate } : v)),
  };
}

/** The aggregate currently applied to a value field (for the dropdown). */
export function aggregateOfValue(spec: PivotSpec, field: string): PivotAggregate {
  return spec.values.find((v) => v.field === field)?.aggregate ?? "sum";
}

/** Change the "Show as" mode of a value field. */
export function setValueShowAs(spec: PivotSpec, field: string, showAs: PivotShowAs): PivotSpec {
  return { ...spec, values: spec.values.map((v) => (v.field === field ? { ...v, showAs } : v)) };
}

/** Patch a Rows/Columns field's per-field settings (Order / Show totals), keyed by field. */
export function setDimSetting(spec: PivotSpec, field: string, patch: Partial<PivotDimSetting>): PivotSpec {
  const prev = spec.dimSettings ?? {};
  return { ...spec, dimSettings: { ...prev, [field]: { ...prev[field], ...patch } } };
}
export function dimSettingOf(spec: PivotSpec, field: string): PivotDimSetting {
  return spec.dimSettings?.[field] ?? {};
}

/** Set the kept values ("Filter by values") for a filter field. `null`/[] = keep all. */
export function setFilterInclude(spec: PivotSpec, field: string, include: string[] | null): PivotSpec {
  return {
    ...spec,
    filters: (spec.filters ?? []).map((f) => (f.field === field ? { ...f, include: include && include.length ? include : undefined } : f)),
  };
}
export function filterIncludeOf(spec: PivotSpec, field: string): string[] | null {
  const f = (spec.filters ?? []).find((x) => x.field === field);
  return f?.include ?? null;
}

/* ─── Styles ──────────────────────────────────────────────────────────────── */

// Design-system tokens: the panel renders inside the host app, so these CSS custom
// properties (defined by the host's Untitled UI theme.css) resolve here — keeping the
// pivot editor on-brand (fiab-yellow, not blue) and theme-aware, matching <Button>.
const T = {
  textPrimary: "var(--color-text-primary)",
  textSecondary: "var(--color-text-secondary)",
  textTertiary: "var(--color-text-tertiary)",
  textQuaternary: "var(--color-text-quaternary)",
  fgQuaternary: "var(--color-fg-quaternary)",
  brand: "var(--color-text-brand-primary)", // fiab-yellow-700 — links/accents
  borderPrimary: "var(--color-border-primary)",
  borderSecondary: "var(--color-border-secondary)",
  bg: "var(--color-bg-primary)",
  bgHover: "var(--color-bg-primary_hover)",
} as const;

const drawer: CSSProperties = {
  position: "absolute",
  top: 0,
  right: 0,
  bottom: 0,
  width: 400,
  background: T.bg,
  borderLeft: `1px solid ${T.borderSecondary}`,
  boxShadow: "-8px 0 24px rgba(16,24,40,0.08)",
  display: "flex",
  flexDirection: "column",
  zIndex: 50,
};
const header: CSSProperties = { display: "flex", alignItems: "center", justifyContent: "space-between", padding: "14px 16px 8px", flexShrink: 0 };
const titleStyle: CSSProperties = { fontSize: 15, fontWeight: 600, color: T.textPrimary };
const closeBtn: CSSProperties = { display: "inline-flex", alignItems: "center", justifyContent: "center", width: 28, height: 28, borderRadius: 8, border: "none", background: "transparent", color: T.textTertiary, cursor: "pointer" };
const bodyStyle: CSSProperties = { display: "flex", flexDirection: "column", flex: 1, minHeight: 0, overflowY: "auto", padding: "0 16px 16px" };
const searchInput: CSSProperties = { width: "100%", height: 34, borderRadius: 8, border: `1px solid ${T.borderPrimary}`, padding: "0 12px", fontSize: 13, boxSizing: "border-box", outline: "none", color: T.textPrimary, background: T.bg };
const chipRemove: CSSProperties = { display: "inline-flex", alignItems: "center", justifyContent: "center", width: 15, height: 15, borderRadius: 4, border: "none", background: "transparent", color: T.fgQuaternary, cursor: "pointer", flexShrink: 0, padding: 0 };

const areaTitle: CSSProperties = { fontSize: 11, fontWeight: 600, color: T.textTertiary, textTransform: "uppercase", letterSpacing: 0.3 };
const emptyHint: CSSProperties = { fontSize: 11.5, color: T.textQuaternary, fontStyle: "italic" };

const DND_MIME = "application/x-levich-pivot-field";

const cardStyle: CSSProperties = { border: `1px solid ${T.borderSecondary}`, borderRadius: 8, background: T.bg, padding: 10, display: "flex", flexDirection: "column", gap: 8 };
const cardHead: CSSProperties = { display: "flex", alignItems: "center", justifyContent: "space-between", gap: 6 };
const cardName: CSSProperties = { fontSize: 13, fontWeight: 600, color: T.textPrimary, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap", cursor: "grab" };
const ctrlLabel: CSSProperties = { fontSize: 11, color: T.textTertiary, marginBottom: 3, display: "block" };
const ctrlRow: CSSProperties = { display: "grid", gridTemplateColumns: "1fr 1fr", gap: 8 };
// Secondary <Button> look: white bg, gray ring, neutral text + quaternary icon (no blue).
const addBtn: CSSProperties = { display: "inline-flex", alignItems: "center", gap: 4, height: 24, padding: "0 8px", borderRadius: 6, border: `1px solid ${T.borderPrimary}`, background: T.bg, color: T.textSecondary, fontSize: 12, fontWeight: 600, cursor: "pointer" };
const sectionHead: CSSProperties = { display: "flex", alignItems: "center", justifyContent: "space-between", margin: "14px 0 6px" };
const selectBtn: CSSProperties = { display: "flex", alignItems: "center", justifyContent: "space-between", gap: 6, width: "100%", height: 30, padding: "0 8px", borderRadius: 6, border: `1px solid ${T.borderPrimary}`, background: T.bg, color: T.textPrimary, fontSize: 12.5, cursor: "pointer", boxSizing: "border-box" };
const popover: CSSProperties = { position: "absolute", zIndex: 60, minWidth: 160, maxHeight: 260, overflowY: "auto", background: T.bg, border: `1px solid ${T.borderSecondary}`, borderRadius: 8, boxShadow: "0 8px 24px rgba(16,24,40,0.12)", padding: 4, marginTop: 4 };
const popItem = (active: boolean): CSSProperties => ({ display: "flex", alignItems: "center", gap: 8, padding: "6px 10px", borderRadius: 6, fontSize: 12.5, color: T.textSecondary, cursor: "pointer", background: active ? T.bgHover : "transparent", whiteSpace: "nowrap" });

/** A design-system dropdown (button + popover) — replaces the native <select>. */
function Select<T extends string>({ value, options, onChange, ariaLabel }: { value: T; options: Array<{ value: T; label: string }>; onChange: (v: T) => void; ariaLabel: string }) {
  const [open, setOpen] = useState(false);
  const ref = useRef<HTMLDivElement | null>(null);
  useEffect(() => {
    if (!open) return;
    const onDown = (e: MouseEvent) => { if (ref.current && !ref.current.contains(e.target as Node)) setOpen(false); };
    document.addEventListener("mousedown", onDown, true);
    return () => document.removeEventListener("mousedown", onDown, true);
  }, [open]);
  const cur = options.find((o) => o.value === value) ?? options[0];
  return (
    <div ref={ref} style={{ position: "relative" }} onClick={(e) => e.stopPropagation()}>
      <button type="button" className="lvpv-btn" style={selectBtn} onClick={() => setOpen((o) => !o)} aria-label={ariaLabel} aria-haspopup="listbox" aria-expanded={open}>
        <span style={{ overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{cur?.label}</span>
        <ChevronDown size={14} style={{ flexShrink: 0, color: T.fgQuaternary }} />
      </button>
      {open && (
        <div style={popover} role="listbox">
          {options.map((o) => (
            <div key={o.value} role="option" className="lvpv-item" aria-selected={o.value === value} style={popItem(o.value === value)} onClick={() => { onChange(o.value); setOpen(false); }}>
              {o.label}
            </div>
          ))}
        </div>
      )}
    </div>
  );
}

/* ─── Component ───────────────────────────────────────────────────────────── */

export interface PivotPanelProps {
  /** Every field available from the source (drawn as the draggable chip list). */
  fields: string[];
  /** The current spec (controlled). */
  spec: PivotSpec;
  /** Fired with the next spec on any change. */
  onChange: (spec: PivotSpec) => void;
  /** Close the drawer (host toggles visibility). */
  onClose?: () => void;
  /** Distinct values of a field (for the "Filter by values" checklist). */
  distinctValues?: (field: string) => string[];
}

// Google-Sheets section order: Rows, Columns, Values, Filters.
const SECTION_ORDER: Array<{ key: PivotArea; label: string }> = [
  { key: "rows", label: "Rows" },
  { key: "columns", label: "Columns" },
  { key: "values", label: "Values" },
  { key: "filters", label: "Filters" },
];

export function PivotPanel({ fields, spec, onChange, onClose, distinctValues }: PivotPanelProps) {
  const [dragField, setDragField] = useState<string | null>(null);
  const [overArea, setOverArea] = useState<PivotArea | null>(null);
  const [addOpen, setAddOpen] = useState<PivotArea | null>(null); // which section's "Add" menu is open
  const [addQuery, setAddQuery] = useState("");
  const [filterOpen, setFilterOpen] = useState<string | null>(null); // which filter field's checklist is open

  const startDrag = (field: string) => (e: DragEvent) => {
    setDragField(field);
    e.dataTransfer.effectAllowed = "move";
    try {
      e.dataTransfer.setData(DND_MIME, field);
      e.dataTransfer.setData("text/plain", field);
    } catch {
      /* some browsers restrict custom MIME in tests — the ref fallback covers it */
    }
  };
  const endDrag = () => { setDragField(null); setOverArea(null); };
  const fieldFrom = (e: DragEvent): string | null => e.dataTransfer.getData(DND_MIME) || e.dataTransfer.getData("text/plain") || dragField;
  const dropOnArea = (area: PivotArea, index?: number) => (e: DragEvent) => {
    e.preventDefault();
    e.stopPropagation();
    const field = fieldFrom(e);
    endDrag();
    if (field) onChange(placeField(spec, field, area, index));
  };
  const allowDrop = (area: PivotArea) => (e: DragEvent) => {
    e.preventDefault();
    e.dataTransfer.dropEffect = "move";
    if (overArea !== area) setOverArea(area);
  };
  const removeChip = (field: string) => onChange(removeField(spec, field));

  const addRef = useRef<HTMLDivElement | null>(null);
  useEffect(() => {
    if (!addOpen) return;
    const onDown = (e: MouseEvent) => { if (addRef.current && !addRef.current.contains(e.target as Node)) { setAddOpen(null); setAddQuery(""); } };
    document.addEventListener("mousedown", onDown, true);
    return () => document.removeEventListener("mousedown", onDown, true);
  }, [addOpen]);

  const filterRef = useRef<HTMLDivElement | null>(null);
  useEffect(() => {
    if (!filterOpen) return;
    const onDown = (e: MouseEvent) => { if (filterRef.current && !filterRef.current.contains(e.target as Node)) setFilterOpen(null); };
    document.addEventListener("mousedown", onDown, true);
    return () => document.removeEventListener("mousedown", onDown, true);
  }, [filterOpen]);

  const addableFields = (area: PivotArea): string[] => {
    const q = addQuery.trim().toLowerCase();
    const inArea = new Set(fieldsInArea(spec, area));
    return fields.filter((f) => !inArea.has(f) && (q ? f.toLowerCase().includes(q) : true));
  };

  const dimCard = (field: string): ReactNode => {
    const s = dimSettingOf(spec, field);
    return (
      <div style={ctrlRow}>
        <div>
          <label style={ctrlLabel}>Order</label>
          <Select value={s.order ?? "asc"} options={ORDER_OPTS} ariaLabel={`Order for ${field}`} onChange={(v) => onChange(setDimSetting(spec, field, { order: v }))} />
        </div>
        <label style={{ display: "flex", alignItems: "center", gap: 6, fontSize: 12.5, color: T.textSecondary, alignSelf: "flex-end", height: 30 }}>
          <input type="checkbox" style={{ accentColor: "var(--color-fiab-yellow-600)" }} checked={s.showTotals ?? true} onChange={(e) => onChange(setDimSetting(spec, field, { showTotals: e.target.checked }))} />
          Show totals
        </label>
      </div>
    );
  };
  const valueCard = (field: string): ReactNode => {
    const v = spec.values.find((x) => x.field === field);
    return (
      <div style={ctrlRow}>
        <div>
          <label style={ctrlLabel}>Summarize by</label>
          <Select value={aggregateOfValue(spec, field)} options={PIVOT_AGGREGATES} ariaLabel={`Summarize ${field}`} onChange={(a) => onChange(setValueAggregate(spec, field, a))} />
        </div>
        <div>
          <label style={ctrlLabel}>Show as</label>
          <Select value={v?.showAs ?? "default"} options={PIVOT_SHOW_AS} ariaLabel={`Show ${field} as`} onChange={(m) => onChange(setValueShowAs(spec, field, m))} />
        </div>
      </div>
    );
  };
  const filterCard = (field: string): ReactNode => {
    const include = filterIncludeOf(spec, field);
    const all = distinctValues?.(field) ?? [];
    const kept = include ?? all;
    const label = !include || include.length === all.length ? "Showing all items" : `Showing ${include.length} of ${all.length}`;
    return (
      <div ref={filterOpen === field ? filterRef : undefined} style={{ position: "relative" }}>
        <label style={ctrlLabel}>Status</label>
        <button type="button" className="lvpv-btn" style={selectBtn} onClick={() => setFilterOpen((f) => (f === field ? null : field))} aria-haspopup="dialog">
          <span style={{ overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{label}</span>
          <ChevronDown size={14} style={{ flexShrink: 0, color: T.fgQuaternary }} />
        </button>
        {filterOpen === field && (
          <div style={{ ...popover, width: 220, maxHeight: 300 }} role="dialog" aria-label={`Filter ${field}`}>
            <div style={{ display: "flex", gap: 10, padding: "4px 8px 6px", fontSize: 12, fontWeight: 600 }}>
              <button type="button" className="lvpv-link" style={{ color: T.brand, fontWeight: 600, background: "none", border: "none", cursor: "pointer", padding: 0 }} onClick={() => onChange(setFilterInclude(spec, field, null))}>Select all</button>
              <button type="button" className="lvpv-link" style={{ color: T.brand, fontWeight: 600, background: "none", border: "none", cursor: "pointer", padding: 0 }} onClick={() => onChange(setFilterInclude(spec, field, []))}>Clear</button>
            </div>
            {all.length === 0 && <div style={{ ...emptyHint, padding: "6px 10px" }}>No values.</div>}
            {all.map((val) => {
              const checked = kept.includes(val);
              return (
                <label key={val} className="lvpv-item" style={popItem(false)}>
                  <input
                    type="checkbox"
                    style={{ accentColor: "var(--color-fiab-yellow-600)" }}
                    checked={checked}
                    onChange={() => {
                      const next = checked ? kept.filter((x) => x !== val) : [...kept, val];
                      onChange(setFilterInclude(spec, field, next.length === all.length ? null : next));
                    }}
                  />
                  <span style={{ overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{val === "" ? "(blank)" : val}</span>
                </label>
              );
            })}
          </div>
        )}
      </div>
    );
  };

  return (
    <aside style={drawer} role="dialog" aria-label="Pivot table editor">
      {/* Scoped hover states so the buttons feel like the host's <Button> (which uses
          bg-primary_hover on hover). Inline styles can't express :hover. */}
      <style>{`
        .lvpv-btn:hover { background: var(--color-bg-primary_hover); }
        .lvpv-icon-btn:hover { background: var(--color-bg-primary_hover); color: var(--color-text-secondary); }
        .lvpv-item:hover { background: var(--color-bg-primary_hover); }
        .lvpv-link:hover { text-decoration: underline; }
      `}</style>
      <header style={header}>
        <span style={titleStyle}>Pivot table editor</span>
        {onClose && (
          <button type="button" aria-label="Close" className="lvpv-icon-btn" onClick={onClose} style={closeBtn}>
            <X size={18} />
          </button>
        )}
      </header>

      <div style={bodyStyle}>
        {SECTION_ORDER.map(({ key, label }) => {
          const placed = fieldsInArea(spec, key);
          const active = overArea === key;
          return (
            <div key={key} data-testid={`area-${key}`}>
              <div style={sectionHead}>
                <span style={areaTitle}>{label}</span>
                <div ref={addOpen === key ? addRef : undefined} style={{ position: "relative" }}>
                  <button type="button" className="lvpv-btn" style={addBtn} onClick={() => { setAddOpen((a) => (a === key ? null : key)); setAddQuery(""); }} data-testid={`add-${key}`} aria-label={`Add to ${label}`}>
                    <Plus size={13} /> Add
                  </button>
                  {addOpen === key && (
                    <div style={{ ...popover, right: 0, width: 200 }} role="menu">
                      <input style={{ ...searchInput, height: 30, margin: "2px 2px 6px" }} placeholder="Search fields" value={addQuery} autoFocus onChange={(e) => setAddQuery(e.target.value)} aria-label="Search fields to add" />
                      {addableFields(key).map((f) => (
                        <div key={f} role="menuitem" className="lvpv-item" style={popItem(false)} onClick={() => { onChange(placeField(spec, f, key)); setAddOpen(null); setAddQuery(""); }} data-testid={`add-field-${key}-${f}`}>
                          {f}
                        </div>
                      ))}
                      {addableFields(key).length === 0 && <div style={{ ...emptyHint, padding: "6px 10px" }}>No fields.</div>}
                    </div>
                  )}
                </div>
              </div>
              <div
                onDragOver={allowDrop(key)}
                onDragLeave={() => setOverArea((a) => (a === key ? null : a))}
                onDrop={dropOnArea(key)}
                style={{ display: "flex", flexDirection: "column", gap: 8, minHeight: 8, padding: active ? 6 : 0, borderRadius: 8, background: active ? "var(--color-utility-brand-50)" : "transparent", border: active ? "1px dashed var(--color-border-brand_alt)" : "1px solid transparent" }}
              >
                {placed.length === 0 && <span style={emptyHint}>No fields — use Add or drag a field here.</span>}
                {placed.map((field, i) => (
                  <div
                    key={field}
                    style={cardStyle}
                    onDragOver={(e) => { e.preventDefault(); e.stopPropagation(); }}
                    onDrop={dropOnArea(key, i)}
                    data-testid={`chip-${key}-${field}`}
                  >
                    <div style={cardHead}>
                      <span style={cardName} draggable onDragStart={startDrag(field)} onDragEnd={endDrag} title={field}>{field}</span>
                      <button type="button" aria-label={`Remove ${field}`} onClick={() => removeChip(field)} style={chipRemove}><X size={13} /></button>
                    </div>
                    {(key === "rows" || key === "columns") && dimCard(field)}
                    {key === "values" && valueCard(field)}
                    {key === "filters" && filterCard(field)}
                  </div>
                ))}
              </div>
            </div>
          );
        })}
      </div>
    </aside>
  );
}

export default PivotPanel;
