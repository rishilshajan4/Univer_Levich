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
 * Styling mirrors the package's other panels (white surface, #eaecf0 borders,
 * #101828/#475467 text) — see `find-replace-modal.tsx` / `levich-toolbar.tsx`.
 */
import { useMemo, useState, type CSSProperties, type DragEvent } from "react";
import { X } from "@untitledui/icons";
import type { PivotAggregate, PivotSpec, PivotValueField } from "../core/types";

/* ─── Spec model (pure, testable) ─────────────────────────────────────────── */

/** The four drop areas a field can live in. */
export type PivotArea = "filters" | "columns" | "rows" | "values";

export const PIVOT_AREAS: Array<{ key: PivotArea; label: string }> = [
  { key: "filters", label: "Filters" },
  { key: "columns", label: "Columns" },
  { key: "rows", label: "Rows" },
  { key: "values", label: "Values" },
];

export const PIVOT_AGGREGATES: Array<{ value: PivotAggregate; label: string }> = [
  { value: "sum", label: "Sum" },
  { value: "count", label: "Count" },
  { value: "average", label: "Average" },
  { value: "min", label: "Min" },
  { value: "max", label: "Max" },
  { value: "countNumbers", label: "Count Numbers" },
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

/* ─── Styles ──────────────────────────────────────────────────────────────── */

const drawer: CSSProperties = {
  position: "absolute",
  top: 0,
  right: 0,
  bottom: 0,
  width: 320,
  background: "#fff",
  borderLeft: "1px solid #eaecf0",
  boxShadow: "-8px 0 24px rgba(16,24,40,0.08)",
  display: "flex",
  flexDirection: "column",
  zIndex: 50,
};
const header: CSSProperties = { display: "flex", alignItems: "center", justifyContent: "space-between", padding: "14px 16px 8px", flexShrink: 0 };
const titleStyle: CSSProperties = { fontSize: 15, fontWeight: 600, color: "#101828" };
const closeBtn: CSSProperties = { display: "inline-flex", alignItems: "center", justifyContent: "center", width: 28, height: 28, borderRadius: 8, border: "none", background: "transparent", color: "#475467", cursor: "pointer" };
const bodyStyle: CSSProperties = { display: "flex", flexDirection: "column", flex: 1, minHeight: 0, overflowY: "auto", padding: "0 16px 16px" };
const sectionLabel: CSSProperties = { fontSize: 11, fontWeight: 700, letterSpacing: 0.3, color: "#98a2b3", margin: "14px 2px 6px" };
const searchInput: CSSProperties = { width: "100%", height: 34, borderRadius: 8, border: "1px solid #d0d5dd", padding: "0 12px", fontSize: 13, boxSizing: "border-box", outline: "none", color: "#101828" };
const fieldListStyle: CSSProperties = { display: "flex", flexWrap: "wrap", gap: 6, marginTop: 8, maxHeight: 132, overflowY: "auto", paddingBottom: 2 };
const areaGrid: CSSProperties = { display: "grid", gridTemplateColumns: "1fr 1fr", gap: 8, marginTop: 8 };
const chipBase: CSSProperties = {
  display: "inline-flex",
  alignItems: "center",
  gap: 6,
  padding: "4px 8px",
  borderRadius: 6,
  border: "1px solid #eaecf0",
  background: "#f9fafb",
  color: "#344054",
  fontSize: 12.5,
  cursor: "grab",
  userSelect: "none",
  maxWidth: "100%",
};
const chipRemove: CSSProperties = { display: "inline-flex", alignItems: "center", justifyContent: "center", width: 15, height: 15, borderRadius: 4, border: "none", background: "transparent", color: "#98a2b3", cursor: "pointer", flexShrink: 0, padding: 0 };
const aggSelect: CSSProperties = { height: 22, border: "1px solid #d0d5dd", borderRadius: 5, background: "#fff", color: "#475467", fontSize: 11.5, padding: "0 2px", cursor: "pointer", maxWidth: 96 };

const areaBoxBase: CSSProperties = {
  minHeight: 76,
  border: "1px dashed #d0d5dd",
  borderRadius: 8,
  padding: 8,
  display: "flex",
  flexDirection: "column",
  gap: 6,
  background: "#fff",
  transition: "background-color .12s ease, border-color .12s ease",
};
const areaTitle: CSSProperties = { fontSize: 11, fontWeight: 600, color: "#667085", textTransform: "uppercase", letterSpacing: 0.3 };
const emptyHint: CSSProperties = { fontSize: 11.5, color: "#c0c6d0", fontStyle: "italic" };

const DND_MIME = "application/x-levich-pivot-field";

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
}

export function PivotPanel({ fields, spec, onChange, onClose }: PivotPanelProps) {
  const [query, setQuery] = useState("");
  const [dragField, setDragField] = useState<string | null>(null);
  const [overArea, setOverArea] = useState<PivotArea | null>(null);

  const unusedFields = useMemo(() => {
    const q = query.trim().toLowerCase();
    return fields.filter((f) => (q ? f.toLowerCase().includes(q) : true));
  }, [fields, query]);

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
  const endDrag = () => {
    setDragField(null);
    setOverArea(null);
  };

  const fieldFrom = (e: DragEvent): string | null => {
    const fromData = e.dataTransfer.getData(DND_MIME) || e.dataTransfer.getData("text/plain");
    return fromData || dragField;
  };

  const dropOnArea = (area: PivotArea, index?: number) => (e: DragEvent) => {
    e.preventDefault();
    e.stopPropagation();
    const field = fieldFrom(e);
    endDrag();
    if (!field) return;
    onChange(placeField(spec, field, area, index));
  };

  const allowDrop = (area: PivotArea) => (e: DragEvent) => {
    e.preventDefault();
    e.dataTransfer.dropEffect = "move";
    if (overArea !== area) setOverArea(area);
  };

  const removeChip = (field: string) => onChange(removeField(spec, field));

  return (
    <aside style={drawer} role="dialog" aria-label="Pivot table editor">
      <header style={header}>
        <span style={titleStyle}>Pivot table editor</span>
        {onClose && (
          <button type="button" aria-label="Close" onClick={onClose} style={closeBtn}>
            <X size={18} />
          </button>
        )}
      </header>

      <div style={bodyStyle}>
        <div style={sectionLabel}>FIELDS</div>
        <input style={searchInput} placeholder="Search fields" value={query} onChange={(e) => setQuery(e.target.value)} aria-label="Search fields" />
        <div style={fieldListStyle}>
          {unusedFields.map((field) => {
            const area = areaOfField(spec, field);
            return (
              <span
                key={field}
                draggable
                onDragStart={startDrag(field)}
                onDragEnd={endDrag}
                title={area ? `${field} — in ${area}` : field}
                style={{ ...chipBase, opacity: area ? 0.55 : 1, borderColor: area ? "#c9cfd8" : "#eaecf0" }}
                data-testid={`field-chip-${field}`}
              >
                {field}
              </span>
            );
          })}
          {unusedFields.length === 0 && <span style={emptyHint}>No fields match your search.</span>}
        </div>

        <div style={sectionLabel}>AREAS</div>
        <div style={areaGrid}>
          {PIVOT_AREAS.map(({ key, label }) => {
            const placed = fieldsInArea(spec, key);
            const active = overArea === key;
            return (
              <div
                key={key}
                onDragOver={allowDrop(key)}
                onDragLeave={() => setOverArea((a) => (a === key ? null : a))}
                onDrop={dropOnArea(key)}
                style={{ ...areaBoxBase, background: active ? "#fef3c7" : "#fff", borderColor: active ? "#fde68a" : "#d0d5dd" }}
                data-testid={`area-${key}`}
                aria-label={label}
              >
                <span style={areaTitle}>{label}</span>
                {placed.map((field, i) => (
                  <span
                    key={field}
                    draggable
                    onDragStart={startDrag(field)}
                    onDragEnd={endDrag}
                    // Drop BEFORE this chip → reorder within the area.
                    onDragOver={(e) => { e.preventDefault(); e.stopPropagation(); }}
                    onDrop={dropOnArea(key, i)}
                    style={{ ...chipBase, background: "#fff", cursor: "grab", justifyContent: "space-between", width: "100%", boxSizing: "border-box" }}
                    data-testid={`chip-${key}-${field}`}
                  >
                    <span style={{ overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{field}</span>
                    <span style={{ display: "inline-flex", alignItems: "center", gap: 4, flexShrink: 0 }}>
                      {key === "values" && (
                        <select
                          value={aggregateOfValue(spec, field)}
                          onChange={(e) => onChange(setValueAggregate(spec, field, e.target.value as PivotAggregate))}
                          onClick={(e) => e.stopPropagation()}
                          style={aggSelect}
                          aria-label={`Aggregate for ${field}`}
                          data-testid={`agg-${field}`}
                        >
                          {PIVOT_AGGREGATES.map((a) => (
                            <option key={a.value} value={a.value}>{a.label}</option>
                          ))}
                        </select>
                      )}
                      <button type="button" aria-label={`Remove ${field}`} onClick={() => removeChip(field)} style={chipRemove}>
                        <X size={12} />
                      </button>
                    </span>
                  </span>
                ))}
                {placed.length === 0 && <span style={emptyHint}>Drop fields here</span>}
              </div>
            );
          })}
        </div>
      </div>
    </aside>
  );
}

export default PivotPanel;
