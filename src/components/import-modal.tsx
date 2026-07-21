/**
 * ImportModal — Google-Sheets-style "Import file" dialog shown after a user
 * picks a .xlsx/.csv. Lets them choose where the parsed data lands. Options that
 * don't apply to a standalone single-sheet component (create new spreadsheet /
 * insert new sheets) are shown disabled, matching Google's layout.
 */
import { useState, type CSSProperties } from "react";
import { Z_BASE } from "../core/z-index";

export type ImportLocation = "new-spreadsheet" | "new-sheets" | "replace-spreadsheet" | "replace-sheet" | "append" | "at-cell";

interface Option {
  value: ImportLocation;
  label: string;
  disabled?: boolean;
}

const OPTIONS: Option[] = [
  { value: "new-spreadsheet", label: "Create new spreadsheet" },
  { value: "new-sheets", label: "Insert new sheet(s)" },
  { value: "replace-spreadsheet", label: "Replace spreadsheet" },
  { value: "replace-sheet", label: "Replace current sheet" },
  { value: "append", label: "Append to current sheet" },
  { value: "at-cell", label: "Replace data at selected cell" },
];

export interface ImportModalProps {
  open: boolean;
  fileName: string;
  onCancel: () => void;
  onImport: (location: ImportLocation) => void;
}

const overlay: CSSProperties = { position: "fixed", inset: 0, background: "rgba(16,24,40,0.45)", display: "flex", alignItems: "center", justifyContent: "center", zIndex: Z_BASE + 4000 };
const card: CSSProperties = { width: 460, maxWidth: "92vw", background: "#fff", borderRadius: 14, boxShadow: "0 20px 48px rgba(16,24,40,0.28)", padding: "22px 24px", fontFamily: "'Work Sans', ui-sans-serif, system-ui, sans-serif" };

export function ImportModal({ open, fileName, onCancel, onImport }: ImportModalProps) {
  const [choice, setChoice] = useState<ImportLocation>("replace-sheet");
  if (!open) return null;
  return (
    <div style={overlay} onMouseDown={onCancel}>
      <div style={card} onMouseDown={(e) => e.stopPropagation()}>
        <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 16 }}>
          <span style={{ fontSize: 19, fontWeight: 600, color: "#101828" }}>Import file</span>
          <button type="button" aria-label="Close" onClick={onCancel} style={{ border: "none", background: "transparent", fontSize: 20, color: "#667085", cursor: "pointer", lineHeight: 1 }}>
            ✕
          </button>
        </div>

        <div style={{ fontSize: 12, color: "#667085", marginBottom: 4 }}>File</div>
        <div style={{ fontSize: 14, fontWeight: 600, color: "#101828", marginBottom: 18, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{fileName}</div>

        <div style={{ fontSize: 12, color: "#667085", marginBottom: 8 }}>Import location</div>
        <div style={{ display: "flex", flexDirection: "column", gap: 2, marginBottom: 22 }}>
          {OPTIONS.map((o) => (
            <label
              key={o.value}
              style={{ display: "flex", alignItems: "center", gap: 10, padding: "8px 8px", borderRadius: 8, cursor: o.disabled ? "default" : "pointer", color: o.disabled ? "#98a2b3" : "#344054", fontSize: 14 }}
              onMouseEnter={(e) => !o.disabled && (e.currentTarget.style.background = "#f9fafb")}
              onMouseLeave={(e) => (e.currentTarget.style.background = "transparent")}
            >
              <input
                type="radio"
                name="import-location"
                disabled={o.disabled}
                checked={!o.disabled && choice === o.value}
                onChange={() => !o.disabled && setChoice(o.value as ImportLocation)}
                style={{ accentColor: "#101828", width: 16, height: 16 }}
              />
              {o.label}
            </label>
          ))}
        </div>

        <div style={{ display: "flex", justifyContent: "flex-end", gap: 10 }}>
          <button
            type="button"
            onClick={onCancel}
            style={{ padding: "9px 18px", borderRadius: 8, border: "1px solid #d0d5dd", background: "#fff", color: "#344054", fontSize: 14, fontWeight: 600, cursor: "pointer", fontFamily: "inherit" }}
          >
            Cancel
          </button>
          <button
            type="button"
            onClick={() => onImport(choice)}
            onMouseEnter={(e) => (e.currentTarget.style.background = "#000")}
            onMouseLeave={(e) => (e.currentTarget.style.background = "#101828")}
            style={{ padding: "9px 22px", borderRadius: 8, border: "none", background: "#101828", color: "#fff", fontSize: 14, fontWeight: 600, cursor: "pointer", fontFamily: "inherit", transition: "background-color .15s ease" }}
          >
            Import data
          </button>
        </div>
      </div>
    </div>
  );
}
