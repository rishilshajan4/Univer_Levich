/**
 * RenameModal — Google-Sheets-style rename dialog. Replaces the native
 * `window.prompt`, matching the Levich modal look (see ImportModal). Validates
 * inline: trims, blocks empty, and blocks names already taken by another sheet
 * (the `taken` set). Enter submits, Esc / overlay-click / Cancel closes.
 */
import { useEffect, useRef, useState, type CSSProperties } from "react";

export interface RenameModalProps {
  open: boolean;
  /** Dialog heading, e.g. "Rename sheet". */
  title?: string;
  /** Current name, pre-filled and text-selected on open. */
  current: string;
  /** Lower-cased names already in use (excluding `current`) — blocks duplicates. */
  taken?: Set<string>;
  onCancel: () => void;
  onRename: (name: string) => void;
}

const overlay: CSSProperties = { position: "fixed", inset: 0, background: "rgba(16,24,40,0.45)", display: "flex", alignItems: "center", justifyContent: "center", zIndex: 4000 };
const card: CSSProperties = { width: 420, maxWidth: "92vw", background: "#fff", borderRadius: 14, boxShadow: "0 20px 48px rgba(16,24,40,0.28)", padding: "22px 24px", fontFamily: "'Work Sans', ui-sans-serif, system-ui, sans-serif" };

export function RenameModal({ open, title = "Rename", current, taken, onCancel, onRename }: RenameModalProps) {
  const [value, setValue] = useState(current);
  const inputRef = useRef<HTMLInputElement>(null);

  // Reset + focus/select whenever the dialog (re)opens.
  useEffect(() => {
    if (!open) return;
    setValue(current);
    const t = setTimeout(() => inputRef.current?.select(), 0);
    return () => clearTimeout(t);
  }, [open, current]);

  if (!open) return null;

  const trimmed = value.trim();
  const duplicate = !!trimmed && trimmed.toLowerCase() !== current.toLowerCase() && !!taken?.has(trimmed.toLowerCase());
  const error = !trimmed ? "Name can't be empty" : duplicate ? "That name is already taken" : "";
  const canSubmit = !error && trimmed !== current;

  const submit = () => {
    if (!canSubmit) return;
    onRename(trimmed);
  };

  return (
    <div style={overlay} onMouseDown={onCancel}>
      <div style={card} onMouseDown={(e) => e.stopPropagation()}>
        <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 16 }}>
          <span style={{ fontSize: 19, fontWeight: 600, color: "#101828" }}>{title}</span>
          <button type="button" aria-label="Close" onClick={onCancel} style={{ border: "none", background: "transparent", fontSize: 20, color: "#667085", cursor: "pointer", lineHeight: 1 }}>
            ✕
          </button>
        </div>

        <input
          ref={inputRef}
          value={value}
          onChange={(e) => setValue(e.target.value)}
          onKeyDown={(e) => {
            if (e.key === "Enter") { e.preventDefault(); submit(); }
            if (e.key === "Escape") { e.preventDefault(); onCancel(); }
          }}
          placeholder="Enter a name"
          style={{
            width: "100%", boxSizing: "border-box", padding: "10px 12px", fontSize: 14, color: "#101828",
            border: `1px solid ${error ? "#f04438" : "#d0d5dd"}`, borderRadius: 8, outline: "none",
            fontFamily: "inherit",
          }}
        />
        <div style={{ minHeight: 18, marginTop: 6, fontSize: 12, color: "#f04438" }}>{value.trim() !== current || value.trim() === "" ? error : ""}</div>

        <div style={{ display: "flex", justifyContent: "flex-end", gap: 10, marginTop: 12 }}>
          <button
            type="button"
            onClick={onCancel}
            style={{ padding: "9px 18px", borderRadius: 8, border: "1px solid #d0d5dd", background: "#fff", color: "#344054", fontSize: 14, fontWeight: 600, cursor: "pointer", fontFamily: "inherit" }}
          >
            Cancel
          </button>
          <button
            type="button"
            onClick={submit}
            disabled={!canSubmit}
            onMouseEnter={(e) => canSubmit && (e.currentTarget.style.background = "#000")}
            onMouseLeave={(e) => canSubmit && (e.currentTarget.style.background = "#101828")}
            style={{
              padding: "9px 22px", borderRadius: 8, border: "none", fontSize: 14, fontWeight: 600, fontFamily: "inherit",
              background: canSubmit ? "#101828" : "#d0d5dd", color: "#fff", cursor: canSubmit ? "pointer" : "not-allowed",
              transition: "background-color .15s ease",
            }}
          >
            OK
          </button>
        </div>
      </div>
    </div>
  );
}

export default RenameModal;
