/**
 * Cross-platform keyboard shortcuts (Mac ⌘ + Windows/Linux Ctrl).
 *
 * We bind the shortcuts tied to our own menus/toolbar, the browser-reserved
 * keys we must intercept (⌘F/S/O/P), AND the Bold/Italic/Underline toggles.
 * We route ⌘B/I/U through the Facade style setters (not Univer's native
 * shortcut) so the change goes through the same command path a toolbar click
 * uses — that's what re-lights the toolbar's B/I/U button. We stop propagation
 * so Univer's native handler doesn't also fire (no double-toggle), and we bail
 * while a cell/formula editor is focused so in-cell rich-text bolding still
 * works. Univer still natively owns Undo/Redo, Cut/Copy/Paste, Select-all,
 * arrow navigation, F2, hyperlink ⌘K, … . Matching uses `event.code` (physical key), so
 * it's keyboard-layout independent, and the accelerator is `metaKey || ctrlKey`
 * so the same table works on macOS and Windows.
 *
 * Shortcut choices follow the researched Excel/Google-Sheets mapping (Sheets-
 * first). NOTE: ⌘⇧R (align-right in Sheets) is intentionally NOT bound — it
 * collides with the browser's hard-reload; use the toolbar for align-right.
 */
import { insertAggregate, type FunctionsApi } from "./functions";
import { printSheet } from "../core/print-sheet";

/* ---- Loose Facade views --------------------------------------------------- */
interface KRange {
  setNumberFormat(p: string): unknown;
  setHorizontalAlignment(a: string): unknown;
  setValue(v: unknown): unknown;
  getCellStyleData?: () => { bl?: number; it?: number; ul?: { s?: number }; st?: { s?: number } } | null;
  setFontWeight(w: string): unknown;
  setFontStyle(s: string): unknown;
  setFontLine(l: string): unknown;
}
interface KSheet {
  getSheetId?: () => string;
  getSheetName?: () => string;
}
interface KWorkbook {
  getActiveRange?: () => KRange | null;
  getActiveSheet?: () => KSheet | null;
  getSheets?: () => KSheet[];
  setActiveSheet?: (s: KSheet | string) => unknown;
  create?: (name: string, rows: number, cols: number) => unknown;
}
interface KApi {
  getActiveWorkbook?: () => KWorkbook | null;
  executeCommand?: (id: string, params?: object) => unknown;
}

export interface ShortcutContext {
  api: unknown;
  /** Open the Find & Replace modal. */
  onFind?: () => void;
  /** Trigger File ▸ Import (file picker + location modal). */
  onImport?: () => void;
  /** Trigger printing (defaults to printing the active workbook). */
  onPrint?: () => void;
  /** Optional host save hook (⌘S). We always suppress the browser save dialog. */
  onSave?: () => void;
}

const pad2 = (n: number) => String(n).padStart(2, "0");

/**
 * Attach the global shortcut handler. `getCtx` is called on every keystroke so
 * the handler always sees the latest api/callbacks (avoids stale closures).
 * Returns a disposer.
 */
export function attachKeyboardShortcuts(getCtx: () => ShortcutContext): () => void {
  const onKey = (e: KeyboardEvent) => {
    const mod = e.metaKey || e.ctrlKey; // ⌘ on macOS, Ctrl on Windows/Linux
    const shift = e.shiftKey;
    const alt = e.altKey;
    const code = e.code;
    if (!mod && !alt && !(shift && code === "F11")) return; // fast bail on plain typing

    const ctx = getCtx();
    const api = ctx.api as KApi | null;
    const wb = () => api?.getActiveWorkbook?.() ?? null;
    const range = () => wb()?.getActiveRange?.() ?? null;
    const take = () => {
      e.preventDefault();
      e.stopImmediatePropagation(); // win over the browser + Univer's own handlers
    };
    // While a cell/formula editor is focused, ⌘B/I/U should format the rich
    // text being typed — that's Univer's job, so we don't intercept it.
    const editing = () => {
      const el = document.activeElement as HTMLElement | null;
      return !!el && (el.isContentEditable || el.tagName === "INPUT" || el.tagName === "TEXTAREA");
    };
    // Toggle a character style via the Facade (same path a toolbar click takes,
    // so the toolbar button re-lights). Reads the active cell's current style to
    // decide on/off.
    const toggleStyle = (k: "bl" | "it" | "ul" | "st") => {
      if (editing()) return; // let Univer's editor handle in-cell formatting
      take();
      try {
        const r = range();
        if (!r) return;
        const s = r.getCellStyleData?.() ?? {};
        if (k === "bl") r.setFontWeight(s.bl ? "normal" : "bold");
        else if (k === "it") r.setFontStyle(s.it ? "normal" : "italic");
        else if (k === "ul") r.setFontLine(s.ul?.s ? "none" : "underline");
        else r.setFontLine(s.st?.s ? "none" : "line-through");
      } catch {
        /* ignore */
      }
    };
    const numFmt = (p: string) => {
      take();
      try {
        range()?.setNumberFormat(p);
      } catch {
        /* ignore */
      }
    };
    const align = (a: string) => {
      take();
      try {
        range()?.setHorizontalAlignment(a);
      } catch {
        /* ignore */
      }
    };
    const moveSheet = (dir: number) => {
      take();
      try {
        const w = wb();
        const sheets = w?.getSheets?.() ?? [];
        const activeId = w?.getActiveSheet?.()?.getSheetId?.();
        const idx = sheets.findIndex((s) => s.getSheetId?.() === activeId);
        const next = sheets[idx + dir];
        if (next) w?.setActiveSheet?.(next);
      } catch {
        /* ignore */
      }
    };

    // ---- Browser-reserved / custom-UI (accelerator = ⌘/Ctrl) ---------------
    if (mod && !alt) {
      // Character-style toggles routed through the Facade so the toolbar's
      // B/I/U buttons reflect the new state (native ⌘B doesn't re-light them).
      if (code === "KeyB" && !shift) { toggleStyle("bl"); return; }             // Bold
      if (code === "KeyI" && !shift) { toggleStyle("it"); return; }             // Italic
      if (code === "KeyU" && !shift) { toggleStyle("ul"); return; }             // Underline
      if (code === "KeyF" && !shift) { take(); ctx.onFind?.(); return; }        // Find
      if (code === "KeyH" && shift) { take(); ctx.onFind?.(); return; }         // Find & Replace
      if (code === "KeyO" && !shift) { take(); ctx.onImport?.(); return; }      // Open / Import
      if (code === "KeyP" && !shift) { take(); (ctx.onPrint ?? (() => printSheet(wb() as never)))(); return; } // Print
      if (code === "KeyS" && !shift) { take(); ctx.onSave?.(); return; }        // Save (suppress browser dialog)
      if (code === "Backslash" && !shift) { take(); try { api?.executeCommand?.("sheet.command.clear-selection-format"); } catch { /* */ } return; } // Clear formatting

      // Number formats (Sheets: Ctrl+Shift+digit) — also accept ⌘ on Mac.
      if (shift) {
        switch (code) {
          case "Digit1": return numFmt("#,##0.00");            // Number
          case "Digit2": return numFmt("h:mm:ss am/pm");       // Time
          case "Digit3": return numFmt("yyyy-mm-dd");          // Date
          case "Digit4": return numFmt('"$"#,##0.00');         // Currency
          case "Digit5": return numFmt("0.00%");               // Percent
          case "KeyL": return align("left");
          case "KeyE": return align("center");
          // KeyR (align right) intentionally unbound — collides with hard-reload.
        }
      }

      // Insert current date / time.
      if (code === "Semicolon") {
        take();
        try {
          const d = new Date();
          const v = shift
            ? `${pad2(d.getHours())}:${pad2(d.getMinutes())}:${pad2(d.getSeconds())}`
            : `${d.getFullYear()}-${pad2(d.getMonth() + 1)}-${pad2(d.getDate())}`;
          range()?.setValue(v);
        } catch {
          /* ignore */
        }
        return;
      }
    }

    // ---- Alt-based -----------------------------------------------------------
    if (alt && !mod) {
      if (shift && code === "Digit5") { toggleStyle("st"); return; } // Strikethrough (Sheets: Alt+Shift+5) — toggles + re-lights the toolbar S̶ button
      if (code === "Equal") { take(); try { insertAggregate(api as unknown as FunctionsApi, "SUM"); } catch { /* */ } return; } // AutoSum
      if (code === "ArrowDown") { return moveSheet(1); }   // next sheet
      if (code === "ArrowUp") { return moveSheet(-1); }    // previous sheet
    }

    // ---- New sheet (Shift+F11) ----------------------------------------------
    if (shift && !mod && !alt && code === "F11") {
      take();
      try {
        const w = wb();
        const existing = new Set((w?.getSheets?.() ?? []).map((s) => s.getSheetName?.() ?? ""));
        // Lowest unused "SheetN" (Google-style), not existing.size+1 which can
        // skip a freed-up number.
        let i = 1;
        while (existing.has(`Sheet${i}`)) i++;
        w?.create?.(`Sheet${i}`, 100, 26);
      } catch {
        /* ignore */
      }
      return;
    }
  };

  window.addEventListener("keydown", onKey, true); // capture phase
  return () => window.removeEventListener("keydown", onKey, true);
}
