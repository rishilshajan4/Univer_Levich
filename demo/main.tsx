import { useMemo, useRef, useState } from "react";
import { createRoot } from "react-dom/client";
import { PocApp } from "./poc-app";
import { LevichSheet, LEVICH_BRAND, type ColumnDef, type LevichSheetHandle, type SheetData } from "../src";
import { takeImportPayload, takeSnapshotPayload } from "../src/core/import-data";
import { FORMULA_TESTS } from "./formula-tests";
import { FORMULA_TESTS_ALL } from "./formula-tests-all";

// "Create new spreadsheet" / "Replace spreadsheet" (import) opens this URL in a
// new tab (or reloads) and stashes the imported workbook. Rich .xlsx imports
// carry a full Univer snapshot (styles / merges / formats / all sheets); CSV or
// a failed parse falls back to a flat grid. Whichever is present becomes the
// initial document instead of a blank sheet.
const IMPORT_SNAPSHOT = takeSnapshotPayload<Record<string, unknown>>();
const IMPORT_PAYLOAD = IMPORT_SNAPSHOT ? null : takeImportPayload();

// PoC (localhost:9100/?poc): lazy multi-sheet loading — the whole workbook is
// converted on the "backend" (scripts/xlsx-poc.mjs), split into per-sheet JSON,
// and the FE loads only the active sheet, fetching others on tab-click.
// See demo/poc-app.tsx.
const POC = typeof location !== "undefined" && new URLSearchParams(location.search).has("poc");

/** Count the visible sheets in an imported snapshot (hidden ones exist but
 *  aren't shown as tabs — matches the source app). */
function visibleSheetCount(snap: Record<string, unknown>): number {
  const sheets = (snap.sheets as Record<string, { hidden?: number }> | undefined) ?? {};
  const list = Object.values(sheets);
  const visible = list.filter((s) => s.hidden !== 1).length;
  return visible || list.length || 1;
}
function gridToSheet(grid: (string | number)[][]): { columns: ColumnDef[]; data: SheetData } {
  const headers = grid[0] ?? [];
  const columns: ColumnDef[] = headers.map((h, i) => ({ key: `c${i}`, header: String(h ?? `Column ${i + 1}`), editable: true, width: 150 }));
  const data: SheetData = grid.slice(1).map((row) => {
    const o: Record<string, string | number> = {};
    columns.forEach((c, i) => (o[c.key] = row[i] ?? ""));
    return o;
  });
  return { columns, data };
}

// A fresh, EMPTY spreadsheet (no dummy data): blank, unheaded columns so the
// grid opens like a new Google Sheet. Import or typing fills it in.
const BLANK_COLUMNS: ColumnDef[] = Array.from({ length: 12 }, (_, i) => ({ key: `c${i}`, header: "", editable: true, width: 120 }));
const BLANK_DATA: SheetData = [];

document.title = "Untitled spreadsheet - FinOpz Sheets";

/** Official FinOpz brand logo (gold mark + wordmark). */
function FinOpzLogo() {
  return (
    <div style={{ display: "inline-flex", alignItems: "center", gap: 6 }} aria-label="FinOpz" role="img">
      <svg width="24" height="24" viewBox="0 0 21 21" fill="none" xmlns="http://www.w3.org/2000/svg" style={{ flexShrink: 0 }}>
        <path d="M20.1172 19.7521C20.1172 19.9541 19.9534 20.1179 19.7514 20.1179H13.7772C13.5752 20.1179 13.4114 19.9541 13.4114 19.7521V13.4121H20.1172V19.7521Z" fill="#EFC71D" />
        <path d="M20.1172 6.70576H13.4114V4.76837e-07H19.7514C19.9534 4.76837e-07 20.1172 0.163761 20.1172 0.365769V6.70576Z" fill="#EFC71D" />
        <rect x="13.4121" y="6.70576" width="6.70576" height="6.70576" transform="rotate(180 13.4121 6.70576)" fill="#EFC71D" />
        <path d="M6.70508 6.70576H0.365085C0.163076 6.70576 -0.000684261 6.542 -0.000684261 6.33999V0.365769C-0.000684261 0.163761 0.163076 4.76837e-07 0.365085 4.76837e-07H6.70508V6.70576Z" fill="#EFC71D" />
        <rect x="20.1172" y="13.4108" width="6.70576" height="6.70576" transform="rotate(180 20.1172 13.4108)" fill="#EFC71D" />
        <path d="M13.4121 13.4108H7.07212C6.87011 13.4108 6.70635 13.2471 6.70635 13.0451V6.70508H13.4121V13.4108Z" fill="#EFC71D" />
        <rect x="6.70508" y="20.1179" width="6.70576" height="6.70576" rx="0.365769" transform="rotate(180 6.70508 20.1179)" fill="#EFC71D" />
      </svg>
      <span
        style={{
          fontFamily: "'Work Sans', ui-sans-serif, system-ui, sans-serif",
          fontSize: 22,
          lineHeight: 1,
          fontWeight: 400,
          letterSpacing: "-0.66px",
          color: "#313131",
        }}
      >
        FinOpz
      </span>
    </div>
  );
}

function App() {
  const ref = useRef<LevichSheetHandle>(null);
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const apiRef = useRef<any>(null);
  const [note, setNote] = useState("");
  const [closed, setClosed] = useState(false);
  // Spreadsheet document title, updated by File ▸ Rename (onRename hook).
  const [renamedTitle, setRenamedTitle] = useState<string | null>(null);

  // Seed the active sheet with the whole formula matrix so Univer computes each
  // one live — Result (D) is the real formula, Expected (E) is the known answer,
  // Status (F) is a tolerant PASS/CHECK comparison. Verifies the free formula
  // engine end-to-end (SUM/MIN/MAX/VLOOKUP/HLOOKUP/… and more).
  const verifyFunctions = () => {
    const api = apiRef.current;
    const ws = api?.getActiveWorkbook?.()?.getActiveSheet?.();
    if (!ws) return setNote("Sheet not ready yet");
    const header = ["Category", "Function", "Formula", "Result (live)", "Expected", "Status"];
    // Works for text, numbers, and booleans (direct D=E), with a numeric
    // tolerance for floats and a format-tolerant text fallback (coerce→trim→
    // strip thousands-commas) so a correct value that differs only in type or
    // formatting — e.g. FIXED/DOLLAR/DEC2BIN returning text vs a numeric
    // expected — still counts as PASS. IFERROR turns #VALUE!/#N/A/#NAME? into
    // "✗ ERROR" instead of poisoning the Status cell.
    // Nested IFs so ABS(D-E) is ONLY evaluated when both are numbers — a flat
    // OR/AND would evaluate ABS(text-text) too (spreadsheets don't short-
    // circuit), erroring on text results even when D exactly equals E.
    const status = (r: number) =>
      `=IFERROR(IF(D${r}=E${r},"✓ PASS",IF(AND(ISNUMBER(D${r}),ISNUMBER(E${r})),IF(ABS(D${r}-E${r})<0.01,"✓ PASS","✗ CHECK"),IF(SUBSTITUTE(TRIM(D${r}&""),",","")=SUBSTITUTE(TRIM(E${r}&""),",",""),"✓ PASS","✗ CHECK"))),"✗ ERROR")`;
    const rows: unknown[][] = [header];
    FORMULA_TESTS.forEach((t) => {
      rows.push([t.category, t.label, t.formula.slice(1), t.formula, t.expected, status(rows.length + 1)]);
    });

    // Criteria/rank/lookup functions need REAL cell ranges (Univer rejects
    // {array} constants there). Prove they work against an input block in H1:I5.
    const rangeTests: Array<[string, string, string, number]> = [
      ["SUM(range)", "SUM(H1:H5)", "=SUM(H1:H5)", 150],
      ["MIN(range)", "MIN(H1:H5)", "=MIN(H1:H5)", 10],
      ["MAX(range)", "MAX(H1:H5)", "=MAX(H1:H5)", 50],
      ["AVERAGE(range)", "AVERAGE(H1:H5)", "=AVERAGE(H1:H5)", 30],
      ["SUMIF range >25", 'SUMIF(H1:H5,">25")', '=SUMIF(H1:H5,">25")', 120],
      ["COUNTIF range >25", 'COUNTIF(H1:H5,">25")', '=COUNTIF(H1:H5,">25")', 3],
      ["AVERAGEIF range >25", 'AVERAGEIF(H1:H5,">25")', '=AVERAGEIF(H1:H5,">25")', 40],
      ["SUMIFS range", 'SUMIFS(H1:H5,H1:H5,">=30")', '=SUMIFS(H1:H5,H1:H5,">=30")', 120],
      ["COUNTIFS range", 'COUNTIFS(H1:H5,">=30",I1:I5,"<=400")', '=COUNTIFS(H1:H5,">=30",I1:I5,"<=400")', 2],
      ["AVERAGEIFS range", 'AVERAGEIFS(I1:I5,H1:H5,">=30")', '=AVERAGEIFS(I1:I5,H1:H5,">=30")', 400],
      ["RANK in range", "RANK(30,H1:H5,0)", "=RANK(30,H1:H5,0)", 3],
      ["VLOOKUP range", "VLOOKUP(30,H1:I5,2,FALSE)", "=VLOOKUP(30,H1:I5,2,FALSE)", 300],
    ];
    rangeTests.forEach(([label, text, formula, expected]) => {
      rows.push(["Cell-range", label, text, formula, expected, status(rows.length + 1)]);
    });

    try {
      ws.setColumnCount?.(Math.max(9, ws.getMaxColumns?.() ?? 9));
      ws.setRowCount?.(Math.max(rows.length + 2, ws.getMaxRows?.() ?? rows.length + 2));
      // Input block FIRST (H1:H5 = 10..50, I1:I5 = 100..500) so the range
      // formulas resolve against real data.
      ws.getRange(0, 7, 5, 2)?.setValues([[10, 100], [20, 200], [30, 300], [40, 400], [50, 500]]);
      ws.getRange(0, 0, rows.length, header.length)?.setValues(rows);
      [100, 160, 320, 130, 130, 100].forEach((w, c) => ws.setColumnWidth?.(c, w));
      ws.getRange(0, 0, 1, header.length)?.setFontWeight?.("bold");
      setNote(`Ran ${rows.length - 1} formulas — check the Status column`);
    } catch (e) {
      console.warn("[demo] verify functions failed", e);
      setNote("Verify failed — see console");
    }
  };

  // MASSIVE-SCALE check: run the FULL per-function matrix (one valid, deterministic
  // formula per catalog function) with live Result + Expected + PASS/CHECK — the
  // same treatment as "Verify functions", but for all ~529. "N/A" expecteds
  // (volatile / external / host-dependent) render as "⊘ N/A" (registered, no
  // deterministic check).
  const verifyAllFunctions = () => {
    const api = apiRef.current;
    const ws = api?.getActiveWorkbook?.()?.getActiveSheet?.();
    if (!ws) return setNote("Sheet not ready yet");
    if (!FORMULA_TESTS_ALL.length) return setNote("Full 529 matrix not loaded yet");

    const status = (r: number) =>
      `=IFERROR(IF(D${r}=E${r},"✓ PASS",IF(AND(ISNUMBER(D${r}),ISNUMBER(E${r})),IF(ABS(D${r}-E${r})<0.01,"✓ PASS","✗ CHECK"),IF(SUBSTITUTE(TRIM(D${r}&""),",","")=SUBSTITUTE(TRIM(E${r}&""),",",""),"✓ PASS","✗ CHECK"))),"✗ ERROR")`;
    const header = ["Category", "Function", "Formula", "Result (live)", "Expected", "Status"];
    const rows: unknown[][] = [header];
    FORMULA_TESTS_ALL.forEach((t) => {
      const r = rows.length + 1;
      const isNA = t.expected === "N/A";
      rows.push([t.category, t.label, t.formula.slice(1), t.formula, t.expected, isNA ? "⊘ N/A" : status(r)]);
    });

    try {
      ws.setColumnCount?.(Math.max(9, ws.getMaxColumns?.() ?? 9));
      ws.setRowCount?.(Math.max(rows.length + 2, ws.getMaxRows?.() ?? rows.length + 2));
      ws.getRange(0, 7, 5, 2)?.setValues([[10, 100], [20, 200], [30, 300], [40, 400], [50, 500]]);
      ws.getRange(0, 0, rows.length, header.length)?.setValues(rows);
      [110, 180, 320, 150, 140, 110].forEach((w, c) => ws.setColumnWidth?.(c, w));
      ws.getRange(0, 0, 1, header.length)?.setFontWeight?.("bold");
      setNote(`Computing ${FORMULA_TESTS_ALL.length} functions…`);
    } catch (e) {
      console.warn("[demo] verify all — seed failed", e);
      return setNote("Verify-all failed — see console");
    }

    // After the engine finishes, read the Status column (F) and summarize.
    window.setTimeout(() => {
      try {
        const st = ws.getRange(1, 5, FORMULA_TESTS_ALL.length, 1)?.getValues?.() ?? [];
        let pass = 0, check = 0, err = 0, na = 0;
        const nonPass: string[] = [];
        st.forEach((row, i) => {
          const cell = row?.[0];
          const v = String((cell && typeof cell === "object" ? (cell as { v?: unknown }).v : cell) ?? "");
          if (v.includes("PASS")) pass++;
          else if (v.includes("N/A")) na++;
          else if (v.includes("ERROR")) { err++; nonPass.push(FORMULA_TESTS_ALL[i].label); }
          else if (v.includes("CHECK")) { check++; nonPass.push(FORMULA_TESTS_ALL[i].label); }
        });
        console.info(`[levich] ALL → ${pass} PASS · ${check} CHECK · ${err} ERROR · ${na} N/A. Non-pass:`, nonPass);
        setNote(`${pass} PASS · ${check} CHECK · ${err} ERROR · ${na} N/A of ${FORMULA_TESTS_ALL.length}`);
      } catch (e) {
        console.warn("[demo] verify all — summary read failed", e);
      }
    }, 5000);
  };

  const imported = useMemo(() => (IMPORT_PAYLOAD ? gridToSheet(IMPORT_PAYLOAD) : null), []);
  const data = imported ? imported.data : BLANK_DATA;
  const columns = imported ? imported.columns : BLANK_COLUMNS;

  const download = async () => {
    const n = (await ref.current?.exportXlsx("levich-demo.xlsx")) ?? 0;
    setNote(n ? `Exported ${n} rows to .xlsx` : "Sheet not ready yet");
  };

  if (closed) {
    return (
      <div style={{ display: "grid", placeItems: "center", height: "100vh", color: "#667085" }}>
        <button
          onClick={() => setClosed(false)}
          style={{ padding: "10px 18px", borderRadius: 8, border: `1px solid ${LEVICH_BRAND}`, background: "#fff", color: LEVICH_BRAND, fontWeight: 600, cursor: "pointer" }}
        >
          Reopen Levich Sheet
        </button>
      </div>
    );
  }

  const title = renamedTitle ?? (IMPORT_SNAPSHOT ? "Imported workbook" : imported ? "Imported spreadsheet" : "Untitled spreadsheet");
  const subtitle = IMPORT_SNAPSHOT ? `${visibleSheetCount(IMPORT_SNAPSHOT)} sheet(s) · rich import` : imported ? `${data.length} rows imported` : "Blank spreadsheet";

  return (
    <div style={{ display: "flex", flexDirection: "column", height: "100vh" }}>
      <header
        style={{
          display: "flex",
          alignItems: "center",
          gap: 16,
          padding: "10px 20px",
          borderBottom: "1px solid #e4e7ec",
          background: "#fff",
        }}
      >
        <FinOpzLogo />

        <div style={{ width: 1, height: 28, background: "#e4e7ec" }} />

        <div style={{ display: "flex", flexDirection: "column", lineHeight: 1.25 }}>
          <span style={{ fontSize: 15, fontWeight: 700, color: "#101828" }}>{title}</span>
          <span style={{ fontSize: 12, color: "#667085" }}>{subtitle}</span>
        </div>

        <div style={{ marginLeft: "auto", display: "flex", alignItems: "center", gap: 12 }}>
          {note && <span style={{ color: "#067647", fontSize: 13, fontWeight: 500 }}>{note}</span>}

          <button
            onClick={verifyFunctions}
            style={{
              display: "inline-flex",
              alignItems: "center",
              padding: "8px 16px",
              borderRadius: 8,
              border: `1px solid ${LEVICH_BRAND}`,
              background: "#fff",
              color: LEVICH_BRAND,
              fontSize: 14,
              fontWeight: 600,
              cursor: "pointer",
            }}
          >
            Verify functions
          </button>

          <button
            onClick={verifyAllFunctions}
            style={{
              display: "inline-flex",
              alignItems: "center",
              padding: "8px 16px",
              borderRadius: 8,
              border: `1px solid ${LEVICH_BRAND}`,
              background: "#fff",
              color: LEVICH_BRAND,
              fontSize: 14,
              fontWeight: 600,
              cursor: "pointer",
            }}
          >
            Verify ALL 529
          </button>

          <button
            onClick={download}
            onMouseEnter={(e) => (e.currentTarget.style.background = "#000")}
            onMouseLeave={(e) => (e.currentTarget.style.background = "#101828")}
            style={{
              display: "inline-flex",
              alignItems: "center",
              gap: 8,
              padding: "8px 16px",
              borderRadius: 8,
              border: "none",
              background: "#101828",
              color: "#fff",
              fontSize: 14,
              fontWeight: 600,
              boxShadow: "0 1px 2px rgba(16,24,40,0.05)",
              cursor: "pointer",
              transition: "background-color 0.15s ease",
            }}
          >
            Download .xlsx
          </button>

          <button
            onClick={() => setClosed(true)}
            aria-label="Close"
            style={{
              width: 32,
              height: 32,
              borderRadius: 8,
              border: "none",
              background: "transparent",
              color: "#667085",
              fontSize: 18,
              cursor: "pointer",
            }}
          >
            ✕
          </button>
        </div>
      </header>

      <div style={{ flex: 1, minHeight: 0 }}>
        <LevichSheet
          ref={ref}
          data={data}
          columns={columns}
          snapshot={IMPORT_SNAPSHOT ?? undefined}
          freeze={imported ? { rows: 1 } : false}
          currencySymbol="$"
          onReady={(api) => { apiRef.current = api; }}
          // File ▸ Rename renames the SPREADSHEET (document) — reflect it in the header.
          onRename={(name) => { setRenamedTitle(name); document.title = `${name} - FinOpz Sheets`; }}
          // Real-world routing: document-level imports (create / replace a whole
          // document) are owned by the FinOpz backend. In PRODUCTION you'd POST/
          // PUT to your API and `return true` so the sheet skips its built-in
          // behavior. On THIS demo there's no backend, so we only log where that
          // call would go and `return false` — letting the sheet's built-in
          // behavior run (new tab / clear+write) so it still works on localhost.
          onImport={(grid, location) => {
            if (location === "new-spreadsheet" || location === "replace-spreadsheet") {
              // PROD: await fetch("/api/spreadsheets", { method: location === "new-spreadsheet" ? "POST" : "PUT", body: JSON.stringify(grid) }); return true;
              console.log(`[FinOpz BE] would ${location} — ${grid.length} rows (no backend on localhost → using built-in behavior)`);
            }
            return false; // demo: always use the sheet's built-in local behavior
          }}
        />
      </div>
    </div>
  );
}

createRoot(document.getElementById("root")!).render(POC ? <PocApp /> : <App />);
