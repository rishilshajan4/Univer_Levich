import { describe, expect, it } from "vitest";
import { emptyFormulaCells, findHashCell } from "../../src/core/snapshot-scan";

// Minimal raw Univer snapshot: one sheet, cellData keyed [row][col] = { v, f }.
function snap(cells: Record<number, Record<number, { v?: unknown; f?: string }>>) {
  return { sheetOrder: ["s1"], sheets: { s1: { cellData: cells } } };
}

describe("activeCellData targets the HYDRATED sheet (shell-workbook model)", () => {
  it("scans the first sheet WITH cell data, not sheetOrder[0] when 0 is an empty shell", () => {
    // s1 = empty shell (active sheet is s2, which carries the real data + the # marker).
    const shell = {
      sheetOrder: ["s1", "s2"],
      sheets: {
        s1: { cellData: {} },
        s2: { cellData: { 0: { 0: { v: "Header" } }, 3: { 1: { v: "#review" } } } },
      },
    };
    expect(findHashCell(shell)).toEqual({ row: 3, column: 1 });
  });
});

describe("findHashCell (feature #2 — open at the # cell)", () => {
  it("finds the first cell whose text starts with '#'", () => {
    const s = snap({
      0: { 0: { v: "Header" } },
      2: { 1: { v: "#review here" } },
    });
    expect(findHashCell(s)).toEqual({ row: 2, column: 1 });
  });

  it("ignores Excel error tokens like #REF! / #NAME?", () => {
    const s = snap({ 0: { 0: { v: "#REF!" } }, 1: { 0: { v: "#NAME?" } }, 3: { 2: { v: "#anchor" } } });
    expect(findHashCell(s)).toEqual({ row: 3, column: 2 });
  });

  it("ignores coded IDs like '#123' (only '#', '# note', '#letter' are markers)", () => {
    expect(findHashCell(snap({ 0: { 0: { v: "#123" } }, 1: { 0: { v: "#4A" } } }))).toBeNull();
    expect(findHashCell(snap({ 0: { 0: { v: "#" } } }))).toEqual({ row: 0, column: 0 });
    expect(findHashCell(snap({ 0: { 0: { v: "# note" } } }))).toEqual({ row: 0, column: 0 });
  });

  it("returns null when there is no # marker", () => {
    expect(findHashCell(snap({ 0: { 0: { v: "plain" } } }))).toBeNull();
    expect(findHashCell({})).toBeNull();
  });
});

describe("emptyFormulaCells (feature #12 — recompute only uncached formula cells)", () => {
  it("returns formula cells that have NO cached value", () => {
    const s = snap({
      1: { 0: { f: "=SUM(A1:A3)" } }, // no v → needs compute
      2: { 0: { f: "=SUM(B1:B3)", v: "" } }, // empty string → needs compute
    });
    expect(emptyFormulaCells(s)).toEqual([
      { row: 1, column: 0, formula: "=SUM(A1:A3)" },
      { row: 2, column: 0, formula: "=SUM(B1:B3)" },
    ]);
  });

  it("NEVER includes formula cells with a cached value — including a genuine 0 (the #12 fix)", () => {
    const s = snap({
      1: { 0: { f: "=A1-A2", v: 0 } }, // cached zero total → MUST be preserved, not recomputed
      2: { 0: { f: "=SUM(C1:C9)", v: 1234 } }, // cached non-zero → preserved
      3: { 0: { v: 5 } }, // plain value, no formula
    });
    expect(emptyFormulaCells(s)).toEqual([]);
  });

  it("SKIPS uncached cross-sheet formulas (the #NAME? fix) — recomputing them against empty shells would render #NAME?/#REF!", () => {
    const s = snap({
      1: { 0: { f: "=SUM(A1:A3)" } }, // same-sheet, uncached → still recomputed
      2: { 0: { f: "='P&L'!B12" } }, // cross-sheet ref → skip
      3: { 0: { f: "=SUM(Detail!A:A)" } }, // cross-sheet range → skip
      4: { 0: { f: "=Balance!C1 + D1", v: "" } }, // cross-sheet + empty cache → skip
    });
    expect(emptyFormulaCells(s)).toEqual([{ row: 1, column: 0, formula: "=SUM(A1:A3)" }]);
  });
});
