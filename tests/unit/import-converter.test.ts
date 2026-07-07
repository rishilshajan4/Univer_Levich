/**
 * Import converter + Node-entry tests.
 *
 * The `.xlsx` → Univer-snapshot converter (`parseXlsxToSnapshot`) is the most
 * fidelity-critical module in the package and is now also consumed headlessly by
 * the backend via `@levichco/finsheets/node`. These round-trip tests build a real
 * workbook with exceljs, convert it, and assert the snapshot — guarding both the
 * converter and the Node-safe barrel against regressions.
 */
import { describe, it, expect } from "vitest";
import ExcelJS from "exceljs";
import { parseXlsxToSnapshot } from "../../src/core/xlsx-to-snapshot";
import { buildShellWorkbook } from "../../src/core/shell-workbook";
import { diffSheet } from "../../src/features/version-diff";
import * as nodeEntry from "../../src/node";

type FileLike = { name: string; arrayBuffer: () => Promise<ArrayBuffer> };

async function makeXlsx(build: (ws: ExcelJS.Worksheet) => void): Promise<FileLike> {
  const wb = new ExcelJS.Workbook();
  build(wb.addWorksheet("Sheet1"));
  // writeBuffer() returns a Node Buffer (a Uint8Array subclass) or a Uint8Array;
  // copy into a fresh, exactly-sized ArrayBuffer so exceljs/jszip can read it.
  const out = (await wb.xlsx.writeBuffer()) as unknown;
  const u8 = out instanceof Uint8Array ? out : new Uint8Array(out as ArrayBuffer);
  const ab = u8.slice().buffer;
  return { name: "test.xlsx", arrayBuffer: async () => ab };
}

/* eslint-disable @typescript-eslint/no-explicit-any */

describe("parseXlsxToSnapshot (import converter)", () => {
  it("converts values, formulas, and bold style", async () => {
    const file = await makeXlsx((ws) => {
      ws.getCell("A1").value = "Hello";
      ws.getCell("A1").font = { bold: true };
      ws.getCell("B1").value = 42;
      ws.getCell("A2").value = { formula: "B1*2", result: 84 } as ExcelJS.CellFormulaValue;
    });
    const snap = (await parseXlsxToSnapshot(file as unknown as File)) as any;
    const sid = snap.sheetOrder[0];
    const cd = snap.sheets[sid].cellData;

    expect(cd[0][0].v).toBe("Hello");
    expect(cd[0][1].v).toBe(42);
    expect(cd[1][0].f).toBe("=B1*2");

    const s = cd[0][0].s;
    const style = typeof s === "string" ? snap.styles[s] : s;
    expect(style?.bl).toBe(1); // bold flag
  });

  it("preserves a cell merge", async () => {
    const file = await makeXlsx((ws) => {
      ws.getCell("A1").value = "Merged";
      ws.mergeCells("A1:B2");
    });
    const snap = (await parseXlsxToSnapshot(file as unknown as File)) as any;
    const sid = snap.sheetOrder[0];
    expect(Array.isArray(snap.sheets[sid].mergeData)).toBe(true);
    expect(snap.sheets[sid].mergeData.length).toBeGreaterThan(0);
  });
});

describe("buildShellWorkbook", () => {
  it("hydrates only the active sheet; others are empty shells", () => {
    const wb = buildShellWorkbook({
      documentId: "doc", title: "T",
      manifest: [
        { order: 0, sheetId: "s1", name: "One", hidden: 0 },
        { order: 1, sheetId: "s2", name: "Two", hidden: 0 },
      ],
      activeSheetId: "s1",
      activeSnapshot: { sheets: { s1: { id: "s1", name: "One", cellData: { 0: { 0: { v: "x" } } } } }, styles: {}, resources: [] },
    }) as any;

    expect(wb.sheetOrder).toEqual(["s1", "s2"]);
    expect(wb.sheets.s1.cellData[0][0].v).toBe("x"); // active hydrated
    expect(wb.sheets.s2.cellData).toEqual({}); // shell empty
  });
});

describe("diffSheet (Highlight changes)", () => {
  it("detects a changed cell and reports noChanges when identical", () => {
    const doc = (v: unknown) => ({ manifest: [], title: "", sheets: { s1: { sheets: { s1: { cellData: { 0: { 0: { v } } } } } } } });
    const changed = diffSheet(doc(1) as any, doc(2) as any, "s1");
    expect(changed.changed.has("0:0")).toBe(true);
    expect(changed.noChanges).toBe(false);

    const same = diffSheet(doc(1) as any, doc(1) as any, "s1");
    expect(same.noChanges).toBe(true);
  });
});

describe("@levichco/finsheets/node entry", () => {
  it("re-exports the pure-data converter functions", () => {
    expect(typeof nodeEntry.parseXlsxToSnapshot).toBe("function");
    expect(typeof nodeEntry.buildExcelWorkbook).toBe("function");
    expect(typeof nodeEntry.buildShellWorkbook).toBe("function");
    expect(typeof nodeEntry.diffSheet).toBe("function");
    expect(typeof nodeEntry.highlightSnapshot).toBe("function");
  });
});
