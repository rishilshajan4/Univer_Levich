import { describe, expect, it } from "vitest";
import { ROW_TOTAL, computePivotModel, renderPivotModel } from "../../src/features/pivot-model";
import type { PivotAggregate, PivotSource, PivotSpec } from "../../src/core/types";

describe("row Total with a BLANK column-field value (regression: must include the blank column)", () => {
    it("counts the blank-column data in the row Total + grand (not dropped)", () => {
        const src: PivotSource = {
            fields: ["region", "type", "amount"],
            rows: [
                { region: "X", type: "P", amount: 10 }, // column value present
                { region: "X", type: "", amount: 100 }, // BLANK column value → path ""
                { region: "Y", type: "P", amount: 5 },
            ],
        };
        const m = computePivotModel(src, { rows: ["region"], columns: ["type"], values: [{ field: "amount", aggregate: "sum" }] });
        const x = m.rowTree.find((n) => n.key === "X")!;
        // X row Total must be 10 + 100 = 110 (the blank-"type" 100 is NOT dropped).
        expect(x.values.get(`${ROW_TOTAL}␟0`)).toBe(110);
        // Grand row Total = 110 + 5 = 115.
        expect(m.grand.get(`${ROW_TOTAL}␟0`)).toBe(115);
        // The real "P" column still reads correctly (X=10).
        expect(x.values.get(`P␟0`)).toBe(10);
    });
});

const source: PivotSource = {
  fields: ["region", "product", "amount"],
  rows: [
    { region: "West", product: "A", amount: 100 },
    { region: "West", product: "A", amount: 50 },
    { region: "West", product: "B", amount: 30 },
    { region: "East", product: "A", amount: 200 },
    { region: "East", product: "B", amount: 20 },
  ],
};

describe("computePivotModel", () => {
  it("nests rows and sums per group + subtotals + grand total", () => {
    const spec: PivotSpec = { rows: ["region", "product"], columns: [], values: [{ field: "amount", aggregate: "sum" }] };
    const m = computePivotModel(source, spec);
    const west = m.rowTree.find((n) => n.key === "West")!;
    const total = (n: typeof west) => n.values.get(`${ROW_TOTAL}␟0`); // cellKey("", 0)
    expect(total(west)).toBe(180); // 100+50+30 — subtotal over underlying values
    const westA = west.children.find((n) => n.key === "A")!;
    expect(total(westA)).toBe(150);
    expect(m.grand.get(`${ROW_TOTAL}␟0`)).toBe(400); // 180 + 220
  });

  it("average TOTAL is over the union of underlying values, not an average-of-averages (Excel-exact)", () => {
    const spec: PivotSpec = { rows: ["region"], columns: ["product"], values: [{ field: "amount", aggregate: "average" }] };
    const m = computePivotModel(source, spec);
    const west = m.rowTree.find((n) => n.key === "West")!;
    // West: A avg = (100+50)/2 = 75; B avg = 30. Row TOTAL avg must be (100+50+30)/3 = 60,
    // NOT (75+30)/2 = 52.5.
    expect(west.values.get(`${ROW_TOTAL}␟0`)).toBeCloseTo(60, 6);
    // Grand average over all 5 rows = 400/5 = 80.
    expect(m.grand.get(`${ROW_TOTAL}␟0`)).toBeCloseTo(80, 6);
  });

  it("min/max ignore non-numeric cells (Excel-consistent — a stray 'N/A' must not poison the group)", () => {
    const src: PivotSource = {
      fields: ["region", "amount"],
      rows: [
        { region: "West", amount: 100 },
        { region: "West", amount: "N/A" },
        { region: "West", amount: 40 },
      ],
    };
    const min = computePivotModel(src, { rows: ["region"], columns: [], values: [{ field: "amount", aggregate: "min" }] });
    const max = computePivotModel(src, { rows: ["region"], columns: [], values: [{ field: "amount", aggregate: "max" }] });
    expect(min.rowTree[0].values.get(`${ROW_TOTAL}␟0`)).toBe(40); // not NaN
    expect(max.rowTree[0].values.get(`${ROW_TOTAL}␟0`)).toBe(100); // not NaN
  });

  it("supports multiple value fields with independent aggregations", () => {
    const spec: PivotSpec = {
      rows: ["region"],
      columns: [],
      values: [
        { field: "amount", aggregate: "sum" },
        { field: "amount", aggregate: "count" },
      ],
    };
    const m = computePivotModel(source, spec);
    const west = m.rowTree.find((n) => n.key === "West")!;
    expect(west.values.get(`${ROW_TOTAL}␟0`)).toBe(180); // sum
    expect(west.values.get(`${ROW_TOTAL}␟1`)).toBe(3); // count
  });
});

/** Brute-force reference: scan the raw rows for a (rowPath-prefix, colPath|null)
 *  group and aggregate directly — the O(n²) definition we optimise away. */
function bruteAgg(
  rows: Array<Record<string, unknown>>,
  rowFields: string[],
  colFields: string[],
  rowPrefix: string[] | null,
  colPath: string[] | null,
  field: string,
  agg: PivotAggregate,
): number {
  const nums: number[] = [];
  for (const r of rows) {
    if (rowPrefix) {
      let ok = true;
      for (let i = 0; i < rowPrefix.length; i++) if (String(r[rowFields[i]] ?? "") !== rowPrefix[i]) { ok = false; break; }
      if (!ok) continue;
    }
    if (colPath) {
      let ok = true;
      for (let i = 0; i < colPath.length; i++) if (String(r[colFields[i]] ?? "") !== colPath[i]) { ok = false; break; }
      if (!ok) continue;
    }
    const v = r[field];
    const n = typeof v === "number" ? v : Number(v);
    nums.push(Number.isFinite(n) ? n : NaN);
  }
  switch (agg) {
    case "count": return nums.length;
    case "countNumbers": return nums.filter(Number.isFinite).length;
    case "average": { const f = nums.filter(Number.isFinite); return f.length ? f.reduce((s, x) => s + x, 0) / f.length : 0; }
    case "min": { const f = nums.filter(Number.isFinite); return f.length ? Math.min(...f) : 0; }
    case "max": { const f = nums.filter(Number.isFinite); return f.length ? Math.max(...f) : 0; }
    default: return nums.reduce((s, x) => s + (Number.isFinite(x) ? x : 0), 0);
  }
}

describe("computePivotModel — roll-up accumulators", () => {
  it("matches a brute-force reference on a small multi-level, multi-col, multi-value case", () => {
    const src: PivotSource = {
      fields: ["region", "product", "year", "amount"],
      rows: [
        { region: "West", product: "A", year: "2023", amount: 100 },
        { region: "West", product: "A", year: "2024", amount: 50 },
        { region: "West", product: "B", year: "2023", amount: 30 },
        { region: "West", product: "B", year: "2024", amount: "N/A" },
        { region: "East", product: "A", year: "2023", amount: 200 },
        { region: "East", product: "A", year: "2024", amount: 20 },
        { region: "East", product: "B", year: "2023", amount: 5 },
      ],
    };
    const aggs: PivotAggregate[] = ["sum", "count", "average", "min", "max", "countNumbers"];
    for (const agg of aggs) {
      const spec: PivotSpec = { rows: ["region", "product"], columns: ["year"], values: [{ field: "amount", aggregate: agg }] };
      const m = computePivotModel(src, spec);
      const cols = m.colLeaves.filter((c) => c !== "");
      for (const region of m.rowTree) {
        // node level
        for (const col of cols) {
          expect(region.values.get(`${col}␟0`)).toBeCloseTo(bruteAgg(src.rows, spec.rows, spec.columns, [region.key], col.split("␟"), "amount", agg), 6);
        }
        // row Total (all columns)
        expect(region.values.get(`${ROW_TOTAL}␟0`)).toBeCloseTo(bruteAgg(src.rows, spec.rows, spec.columns, [region.key], null, "amount", agg), 6);
        for (const prod of region.children) {
          expect(prod.values.get(`${ROW_TOTAL}␟0`)).toBeCloseTo(bruteAgg(src.rows, spec.rows, spec.columns, [region.key, prod.key], null, "amount", agg), 6);
        }
      }
      // grand
      expect(m.grand.get(`${ROW_TOTAL}␟0`)).toBeCloseTo(bruteAgg(src.rows, spec.rows, spec.columns, null, null, "amount", agg), 6);
    }
  });

  it("computes a 5k-row × 3-level × 2-col × 2-value pivot in well under a second", () => {
    const rows: Array<Record<string, unknown>> = [];
    for (let i = 0; i < 5000; i++) {
      rows.push({
        region: `R${i % 4}`,
        product: `P${i % 10}`,
        sku: `S${i % 25}`,
        quarter: `Q${i % 4}`,
        channel: `C${i % 3}`,
        amount: (i % 97) + 1,
        qty: (i % 13) + 1,
      });
    }
    const src: PivotSource = { fields: Object.keys(rows[0]), rows };
    const spec: PivotSpec = {
      rows: ["region", "product", "sku"],
      columns: ["quarter", "channel"],
      values: [
        { field: "amount", aggregate: "sum" },
        { field: "qty", aggregate: "average" },
      ],
    };
    const t0 = performance.now();
    const m = computePivotModel(src, spec);
    const ms = performance.now() - t0;
    expect(m.rowTree.length).toBe(4);
    expect(ms).toBeLessThan(1000);

    // Spot-check exactness against brute force on one deep node + grand.
    const first = m.rowTree[0].children[0].children[0];
    expect(first.values.get(`${ROW_TOTAL}␟0`)).toBeCloseTo(
      bruteAgg(rows, spec.rows, spec.columns, [m.rowTree[0].key, m.rowTree[0].children[0].key, first.key], null, "amount", "sum"),
      6,
    );
    expect(m.grand.get(`${ROW_TOTAL}␟1`)).toBeCloseTo(bruteAgg(rows, spec.rows, spec.columns, null, null, "qty", "average"), 6);
  });
});

describe("renderPivotModel", () => {
  it("renders header + grouped rows + grand total into cells", () => {
    const spec: PivotSpec = { rows: ["region", "product"], columns: [], values: [{ field: "amount", aggregate: "sum" }] };
    const out = renderPivotModel(computePivotModel(source, spec));
    // Grand-total label is present in column 0 of the last row.
    const lastRow = out.cells[out.rowCount - 1];
    expect(lastRow[0]?.v).toBe("Grand Total");
    // Grand-total value cell exists and equals 400.
    const gtVal = Object.values(lastRow).find((c) => typeof c?.v === "number");
    expect(gtVal?.v).toBe(400);
    expect(out.columnCount).toBeGreaterThan(1);
  });

  it("collapsed groups hide children but keep the group row", () => {
    const spec: PivotSpec = { rows: ["region", "product"], columns: [], values: [{ field: "amount", aggregate: "sum" }], collapsed: ["West"] };
    const full = renderPivotModel(computePivotModel(source, { ...spec, collapsed: [] }));
    const collapsed = renderPivotModel(computePivotModel(source, spec));
    expect(collapsed.rowCount).toBeLessThan(full.rowCount); // West's children (A,B) + subtotal are hidden
  });
});

describe("computePivotModel — the extra Google aggregates + Show-as", () => {
  const src = {
    fields: ["g", "x"],
    rows: [
      { g: "A", x: 2 },
      { g: "A", x: 4 },
      { g: "A", x: 4 },
      { g: "B", x: 10 },
    ],
  };
  const model = (agg: string) => computePivotModel(src as never, { rows: ["g"], columns: [], values: [{ field: "x", aggregate: agg as never }] });
  const total = (m: ReturnType<typeof computePivotModel>, key: string) => m.rowTree.find((n) => n.key === key)!.values.get(`${ROW_TOTAL}␟0`);

  it("MEDIAN", () => { const m = model("median"); expect(total(m, "A")).toBe(4); }); // [2,4,4] → 4
  it("COUNTUNIQUE", () => { const m = model("countunique"); expect(total(m, "A")).toBe(2); }); // {2,4}
  it("PRODUCT", () => { const m = model("product"); expect(total(m, "A")).toBe(32); }); // 2*4*4
  it("STDEVP (population)", () => { const m = model("stdevp"); expect(total(m, "A")).toBeCloseTo(0.9428, 3); });
  it("VAR (sample)", () => { const m = model("var"); expect(total(m, "A")).toBeCloseTo(4 / 3, 6); });
  it("STDEV of a single value is 0 (n<2)", () => { const m = model("stdev"); expect(total(m, "B")).toBe(0); });

  it("Show-as % of grand total re-bases each row's total", () => {
    const m = computePivotModel(src as never, { rows: ["g"], columns: [], values: [{ field: "x", aggregate: "sum", showAs: "pctOfGrand" }] });
    const rendered = renderPivotModel(m);
    // A sum=10, B sum=10, grand=20 → each row total is 0.5 (50%).
    const cellVals = Object.values(rendered.cells).flatMap((row) => Object.values(row)).map((c) => c?.v);
    expect(cellVals).toContain(0.5);
  });
});

describe("empty / rows-only pivots (Google-Sheets behavior — no invented COUNT, no phantom Grand Total)", () => {
  it("a fully-empty spec renders NOTHING (no invented count, no Grand Total artifact)", () => {
    const empty: PivotSpec = { rows: [], columns: [], values: [] };
    const region = renderPivotModel(computePivotModel(source, empty));
    expect(region).toEqual({ cells: {}, rowCount: 0, columnCount: 0 });
  });

  it("rows-only (no values) groups by the row field but renders NO value cells / no numeric Grand Total", () => {
    const spec: PivotSpec = { rows: ["region"], columns: [], values: [] };
    const m = computePivotModel(source, spec);
    // Grouping still happens (distinct regions), but the model carries no invented value.
    expect(m.rowTree.map((n) => n.key).sort()).toEqual(["East", "West"]);
    expect(m.values.length).toBe(0);
    const region = renderPivotModel(m);
    // Only the row-label column exists (col 0); no value/Grand-Total columns.
    expect(region.columnCount).toBe(1);
    // No "Grand Total" row is emitted when there are no values to total.
    const hasGrandTotal = Object.values(region.cells).some((row) => Object.values(row).some((c) => (c as { v?: unknown }).v === "Grand Total"));
    expect(hasGrandTotal).toBe(false);
  });
});
