/**
 * Tests for the interactive pivot panel: the pure spec-mutation helpers (the
 * real logic) and a jsdom render asserting the field chips + drop areas exist
 * and that dropping / removing a field fires onChange with the right spec.
 */
import { describe, expect, it, vi } from "vitest";
import { fireEvent, render, screen } from "@testing-library/react";

// The @untitledui/icons ESM barrel re-exports extensionless paths that Node's
// ESM resolver (used by vitest) can't resolve. The panel only needs a tiny icon
// glyph, so stub the package with a trivial component — keeps the test focused
// on behavior, not the icon set.
vi.mock("@untitledui/icons", () => ({
  X: (props: Record<string, unknown>) => <span data-icon="x" {...props} />,
  Plus: (props: Record<string, unknown>) => <span data-icon="plus" {...props} />,
  ChevronDown: (props: Record<string, unknown>) => <span data-icon="chevron" {...props} />,
}));
import {
  PivotPanel,
  areaOfField,
  fieldsInArea,
  placeField,
  removeField,
  removeFromArea,
  clearAll,
  setValueAggregate,
  aggregateOfValue,
} from "../../src/features/pivot-panel";
import type { PivotSpec } from "../../src/core/types";

const baseSpec: PivotSpec = {
  rows: ["Region"],
  columns: [],
  values: [{ field: "Amount", aggregate: "sum" }],
};

describe("pivot-panel spec helpers", () => {
  it("reports the fields in each area", () => {
    const spec: PivotSpec = { rows: ["Region"], columns: ["Quarter"], values: [{ field: "Amount", aggregate: "sum" }], filters: [{ field: "Year" }] };
    expect(fieldsInArea(spec, "rows")).toEqual(["Region"]);
    expect(fieldsInArea(spec, "columns")).toEqual(["Quarter"]);
    expect(fieldsInArea(spec, "values")).toEqual(["Amount"]);
    expect(fieldsInArea(spec, "filters")).toEqual(["Year"]);
  });

  it("finds the area a field is placed in", () => {
    expect(areaOfField(baseSpec, "Region")).toBe("rows");
    expect(areaOfField(baseSpec, "Amount")).toBe("values");
    expect(areaOfField(baseSpec, "Nope")).toBeNull();
  });

  it("places a field into an area (and removes it from its old one)", () => {
    const next = placeField(baseSpec, "Region", "columns");
    expect(next.rows).toEqual([]);
    expect(next.columns).toEqual(["Region"]);
  });

  it("defaults a fresh Values field to sum, preserving an existing aggregate on move", () => {
    const added = placeField(baseSpec, "Units", "values");
    expect(added.values.find((v) => v.field === "Units")?.aggregate).toBe("sum");

    const withAvg = setValueAggregate(added, "Units", "average");
    // Reordering within Values keeps the average aggregate.
    const reordered = placeField(withAvg, "Units", "values", 0);
    expect(reordered.values[0].field).toBe("Units");
    expect(aggregateOfValue(reordered, "Units")).toBe("average");
  });

  it("reorders within an area by index", () => {
    const spec: PivotSpec = { rows: ["A", "B", "C"], columns: [], values: [] };
    const moved = placeField(spec, "C", "rows", 0);
    expect(moved.rows).toEqual(["C", "A", "B"]);
  });

  it("removes a field from every area", () => {
    expect(removeField(baseSpec, "Region").rows).toEqual([]);
    expect(removeField(baseSpec, "Amount").values).toEqual([]);
  });

  it("changes a value field's aggregate", () => {
    const next = setValueAggregate(baseSpec, "Amount", "count");
    expect(next.values[0].aggregate).toBe("count");
  });

  it("ADDs a field to Values while keeping it in Rows (Google-Sheets multi-area)", () => {
    // No `moveFrom` = a fields-list ADD: Amount lives in Values already; putting Region in
    // Values must NOT pull Region out of Rows.
    const next = placeField(baseSpec, "Region", "values");
    expect(next.rows).toEqual(["Region"]);
    expect(fieldsInArea(next, "values")).toEqual(["Amount", "Region"]);
  });

  it("MOVEs a field out of its source section when dragged from a chip", () => {
    // `moveFrom: "rows"` = the chip was dragged from Rows → it should vacate Rows.
    const next = placeField(baseSpec, "Region", "values", undefined, "rows");
    expect(next.rows).toEqual([]);
    expect(fieldsInArea(next, "values")).toEqual(["Amount", "Region"]);
  });

  it("keeps Rows and Columns mutually exclusive", () => {
    const next = placeField(baseSpec, "Region", "columns"); // even without moveFrom
    expect(next.rows).toEqual([]);
    expect(next.columns).toEqual(["Region"]);
  });

  it("removeFromArea drops a field from ONE area only", () => {
    const both = placeField(baseSpec, "Region", "values"); // Region now in rows + values
    const afterRowsX = removeFromArea(both, "Region", "rows");
    expect(afterRowsX.rows).toEqual([]);
    expect(fieldsInArea(afterRowsX, "values")).toEqual(["Amount", "Region"]); // still summed
  });

  it("clearAll empties every area", () => {
    const full: PivotSpec = { rows: ["A"], columns: ["B"], values: [{ field: "C", aggregate: "sum" }], filters: [{ field: "D" }], dimSettings: { A: { order: "desc" } } };
    const empty = clearAll(full);
    expect(empty.rows).toEqual([]);
    expect(empty.columns).toEqual([]);
    expect(empty.values).toEqual([]);
    expect(empty.filters).toEqual([]);
  });
});

describe("<PivotPanel />", () => {
  const fields = ["Region", "Quarter", "Amount"];

  it("renders the four sections each with an Add button", () => {
    render(<PivotPanel fields={fields} spec={baseSpec} onChange={() => {}} />);
    for (const a of ["filters", "columns", "rows", "values"]) {
      expect(screen.getByTestId(`area-${a}`)).toBeTruthy();
      expect(screen.getByTestId(`add-${a}`)).toBeTruthy();
    }
  });

  it("adds a field to a section via the Add menu", () => {
    const onChange = vi.fn();
    render(<PivotPanel fields={fields} spec={baseSpec} onChange={onChange} />);
    fireEvent.click(screen.getByTestId("add-columns")); // open the Columns Add menu
    fireEvent.click(screen.getByTestId("add-field-columns-Quarter")); // pick Quarter
    expect(onChange).toHaveBeenCalledTimes(1);
    expect(onChange.mock.calls[0][0].columns).toEqual(["Quarter"]);
  });

  it("fires onChange removing a field when its card × is clicked", () => {
    const onChange = vi.fn();
    render(<PivotPanel fields={fields} spec={baseSpec} onChange={onChange} />);
    fireEvent.click(screen.getByLabelText("Remove Region"));
    expect(onChange).toHaveBeenCalledTimes(1);
    expect(onChange.mock.calls[0][0].rows).toEqual([]);
  });

  it("changes the aggregate via the Summarize-by DS dropdown", () => {
    const onChange = vi.fn();
    render(<PivotPanel fields={fields} spec={baseSpec} onChange={onChange} />);
    fireEvent.click(screen.getByLabelText("Summarize Amount")); // open the Select
    fireEvent.click(screen.getByText("AVERAGE")); // pick AVERAGE
    expect(onChange).toHaveBeenCalledTimes(1);
    expect(onChange.mock.calls[0][0].values[0].aggregate).toBe("average");
  });

  it("shows a value's Show-as options", () => {
    const onChange = vi.fn();
    render(<PivotPanel fields={fields} spec={baseSpec} onChange={onChange} />);
    fireEvent.click(screen.getByLabelText("Show Amount as")); // open Show-as Select
    fireEvent.click(screen.getByText("% of grand total"));
    expect(onChange).toHaveBeenCalledTimes(1);
    expect(onChange.mock.calls[0][0].values[0].showAs).toBe("pctOfGrand");
  });

  it("renders a persistent Fields rail with a draggable chip per source field", () => {
    render(<PivotPanel fields={fields} spec={baseSpec} onChange={() => {}} />);
    for (const f of fields) {
      const chip = screen.getByTestId(`field-${f}`);
      expect(chip).toBeTruthy();
      expect(chip.getAttribute("draggable")).toBe("true"); // draggable even when already placed
    }
  });

  it("offers a Sort-by control on a Rows card listing the value fields", () => {
    const onChange = vi.fn();
    render(<PivotPanel fields={fields} spec={baseSpec} onChange={onChange} />);
    fireEvent.click(screen.getByLabelText("Sort Region by")); // open the Sort-by Select
    fireEvent.click(screen.getByText("Sum of Amount")); // sort Region groups by the Amount total
    expect(onChange).toHaveBeenCalledTimes(1);
    expect(onChange.mock.calls[0][0].dimSettings?.Region?.sortBy).toBe("Amount");
  });

  it("Clear all resets the pivot to empty", () => {
    const onChange = vi.fn();
    render(<PivotPanel fields={fields} spec={baseSpec} onChange={onChange} />);
    fireEvent.click(screen.getByTestId("pivot-clear-all"));
    expect(onChange).toHaveBeenCalledTimes(1);
    expect(onChange.mock.calls[0][0].rows).toEqual([]);
    expect(onChange.mock.calls[0][0].values).toEqual([]);
  });
});
