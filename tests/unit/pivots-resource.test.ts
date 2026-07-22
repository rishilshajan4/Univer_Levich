/**
 * pivots-resource tests — durable persistence of reconstructed interactive pivots
 * via a `LEVICH_PIVOTS_IMPORT` entry in the snapshot's `resources` array (survives
 * a Univer save/load round-trip, unlike the top-level `pivotsImport` escape hatch).
 */
import { describe, it, expect } from "vitest";
import { PIVOTS_IMPORT_RESOURCE, attachPivotsResource, readPivotsResource } from "../../src/core/pivots-resource";
import type { ImportedPivot } from "../../src/core/pivot-import";

const PIVOTS: ImportedPivot[] = [
  {
    location: { row: 2, column: 1 },
    source: { fields: ["region", "amount"], rows: [["West", 10]] },
    spec: { rows: ["region"], columns: [], values: [{ field: "amount", aggregate: "sum" }] },
  },
];

describe("pivots-resource", () => {
  it("attaches pivots under the LEVICH_PIVOTS_IMPORT resource name", () => {
    const snap: { resources?: Array<{ name: string; data: string }> } = {};
    attachPivotsResource(snap, PIVOTS);
    const entry = snap.resources?.find((r) => r.name === PIVOTS_IMPORT_RESOURCE);
    expect(entry).toBeTruthy();
    expect(JSON.parse(entry!.data)).toEqual(PIVOTS);
  });

  it("preserves existing resources and is idempotent (replaces, never duplicates)", () => {
    const snap = { resources: [{ name: "SHEET_DRAWING_PLUGIN", data: "{}" }] };
    attachPivotsResource(snap, PIVOTS);
    attachPivotsResource(snap, PIVOTS);
    expect(snap.resources.filter((r) => r.name === PIVOTS_IMPORT_RESOURCE)).toHaveLength(1);
    expect(snap.resources.find((r) => r.name === "SHEET_DRAWING_PLUGIN")).toBeTruthy();
  });

  it("is a no-op for an empty pivot array (nothing to persist)", () => {
    const snap: { resources?: Array<{ name: string; data: string }> } = {};
    attachPivotsResource(snap, []);
    expect(snap.resources).toBeUndefined();
  });

  it("round-trips: readPivotsResource recovers what attachPivotsResource wrote", () => {
    const snap: { resources?: Array<{ name: string; data: string }> } = {};
    attachPivotsResource(snap, PIVOTS);
    // simulate save/load: serialize whole snapshot and parse back
    const reloaded = JSON.parse(JSON.stringify(snap));
    expect(readPivotsResource(reloaded)).toEqual(PIVOTS);
  });

  it("readPivotsResource returns [] when the resource is absent or malformed", () => {
    expect(readPivotsResource(null)).toEqual([]);
    expect(readPivotsResource({})).toEqual([]);
    expect(readPivotsResource({ resources: [] })).toEqual([]);
    expect(readPivotsResource({ resources: [{ name: PIVOTS_IMPORT_RESOURCE, data: "not json" }] })).toEqual([]);
    expect(readPivotsResource({ resources: [{ name: PIVOTS_IMPORT_RESOURCE, data: '{"not":"array"}' }] })).toEqual([]);
  });
});
