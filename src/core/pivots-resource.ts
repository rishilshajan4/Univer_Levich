/**
 * Durable persistence for reconstructed interactive pivots.
 *
 * `parseXlsxToSnapshot` reconstructs any imported pivot tables into an array of
 * {@link ImportedPivot} (source + spec + location) and exposes them as the
 * NON-Univer top-level `pivotsImport` key on the parsed snapshot. That key is an
 * escape hatch — Univer ignores unknown top-level keys and, crucially, DROPS them
 * when the workbook round-trips through save/load. So an imported pivot could only
 * be "opened interactively" in the SAME browser session that did the import.
 *
 * To make it durable, we ALSO stash the same array inside the snapshot's
 * `resources` array under the custom name {@link PIVOTS_IMPORT_RESOURCE}. Univer
 * preserves unknown `resources` entries verbatim across save + load (they're
 * opaque `{ name, data }` blobs keyed by plugin name), so a host that persists the
 * whole snapshot JSON — or that carries the workbook-level `resources` array in
 * its document manifest — gets the pivots back on ANY open, not just the import
 * session.
 *
 * This module is pure data (type-only imports) so it is safe for both the browser
 * bundle and the Node (`/node`) headless entry.
 */
import type { ImportedPivot } from "./pivot-import";

/**
 * The `resources[].name` under which reconstructed interactive pivots are stashed
 * so they round-trip through Univer save/load. Workbook-level (NOT keyed by sheet)
 * — the whole `ImportedPivot[]` is serialized as one JSON blob.
 */
export const PIVOTS_IMPORT_RESOURCE = "LEVICH_PIVOTS_IMPORT";

/** Minimal shape of a snapshot's `resources` array (opaque plugin blobs). */
type ResourceHolder = { resources?: Array<{ name: string; data: string }> } & Record<string, unknown>;

/**
 * Write `pivots` into `snapshot.resources` under {@link PIVOTS_IMPORT_RESOURCE}
 * (replacing any existing entry), so they survive a save/load round-trip. Mutates
 * and returns the same snapshot. No-op for an empty array (nothing to persist).
 */
export function attachPivotsResource<T extends ResourceHolder>(snapshot: T, pivots: ImportedPivot[]): T {
  if (!pivots.length) return snapshot;
  const resources: Array<{ name: string; data: string }> = (snapshot.resources ??= []);
  const data = JSON.stringify(pivots);
  const existing = resources.find((r) => r.name === PIVOTS_IMPORT_RESOURCE);
  if (existing) existing.data = data;
  else resources.push({ name: PIVOTS_IMPORT_RESOURCE, data });
  return snapshot;
}

/**
 * Read reconstructed interactive pivots back out of a loaded snapshot's
 * `resources` array. Returns `[]` if the resource is absent or unparseable — never
 * throws. Use this at editor-open time to decide whether to show the
 * "open imported pivot interactively" banner on ANY open of the document.
 */
export function readPivotsResource(snapshot: ResourceHolder | null | undefined): ImportedPivot[] {
  const entry = snapshot?.resources?.find((r) => r.name === PIVOTS_IMPORT_RESOURCE);
  if (!entry?.data) return [];
  try {
    const parsed = JSON.parse(entry.data);
    return Array.isArray(parsed) ? (parsed as ImportedPivot[]) : [];
  } catch {
    return [];
  }
}
