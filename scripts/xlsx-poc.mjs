/**
 * PoC "backend" converter. Runs the REAL browser converter (parseXlsxToSnapshot)
 * in Node — proving it ports to a backend Processing Unit unchanged — then:
 *   - splits the workbook into ONE file per sheet:  public/poc-sheets/<id>.json
 *     (with per-sheet PRUNED styles, so each file is small)
 *   - writes the tab manifest:                      public/poc-manifest.json
 *
 * The demo (?poc, demo/poc-app.tsx) then loads only the manifest + the active
 * sheet, fetching others on tab-click. This is exactly what the real service
 * would serve from POST /import + GET /documents/:id/sheets/:id.
 *
 * Usage: node scripts/xlsx-poc.mjs [file.xlsx]
 */
import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { parseXlsxToSnapshot } from "../src/core/xlsx-to-snapshot.ts";

const ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const SRC = process.argv[2] || "/Users/rishilshajan/Downloads/1783032677785-8r0gngya5mu.xlsx";
const OUT = path.join(ROOT, "public");
const SHEETS_DIR = path.join(OUT, "poc-sheets");

const buf = fs.readFileSync(SRC);
const arrayBuffer = buf.buffer.slice(buf.byteOffset, buf.byteOffset + buf.byteLength);
const fakeFile = { name: path.basename(SRC), arrayBuffer: async () => arrayBuffer };

const t0 = Date.now();
const full = await parseXlsxToSnapshot(fakeFile);
console.log(`convert: ${Date.now() - t0} ms · ${full.sheetOrder.length} sheets · ${Object.keys(full.styles || {}).length} styles`);

// Collect the style ids actually referenced by a sheet (cellData / row / col /
// default) so each per-sheet file carries only its own styles, not all 4000+.
function collectStyleIds(sheet) {
  const ids = new Set();
  const add = (s) => { if (typeof s === "string") ids.add(s); };
  const cd = sheet.cellData || {};
  for (const r in cd) for (const c in cd[r]) add(cd[r][c]?.s);
  const rd = sheet.rowData || {}; for (const r in rd) add(rd[r]?.s);
  const colD = sheet.columnData || {}; for (const c in colD) add(colD[c]?.s);
  add(sheet.defaultStyle);
  return ids;
}
function prunedStyles(ids) {
  const out = {};
  for (const id of ids) if (full.styles?.[id]) out[id] = full.styles[id];
  return out;
}
function sheetResources(sheetId) {
  // Every Univer resource here is keyed by sheetId (drawings, filters, …) —
  // keep just this sheet's slice of each.
  const out = [];
  for (const r of full.resources || []) {
    try { const d = JSON.parse(r.data); if (d[sheetId]) out.push({ name: r.name, data: JSON.stringify({ [sheetId]: d[sheetId] }) }); } catch { /* */ }
  }
  return out;
}
// VIEW MODE: strip only CROSS-SHEET formulas; KEEP same-sheet ones.
// A formula referencing ANOTHER (unloaded) sheet — e.g. OpEx Detail pulling from
// Chart Data — can't recompute here (#VALUE!), so we drop it and keep the cached
// value (Excel-on-open behaviour). But a SAME-SHEET formula CAN recompute — and
// must, because some have no cached result in the file: e.g. B3
// =COUNTIFS($A:$A,"# Check") caches nothing, so stripping it left a blank that
// tripped its "if <>0 turn red" conditional format. Keeping it → Univer computes
// 0 → no red, matching Excel.
function hasCrossSheetRef(f) {
  // Ignore error literals (#REF!, #VALUE!, …), then a remaining "!" is a sheet ref.
  return f.replace(/#[A-Z0-9/]+!/gi, "").includes("!");
}
function stripFormulasForView(sheet) {
  const cd = sheet?.cellData;
  if (!cd) return;
  for (const r in cd) for (const c in cd[r]) {
    const cell = cd[r][c];
    if (cell && cell.f !== undefined && hasCrossSheetRef(cell.f)) delete cell.f; // keep cell.v
  }
}
function singleSheetSnapshot(sheetId) {
  const sheet = full.sheets[sheetId];
  stripFormulasForView(sheet);
  return {
    id: sheetId,
    name: sheet?.name ?? sheetId,
    sheetOrder: [sheetId],
    styles: prunedStyles(collectStyleIds(sheet || {})),
    sheets: { [sheetId]: sheet },
    resources: sheetResources(sheetId),
  };
}

fs.rmSync(SHEETS_DIR, { recursive: true, force: true });
fs.mkdirSync(SHEETS_DIR, { recursive: true });

const manifest = [];
let totalBytes = 0, biggest = { name: "", kb: 0 };
full.sheetOrder.forEach((id, i) => {
  const sheet = full.sheets[id];
  const snap = singleSheetSnapshot(id);
  const p = path.join(SHEETS_DIR, `${id}.json`);
  fs.writeFileSync(p, JSON.stringify(snap));
  const kb = fs.statSync(p).size / 1024;
  totalBytes += kb * 1024;
  if (kb > biggest.kb) biggest = { name: sheet?.name ?? id, kb: Math.round(kb) };
  manifest.push({ order: i, sheetId: id, name: sheet?.name ?? id, hidden: sheet?.hidden ? 1 : 0 });
});

fs.writeFileSync(path.join(OUT, "poc-manifest.json"), JSON.stringify(manifest, null, 2));

console.log(`wrote ${manifest.length} per-sheet files → public/poc-sheets/  (${(totalBytes / 1024 / 1024).toFixed(2)} MB total, pruned styles)`);
console.log(`biggest sheet: "${biggest.name}" ${biggest.kb} KB`);
console.log(`wrote public/poc-manifest.json (${manifest.length} tabs)`);
