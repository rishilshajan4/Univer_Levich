/**
 * Reconstruct an INTERACTIVE pivot (`PivotSource` + `PivotSpec`) from an imported
 * `.xlsx` file, so an uploaded pivot table — which otherwise renders as static
 * cells — can be made drag-editable via `LevichSheet`'s `pivotInteractive` prop.
 *
 * Excel stores a pivot table as several XML parts inside the .xlsx (a ZIP):
 *   - `xl/pivotTables/pivotTable*.xml` — the LAYOUT: which fields are on rows /
 *     columns / values (`<rowFields>`, `<colFields>`, `<dataFields>`), each data
 *     field's `subtotal=` (the aggregation), and the on-sheet `location ref`.
 *   - `xl/pivotCache/pivotCacheDefinition*.xml` — the field NAMES + a
 *     `<cacheSource><worksheetSource ref="Sheet1!A1:F200"/>` pointer to the
 *     SOURCE range (or `<consolidation>` / external, which we don't handle).
 *   - `xl/pivotCache/pivotCacheRecords*.xml` — the cached source RECORDS, present
 *     only when the file was saved with "save source data" (`saveData=1`).
 *
 * PRIMARY strategy: read the layout + `worksheetSource ref`, then SLICE that
 * range out of the already-ExcelJS-parsed workbook snapshot (every sheet's
 * `cellData` is in hand) → header row = fields, following rows = records.
 *
 * FALLBACK: if there is no in-workbook source but the cache embedded its records
 * (`saveData=1`), parse `pivotCacheRecords*.xml` (+ shared items from the cache
 * definition) into records.
 *
 * If neither yields a usable source, we return nothing and the caller leaves the
 * static render in place.
 *
 * jszip + saxes are transitive deps of exceljs; both are dynamic-imported (like
 * exceljs itself in xlsx-to-snapshot.ts) so nothing is pulled in unless an
 * actual pivot import happens.
 */
import type { PivotAggregate, PivotSource, PivotSpec, PivotValueField, WorkbookData } from "./types";

/* -------------------------------------------------------------------------- */
/* Parsed-XML intermediate shapes                                             */
/* -------------------------------------------------------------------------- */

/** One field reference inside `<rowFields>` / `<colFields>` (`x="<index>"`). An
 *  index of -2 is Excel's special "data" placeholder (the Σ Values marker) and
 *  is skipped for our purposes. */
type FieldRef = number;

interface DataField {
  /** 0-based index into the cache fields. */
  fld: number;
  /** Aggregation subtotal (`sum` / `count` / `average` / `min` / `max` / …). */
  subtotal: string;
  /** Optional display name from the file. */
  name?: string;
}

interface PivotTableXml {
  /** Cache field names, in cache order (index → name). */
  cacheFields: string[];
  rowFields: FieldRef[];
  colFields: FieldRef[];
  dataFields: DataField[];
  /** On-sheet anchor of the pivot's top-left cell, 0-based. */
  location: { row: number; column: number } | null;
}

interface CacheDefXml {
  /** Field names in cache order. */
  fields: string[];
  /** Per-field shared string/number items (for records that reference by index). */
  sharedItems: Array<Array<string | number | boolean | null>>;
  /** `Sheet!A1:F200` source range, split into sheet + range, if in-workbook. */
  worksheetSource: { sheet: string; ref: string } | null;
  /** Whether the cache stored its own records (`saveData`). */
  saveData: boolean;
}

/* -------------------------------------------------------------------------- */
/* Aggregate mapping                                                          */
/* -------------------------------------------------------------------------- */

/** Map an Excel pivot `subtotal=` value → our `PivotAggregate`. Excel's default
 *  (absent attribute) is `sum` for numeric fields. `countNums` → countNumbers;
 *  `count` (count of all, incl. text) → count. Unsupported (product/stdDev/var)
 *  fall back to `sum`. */
export function mapSubtotal(subtotal: string | undefined): PivotAggregate {
  switch (subtotal) {
    case "count": // countA — count of all values (incl. text)
      return "count";
    case "countNums":
      return "countNumbers";
    case "average":
      return "average";
    case "min":
      return "min";
    case "max":
      return "max";
    case "sum":
    default:
      return "sum";
  }
}

/* -------------------------------------------------------------------------- */
/* saxes parsing helpers                                                      */
/* -------------------------------------------------------------------------- */

interface SaxTag {
  name: string;
  attributes: Record<string, string>;
  isSelfClosing: boolean;
}

/** Minimal streaming walk over an XML string with saxes, invoking callbacks on
 *  open/close tags. Kept tiny so both parsers share one dynamic import. */
async function walkXml(
  xml: string,
  onOpen: (t: SaxTag) => void,
  onClose?: (name: string) => void,
): Promise<void> {
  const { SaxesParser } = await import("saxes");
  const parser = new SaxesParser();
  parser.on("opentag", (t: unknown) => {
    const tag = t as { name: string; attributes: Record<string, { value: string } | string>; isSelfClosing: boolean };
    const attrs: Record<string, string> = {};
    for (const [k, v] of Object.entries(tag.attributes)) attrs[k] = typeof v === "string" ? v : v.value;
    onOpen({ name: localName(tag.name), attributes: attrs, isSelfClosing: tag.isSelfClosing });
  });
  if (onClose) parser.on("closetag", (t: unknown) => onClose(localName((t as { name: string }).name)));
  parser.on("error", () => { /* tolerate malformed fragments — best-effort */ });
  parser.write(xml).close();
}

/** Strip an XML namespace prefix (`x:pivotField` → `pivotField`). */
const localName = (n: string): string => (n.includes(":") ? n.slice(n.indexOf(":") + 1) : n);

/* -------------------------------------------------------------------------- */
/* pivotTable*.xml                                                            */
/* -------------------------------------------------------------------------- */

/** Parse `location ref="B3:F20"` → 0-based top-left {row,column}. */
function parseLocationRef(ref: string | undefined): { row: number; column: number } | null {
  if (!ref) return null;
  const first = ref.split(":")[0];
  return parseCellAddr(first);
}

/** "AB12" → 0-based {row,column}. */
function parseCellAddr(addr: string): { row: number; column: number } | null {
  const m = /^([A-Z]+)(\d+)$/.exec(addr.trim().toUpperCase().replace(/\$/g, ""));
  if (!m) return null;
  let col = 0;
  for (const ch of m[1]) col = col * 26 + (ch.charCodeAt(0) - 64);
  return { row: Number(m[2]) - 1, column: col - 1 };
}

async function parsePivotTableXml(xml: string): Promise<Omit<PivotTableXml, "cacheFields">> {
  const rowFields: FieldRef[] = [];
  const colFields: FieldRef[] = [];
  const dataFields: DataField[] = [];
  let location: { row: number; column: number } | null = null;
  // Track which list we're inside so `<field x="..">` rows go to the right bucket.
  let inRow = false;
  let inCol = false;

  await walkXml(
    xml,
    (t) => {
      switch (t.name) {
        case "location":
          location = parseLocationRef(t.attributes.ref);
          break;
        case "rowFields":
          inRow = true;
          break;
        case "colFields":
          inCol = true;
          break;
        case "field": {
          const x = Number(t.attributes.x);
          if (Number.isFinite(x)) {
            if (inRow) rowFields.push(x);
            else if (inCol) colFields.push(x);
          }
          break;
        }
        case "dataField": {
          const fld = Number(t.attributes.fld);
          if (Number.isFinite(fld)) {
            dataFields.push({ fld, subtotal: t.attributes.subtotal ?? "sum", name: t.attributes.name });
          }
          break;
        }
      }
    },
    (name) => {
      if (name === "rowFields") inRow = false;
      else if (name === "colFields") inCol = false;
    },
  );

  return { rowFields, colFields, dataFields, location };
}

/* -------------------------------------------------------------------------- */
/* pivotCacheDefinition*.xml                                                  */
/* -------------------------------------------------------------------------- */

async function parseCacheDefinition(xml: string): Promise<CacheDefXml> {
  const fields: string[] = [];
  const sharedItems: Array<Array<string | number | boolean | null>> = [];
  let worksheetSource: { sheet: string; ref: string } | null = null;
  let saveData = false;
  // Depth tracking so item-list `<s v="">` entries attach to the current field.
  let curField = -1;
  let inSharedItems = false;

  await walkXml(
    xml,
    (t) => {
      switch (t.name) {
        case "pivotCacheDefinition":
          saveData = t.attributes.saveData !== "0"; // default is "1"/true
          break;
        case "worksheetSource": {
          // ref="A1:F200" sheet="Sheet1"  (name= for a defined-name source).
          const sheet = t.attributes.sheet ?? "";
          const ref = t.attributes.ref ?? "";
          if (ref) worksheetSource = { sheet, ref };
          break;
        }
        case "cacheField":
          fields.push(t.attributes.name ?? `Field${fields.length + 1}`);
          curField = fields.length - 1;
          sharedItems[curField] = [];
          inSharedItems = false;
          break;
        case "sharedItems":
          inSharedItems = true;
          break;
        case "s": // string item
          if (inSharedItems && curField >= 0) sharedItems[curField].push(t.attributes.v ?? "");
          break;
        case "n": // number item
          if (inSharedItems && curField >= 0) sharedItems[curField].push(Number(t.attributes.v));
          break;
        case "b": // boolean item
          if (inSharedItems && curField >= 0) sharedItems[curField].push(t.attributes.v === "1");
          break;
        case "m": // missing/blank item
          if (inSharedItems && curField >= 0) sharedItems[curField].push(null);
          break;
      }
    },
    (name) => {
      if (name === "sharedItems") inSharedItems = false;
    },
  );

  return { fields, sharedItems, worksheetSource, saveData };
}

/* -------------------------------------------------------------------------- */
/* pivotCacheRecords*.xml (fallback)                                          */
/* -------------------------------------------------------------------------- */

/** Parse the embedded records. Each `<r>` is a record; children are values in
 *  cache-field order: `<x v="i"/>` = shared-item index, `<n v=".."/>` = number,
 *  `<s v=".."/>` = inline string, `<b/>` = bool, `<m/>` = blank. */
async function parseCacheRecords(xml: string, def: CacheDefXml): Promise<Array<Record<string, unknown>>> {
  const records: Array<Record<string, unknown>> = [];
  let cur: unknown[] | null = null;
  let colIdx = 0;

  await walkXml(
    xml,
    (t) => {
      switch (t.name) {
        case "r":
          cur = [];
          colIdx = 0;
          break;
        case "x": // reference into the field's shared items
          if (cur) {
            const i = Number(t.attributes.v);
            const shared = def.sharedItems[colIdx] ?? [];
            cur.push(Number.isFinite(i) ? shared[i] ?? "" : "");
            colIdx++;
          }
          break;
        case "n":
          if (cur) { cur.push(Number(t.attributes.v)); colIdx++; }
          break;
        case "s":
          if (cur) { cur.push(t.attributes.v ?? ""); colIdx++; }
          break;
        case "b":
          if (cur) { cur.push(t.attributes.v === "1"); colIdx++; }
          break;
        case "m":
          if (cur) { cur.push(null); colIdx++; }
          break;
      }
    },
    (name) => {
      if (name === "r" && cur) {
        const rec: Record<string, unknown> = {};
        def.fields.forEach((f, i) => (rec[f] = cur![i] ?? null));
        records.push(rec);
        cur = null;
      }
    },
  );

  return records;
}

/* -------------------------------------------------------------------------- */
/* Slice the source range out of the parsed snapshot                          */
/* -------------------------------------------------------------------------- */

/** Read a value out of a snapshot cell (`{v}` literal, resolving nothing). */
function cellVal(cell: unknown): unknown {
  if (cell && typeof cell === "object") {
    const c = cell as { v?: unknown };
    return c.v ?? null;
  }
  return null;
}

/** Find a sheet's `cellData` by its display NAME (case-insensitive). */
function findSheetCellData(snapshot: WorkbookData, sheetName: string): Record<string, Record<string, unknown>> | null {
  const sheets = (snapshot.sheets as Record<string, { name?: string; cellData?: Record<string, Record<string, unknown>> }> | undefined) ?? {};
  const wanted = sheetName.trim().replace(/^'|'$/g, "").toLowerCase();
  for (const s of Object.values(sheets)) {
    const nm = (s.name ?? "").toLowerCase();
    if (nm === wanted) return s.cellData ?? null;
  }
  // If the pivot named no sheet (or a mismatch), fall back to the first sheet.
  const first = Object.values(sheets)[0];
  return first?.cellData ?? null;
}

/** Parse a range "A1:F200" → {startRow,startCol,endRow,endCol} 0-based. */
function parseRange(ref: string): { sr: number; sc: number; er: number; ec: number } | null {
  const [a, b] = ref.split(":");
  const start = parseCellAddr(a);
  const end = parseCellAddr(b ?? a);
  if (!start || !end) return null;
  return {
    sr: Math.min(start.row, end.row),
    sc: Math.min(start.column, end.column),
    er: Math.max(start.row, end.row),
    ec: Math.max(start.column, end.column),
  };
}

/** Slice a source range out of the snapshot → header row = fields, following
 *  rows = records. Trailing all-empty rows are dropped. */
function sliceSource(snapshot: WorkbookData, sheet: string, ref: string): PivotSource | null {
  const cellData = findSheetCellData(snapshot, sheet);
  if (!cellData) return null;
  const rng = parseRange(ref);
  if (!rng) return null;

  const at = (r: number, c: number): unknown => cellVal(cellData[String(r)]?.[String(c)]);

  const fields: string[] = [];
  for (let c = rng.sc; c <= rng.ec; c++) {
    const h = at(rng.sr, c);
    fields.push(h != null && h !== "" ? String(h) : `Field${c - rng.sc + 1}`);
  }
  if (!fields.length) return null;

  const rows: Array<Record<string, unknown>> = [];
  for (let r = rng.sr + 1; r <= rng.er; r++) {
    const rec: Record<string, unknown> = {};
    let hasAny = false;
    for (let c = rng.sc; c <= rng.ec; c++) {
      const v = at(r, c);
      rec[fields[c - rng.sc]] = v ?? null;
      if (v != null && v !== "") hasAny = true;
    }
    if (hasAny) rows.push(rec);
  }
  if (!rows.length) return null;
  return { fields, rows };
}

/* -------------------------------------------------------------------------- */
/* Build the PivotSpec from the parsed layout                                 */
/* -------------------------------------------------------------------------- */

function buildSpec(table: PivotTableXml): PivotSpec {
  const nameOf = (idx: number): string | null => (idx >= 0 && idx < table.cacheFields.length ? table.cacheFields[idx] : null);
  const rows = table.rowFields.map(nameOf).filter((n): n is string => !!n);
  const columns = table.colFields.map(nameOf).filter((n): n is string => !!n);
  const values: PivotValueField[] = table.dataFields
    .map((d): PivotValueField | null => {
      const field = nameOf(d.fld);
      if (!field) return null;
      return { field, aggregate: mapSubtotal(d.subtotal), label: d.name };
    })
    .filter((v): v is PivotValueField => !!v);
  return { rows, columns, values };
}

/* -------------------------------------------------------------------------- */
/* Public: parse pivots out of the raw .xlsx bytes                            */
/* -------------------------------------------------------------------------- */

/** One reconstructed interactive pivot: where it sat + how to re-render it live. */
export interface ImportedPivot {
  /** 0-based top-left cell of the pivot on its sheet (from `location ref`). */
  location: { row: number; column: number };
  source: PivotSource;
  spec: PivotSpec;
}

/**
 * Parse every pivot table out of a workbook's raw bytes + an already-built
 * snapshot (for the in-workbook source slice). Returns an empty array if there
 * are no pivots or none can be reconstructed robustly. Never throws.
 *
 * @param bytes    The raw .xlsx file bytes (`await file.arrayBuffer()`).
 * @param snapshot The snapshot from `parseXlsxToSnapshot` (holds every sheet's cellData).
 */
export async function parsePivotsFromXlsx(bytes: ArrayBuffer, snapshot: WorkbookData): Promise<ImportedPivot[]> {
  interface ZipFile { async(t: "text"): Promise<string> }
  interface Zip { files: Record<string, unknown>; file(path: string): ZipFile | null }
  let zip: Zip;
  try {
    const mod: unknown = await import("jszip");
    const JSZip = ((mod as { default?: unknown }).default ?? mod) as { loadAsync(b: ArrayBuffer): Promise<Zip> };
    zip = await JSZip.loadAsync(bytes);
  } catch (e) {
    console.warn("[levich] pivot-import: jszip load failed", e);
    return [];
  }

  const readText = async (path: string): Promise<string | null> => {
    const f = zip.file(path);
    return f ? f.async("text") : null;
  };

  // Enumerate pivotTable parts.
  const tablePaths = Object.keys(zip.files).filter((p) => /^xl\/pivotTables\/pivotTable\d+\.xml$/i.test(p)).sort();
  if (!tablePaths.length) return [];

  // Pre-load every cache definition + records once (a table maps to a cache via
  // the workbook rels, but we can robustly pair by trying each cache — most
  // workbooks have one cache per table and the field-name lists are consistent).
  const cacheDefPaths = Object.keys(zip.files).filter((p) => /^xl\/pivotCache\/pivotCacheDefinition\d+\.xml$/i.test(p)).sort();
  const caches: Array<{ def: CacheDefXml; recordsPath: string | null }> = [];
  for (const p of cacheDefPaths) {
    const xml = await readText(p);
    if (!xml) continue;
    const def = await parseCacheDefinition(xml);
    // Records file is the sibling with the same numeric suffix.
    const num = /(\d+)\.xml$/.exec(p)?.[1];
    const recordsPath = num ? `xl/pivotCache/pivotCacheRecords${num}.xml` : null;
    caches.push({ def, recordsPath: recordsPath && zip.file(recordsPath) ? recordsPath : null });
  }

  const out: ImportedPivot[] = [];

  for (const tp of tablePaths) {
    try {
      const tableXml = await readText(tp);
      if (!tableXml) continue;
      const parsed = await parsePivotTableXml(tableXml);

      // Pick the cache whose field set covers this table's referenced indices.
      // Prefer the one with an in-workbook worksheetSource; else any.
      const maxIdx = Math.max(-1, ...parsed.rowFields, ...parsed.colFields, ...parsed.dataFields.map((d) => d.fld));
      const cache =
        caches.find((c) => c.def.worksheetSource && c.def.fields.length > maxIdx) ??
        caches.find((c) => c.def.fields.length > maxIdx) ??
        caches[0];
      if (!cache) continue;

      const table: PivotTableXml = { ...parsed, cacheFields: cache.def.fields };
      const spec = buildSpec(table);
      if (!spec.rows.length && !spec.columns.length && !spec.values.length) continue;

      // PRIMARY: slice the in-workbook source range out of the snapshot.
      let source: PivotSource | null = null;
      if (cache.def.worksheetSource) {
        source = sliceSource(snapshot, cache.def.worksheetSource.sheet, cache.def.worksheetSource.ref);
      }
      // FALLBACK: embedded cache records (saveData=1, no in-workbook source).
      if (!source && cache.recordsPath) {
        const recXml = await readText(cache.recordsPath);
        if (recXml) {
          const records = await parseCacheRecords(recXml, cache.def);
          if (records.length) source = { fields: cache.def.fields, rows: records };
        }
      }
      if (!source) continue; // neither worked → leave the static render

      out.push({ location: parsed.location ?? { row: 0, column: 0 }, source, spec });
    } catch (e) {
      console.warn("[levich] pivot-import: failed to parse", tp, e);
    }
  }

  return out;
}
