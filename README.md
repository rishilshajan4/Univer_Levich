# @levichco/finsheets

A reusable, fully **Levich / FinOpz-branded** spreadsheet package that wraps the
free (Apache-2.0) [Univer](https://univer.ai) engine. Drop a Google-Sheets-style
spreadsheet into any React app with a single component — no Univer Pro, no server.

## Install

```bash
npm install @levichco/finsheets react react-dom
```

`react` and `react-dom` (>=18) are peer dependencies.

## Usage

```tsx
import { LevichSheet } from "@levichco/finsheets";
import "@levichco/finsheets/styles.css"; // once, anywhere in your app

export default function App() {
  const columns = [
    { key: "date", header: "Date", format: "date" },
    { key: "amount", header: "Amount", format: "currency" },
    { key: "memo", header: "Memo", editable: true },
  ];
  const data = [
    { date: "2026-01-16", amount: 4906.25, memo: "Opening balance" },
  ];

  return (
    <div style={{ height: "100vh" }}>
      <LevichSheet data={data} columns={columns} currencySymbol="$" />
    </div>
  );
}
```

## Features

- **Google-Sheets-style UI** — File / Edit / View / Insert / Format / Data menus,
  branded toolbar, formula bar, and sheet-tab management (add / delete / duplicate
  / rename / hide / colour / move, with a ▴ dropdown per tab).
- **Rich `.xlsx` import** — preserves cell styles, fills, fonts, theme colours,
  merges, number formats, dates, column widths, row heights, frozen panes,
  floating images, and **all worksheets** (hidden ones respected).
- **Full-fidelity `.xlsx` / CSV / TSV / HTML export**, plus print.
- **Filter & group views**, a full **function catalog** (500+ functions) with
  search, conditional formatting, data validation, notes, and hyperlinks.
- **Cross-platform keyboard shortcuts** (macOS ⌘ / Windows Ctrl) mirroring
  Excel & Google Sheets.
- **Configurable** columns, formats, freezing, locked columns, comments, and a
  computed pivot region — all driven by props.

## Public API (excerpt)

- `<LevichSheet data columns … />` — the single spreadsheet component.
- `parseXlsxToSnapshot(file)` — convert an `.xlsx` File to a Univer snapshot.
- `exportToXlsx(...)`, `LevichMenuBar`, `LevichToolbar`, `Modal`, `LEVICH_BRAND`.

See `dist/index.d.ts` for the full typed surface.

## Roadmap / design docs

- [`docs/VERSION-HISTORY.md`](docs/VERSION-HISTORY.md) — Google-Docs-style version
  history design (Yjs snapshots + PostgreSQL changeset log via a future
  `finsheets-service` backend). Currently the package is frontend-only; this doc
  specifies the collab/history hooks the FE will expose.

## License

Apache-2.0. Wraps Univer's Apache-2.0 free tier — see `NOTICE`.
