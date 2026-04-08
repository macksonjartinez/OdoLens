# OdoLens

OdoLens is a client-side React application for consolidating multiple clinical Excel workbooks into a single formatted output file.

The app accepts multiple `.xlsx` files, validates their worksheet structure, merges the relevant data, and generates a new consolidated workbook while preserving the original Excel template formatting, drawings, and layout.

## What The Project Does

- Uploads multiple hospital or clinic reporting files in `.xlsx` format.
- Validates that every file contains the required worksheets.
- Consolidates row-based data across:
  - `Nacionalidad`
  - `Aseguramiento`
  - `Consolidado`
- Aggregates specific cells in `Informe Económico`.
- Merges doctor columns in `Monitoreo`, including dynamic insertion of new doctor columns when needed.
- Generates a downloadable Excel file named `OdoLens_Consolidado_YYYY-MM-DD.xlsx`.
- Preserves the original Excel template styling by editing the workbook XML directly instead of rebuilding the file from scratch.

## Product Workflow

1. The user drags or selects one or more `.xlsx` files.
2. Files are filtered client-side:
   - only `.xlsx`
   - maximum 50 MB per file
   - duplicate files are ignored by `name + size`
3. Each workbook is parsed with SheetJS.
4. Each workbook is validated before consolidation.
5. Valid files are consolidated into a single in-memory result.
6. The embedded Excel template is updated and downloaded locally.

There is no backend. Files stay in the browser session.

## Validation Rules

Each uploaded workbook must contain these sheets:

- `Nacionalidad`
- `Aseguramiento`
- `Consolidado`
- `Informe Económico`
- `Monitoreo`

Additional validation rules:

- Sheet-name matching is case-insensitive, accent-insensitive, and trimmed.
- `Monitoreo` must contain a row whose first cell is `Nombre del Dr`.
- `Monitoreo` must expose at least one doctor column.
- Damaged or unreadable files are marked invalid.

## Consolidation Rules

### `Nacionalidad`, `Aseguramiento`, `Consolidado`

- The first row is treated as the header row.
- The first column is treated as the row label.
- Rows are matched by normalized label.
- Numeric values are summed across all uploaded files.
- The first-seen display label is preserved in the final output.

### `Informe Económico`

Summed cells:

- `E13`
- `F13`
- `G13`
- `H13`

Copied cells from the first non-empty source:

- `I13`
- `J13`
- `D15`
- `J15`

### `Monitoreo`

- The row starting with `Nombre del Dr` is used as the doctor header row.
- Non-doctor headers such as `Total`, `Procedimiento`, `Actividad`, `Código`, and similar labels are excluded from doctor matching.
- Doctor names are normalized before merge, so minor casing, spacing, and accent differences do not create duplicates.
- Placeholder doctor labels like `Dr 1`, `Dra 2`, etc. are treated specially and separated per source file/column.
- When the consolidated doctor count exceeds the template’s available doctor columns, new columns are inserted into the worksheet XML.
- Total formulas are regenerated so the template keeps working after doctor-column expansion.

## Why The Excel Output Keeps Its Formatting

The export layer does not rebuild the workbook from plain JSON.

Instead, `src/lib/excelWriter.ts` loads an embedded `.xlsx` template as a ZIP archive, updates only the targeted worksheet XML values, recalculates simple formula caches where needed, removes `calcChain.xml`, and writes the workbook back out.

That approach preserves:

- styles
- merged cells
- drawings
- images
- printer settings
- worksheet layout
- most original workbook internals

The embedded template lives in `src/assets/template.ts`.

## Tech Stack

- React 19
- TypeScript
- Vite 6
- Tailwind CSS 4
- SheetJS (`xlsx`) for workbook parsing
- JSZip for binary-level workbook editing
- FileSaver for downloads

## Project Structure

```text
src/
  assets/
    template.ts              Embedded Excel template as base64
  components/
    dropzone/                Upload UI
    filelist/                File status and validation errors
    layout/                  Header and footer
    success/                 Post-export screen
    summary/                 File counters and consolidate action
  constants/
    sheetSchema.ts           Required sheets and upload constraints
  hooks/
    useConsolidation.ts      End-to-end parsing, validation, export flow
    useFileDropzone.ts       File filtering and drag/drop handling
  lib/
    consolidator.ts          Workbook merge rules
    excelReader.ts           SheetJS workbook parsing
    excelWriter.ts           Template-preserving workbook writer
    normalize.ts             Name normalization and numeric parsing helpers
    validator.ts             Structural workbook validation
  types/
    index.ts                 Shared types
```

## Local Development

### Prerequisites

- Node.js
- npm

### Install

```bash
npm install
```

### Start The App

```bash
npm run dev
```

### Build For Production

```bash
npm run build
```

### Preview The Production Build

```bash
npm run preview
```

## Docker

The repo includes a production Docker setup.

### Build The Image

```bash
docker build -t odolens-local .
```

### Run The Container

```bash
docker run --rm -p 8081:80 odolens-local
```

Then open `http://localhost:8081`.

## Notes And Limitations

- This is a frontend-only application. There is no persistence layer or API.
- Output quality depends on the embedded template and the expected worksheet layout remaining stable.
- Formula recalculation is intentionally limited to the template patterns used by this project.
- The header exposes `Informe Estadístico` and `Informe Ejecutivo` tabs, but the current consolidation flow is shared and does not branch by tab yet.
- The repository currently stores the Excel template as generated source code in `src/assets/template.ts`, which keeps deployment simple but makes the file large.

## Useful Commands

```bash
npm run dev
npm run build
npm run preview
docker build -t odolens-local .
docker run --rm -p 8081:80 odolens-local
```

## Suggested Next Improvements

- Move the embedded template into a clearer asset-generation workflow.
- Add automated tests around workbook validation and consolidation.
- Add fixture-based regression tests for template-preserving export behavior.
- Wire the report-type tabs to distinct business flows if those reports diverge.
