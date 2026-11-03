# Folio IR Web

Browser based version of the Folio IR workflow. It keeps the two workspaces, input/output editors, period selection, add/replace modes, preview, report, theme controls, local drafts, and export flow. Workbook reading and updates use deterministic JavaScript functions. The page does not call AI, an API, or a backend.

## Run or publish

Open `index.html` in a modern browser. To publish with GitHub Pages, place the contents of this folder at the repository root and select **Settings → Pages → Deploy from a branch**. The `xlsx.full.min.js` file is required for Excel workbook reading.

## Workflows

### Financial Highlights

- Reads XLSX, XLS, XLSM, CSV, and TSV files locally in the browser.
- Detects period headers and numeric rows arranged with periods in columns.
- Inspects periods found in the workbook and in the existing HTML / JavaScript.
- Requires an explicit period selection and Add or Replace choice.
- Matches table rows and chart series by normalized labels, common aliases, and token similarity; prefers detailed financial rows for tables and summary rows for charts.
- Filters and aligns nested yearly / quarterly chart `period` and `dataset` arrays, updates mapped named series, and sets existing `QoQ` flags based on selected quarterly values (including zero).
- Updates numeric `<counter-number>` KPI attributes from the latest selected period when their nearby heading maps to a source metric.
- Applies explicit period selection first: Add keeps old values for selected overlaps; Replace uses workbook values; selected old-only periods remain. Deselected periods are removed from aligned tables and charts.
- Preserves the surrounding source text and leaves uncertain rows unchanged with a report note.

### Major Shareholders

- Scans workbook sheets for Overview, Free float, Graph, and shareholder list sections.
- Updates Overview / Free float values when a recognized workbook label is followed by its value in the same row and the HTML has a matching label/value pair.
- Updates matching shareholder rows, adds source-listed holders missing from a recognizable HTML table, and updates supported `majorShareholders` / `shareholders` chart arrays from an explicit Graph sheet.
- Reports graph percentage totals that differ from 100% without normalizing the source data.

## Deterministic mapping limits

The browser version uses explicit parsing and matching rules instead of the desktop model's broader interpretation. It handles common Excel period columns, HTML tables, KPI counter components, nested chart source objects, and recognized shareholder sections. It cannot safely infer every custom Vue component, irregular workbook layout, or site-specific data mapping. Those cases are reported for review; an actual workbook and its matching HTML/Vue and JavaScript are needed to add targeted rules without guessing.

The file is read and processed in the current browser tab. Only source code drafts are kept in browser local storage; selected workbooks stay in memory and must be selected again after reload.
