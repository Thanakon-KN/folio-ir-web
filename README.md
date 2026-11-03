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
- Updates HTML table rows when a period header and row label match workbook data. Updates chart `period` / `dataset` arrays when the series label matches a workbook row. Quarterly arrays update `QoQ` where that property is present.
- Preserves the surrounding source text and leaves uncertain rows unchanged with a report note.

### Major Shareholders

- Scans workbook sheets for Overview, Free float, Graph, and shareholder list sections.
- Updates Overview / Free float values when a recognized workbook label is followed by its value in the same row and the HTML has a matching label/value pair.
- Updates matching rows in a standard HTML shareholder table and supported `majorShareholders` chart arrays from an explicit Graph sheet.
- Reports graph percentage totals that differ from 100% without normalizing the source data.

## Deterministic mapping limits

Unlike the desktop AI workflow, this edition does not interpret arbitrary page markup or workbook layouts. Financial tables need period columns and row labels that match the workbook. Chart arrays need a nearby `label` / `name` / `title` that matches a workbook row and `period` / `dataset` arrays. Shareholder updates require recognizable table headers, explicit Graph data, and label/value pairs for Overview / Free float. Unsupported or ambiguous layouts are reported for manual review instead of guessed.

The file is read and processed in the current browser tab. Only source code drafts are kept in browser local storage; selected workbooks stay in memory and must be selected again after reload.
