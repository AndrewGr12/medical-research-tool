# Localization audit and implementation notes

The audit references below are the line numbers in the copied source **before** this change. The source was audited before editing; the original repository was left untouched.

## Audit table

| File:line | What it produces or changes | Category | Render path |
|---|---|---|---|
| `main.js:18` | Translated column names before replacing underscores with spaces. The names come from uploaded headers and are user data. | C | Summary table, Table 1, detail dialog, figure controls/labels, CSV, Word, copy, SVG/PNG/PDF |
| `main.js:135,137,142` | Locale-dependent decimal and missing-count formatting. | B | Summary table, detail dialog, frequency tables, Table 1, CSV, Word, figure SVG/PNG/PDF |
| `main.js:203-204` | Status and error messages. | A | Status/error chrome |
| `main.js:283` | Reads uploaded file contents with `file.text()`; this is file ingestion, not UI localization. | C | Input only |
| `main.js:316-323` | File summary counts and French singular/plural labels. These are file-information chrome. | A | File summary panel |
| `main.js:346,364,691,927,947,1108,1187` | Locale-sensitive ordering for variable names, categories, groups, and figure labels. The ordered values are user data; the resulting order is part of analysis output. | B / C | Summary table, frequency table, Table 1, figure labels/legends |
| `main.js:390` | Summary-table counts, missing text, variable names, and table cells. | B | On-screen summary table |
| `main.js:416-436` | Per-variable detail labels and statistics. | B | Detail dialog |
| `main.js:500,602-606` | Box-plot title and labels drawn on canvas. | B | Detail dialog figure; canvas export/preview |
| `main.js:610-611` | Translated summary-table headers and methods used by copy/download builders. | B | Copy, CSV, TSV |
| `main.js:696-699,712` | Category counts, percentages, denominators, and category summary text. | B | Categorical frequency tables |
| `main.js:921` | Exact-test limit in an explanatory message. | B | Table 1 p-value detail/title |
| `main.js:948-989` | French branch, Table 1 count formatting, missing-cell labels, rows, and values. | B | On-screen Table 1, CSV, Word, copy |
| `main.js:1002` | “Select variables” menu label and selected count. | A | Table 1 controls |
| `main.js:1011-1029` | `localeTable1()` translated Table 1 headings, rows, and notes; French also replaced the notes and missing-cell format. | B | On-screen Table 1, CSV, Word, copy |
| `main.js:1104,1115-1119` | Figure tick values, labels, accessible title, and axis text. | B | Figure SVG, preview, SVG/PNG/PDF exports |
| `main.js:1180,1187-1191` | Figure count labels, missingness labels, and descriptions. | B / C | Figure SVG, preview, SVG/PNG/PDF exports |
| `main.js:1219-1223` | Figure type names and figure-control labels. Variable options are uploaded column names. | A / C | Figure controls (`#figtypes`, `#figctl`) |
| `main.js:1229,1233,1235` | SVG markup, empty-state text, and figure description were translated after construction. | B | Figure preview and SVG/PNG/PDF exports |
| `main.js:1305-1310` | Descriptive-statistics and Table 1 Word export titles, table labels, values, and notes. | B | Word exports |
| `main.js:1447-1452` | Disabled manuscript-copy code also contained locale-dependent output. | B | Dormant copy action |
| `locales.js:93,198-219,257` | `preserve()` excluded `#preview`, raw data, and language controls, while `translatePage()` and its document-wide observer translated other DOM text and attributes. | A/B boundary | Initial render and every later DOM mutation |

### Output container preservation before the change

Before the change, `preserve()` covered `#preview`, `[data-raw]`, `.language-options`, and `[data-locale]`. Of the named output containers, only `#preview` was inside that exclusion. `#stats`, `#t1`, `#t1notes`, `#types`, `#figplot`, `#figdesc`, and the detail dialog content were outside it. The summary and Table 1 captions, table-wrapper `aria-label`s, and the box-plot canvas `aria-label` were also outside it. The detail title had a runtime `data-raw` marker, but the statistics body and canvas did not.

## Root-cause finding

The reported root cause is confirmed. The page-wide `MutationObserver` called `translatePage()` for changes anywhere below `document.body`; the prior `preserve()` list did not protect most output containers. In addition, `main.js` translated variable names, summary/table output, detail labels, figure canvas/SVG text, figure descriptions, and Word exports directly. Output number formatting and several sort orders also followed the selected locale, and Table 1 had explicit French-only branches.

The copy now marks output containers and their captions/accessible labels with `data-output`. Both the initial translation pass and the observer-backed pass skip those subtrees. Summary and Table 1 render/export paths share their builders; figure preview and image/PDF export use the same English SVG. Output number formatting uses `OUT_LOCALE = 'en-US'`, and output collation uses English. The French Table 1 branches are gone. Output-only dictionary keys and regex translations are excluded; UI controls, warnings, and status/errors remain localizable. Uploaded labels and values are not passed through localization.

## Verification

Ran from this directory:

```text
node --check main.js
node --check locales.js
node ux-smoke-tests.cjs
All UX package smoke checks passed.
node statistics-tests.cjs
All grouped-comparison statistics checks passed.
```

The UX smoke test loads the actual output builders in a VM harness using the sample data, varies all five language settings, and compares summary rows, Table 1 rows/notes, figure SVG/descriptions, CSV/TSV, and Word bytes. It also checks output markers, observer protection, fixed formatting/collation, and retained language entries.

## Uncertain or limited areas

- The data preview and variable-types sections are commented out in the current HTML. Their containers are marked for future use, but there is no live preview renderer to exercise.
- File-summary chrome counts continue to format according to the selected UI locale; analysis outputs use the fixed English locale.
- The checks exercise builders and static DOM protection, not a full browser, screen-reader, or visual QA pass.
