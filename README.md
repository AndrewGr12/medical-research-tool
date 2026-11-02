# Data Analysis Assistant — UX optimized copy

This folder is a separate static copy of the website. It can be opened by opening `index.html` in a current browser or served as a static site. No install or build step is required.

## UX review and changes in v2

The existing site already provides an understandable start (language choice, sample data, and upload), file-reading and analysis feedback, patient-ID exclusion, a persistent Tables/Figures switcher, descriptive and grouped outputs, exports, and keyboard/reduced-motion support. This version keeps those workflows and the statistical calculations intact.

## Language behavior

The selected language applies to instructions and interface controls. Analysis output stays in English in the summary table, Table 1, detail dialog, figure labels and descriptions, and exports. Uploaded column names and category values are kept as provided. Output numbers and sorting use fixed English conventions; outputs stay left-to-right when the interface is Arabic.

Output DOM subtrees carry `data-output`, which keeps both initial translation and later DOM mutations from changing their text or accessible labels. The summary and Table 1 export paths reuse their table builders, and figure exports reuse the SVG shown in the preview.

The review found two practical improvements: blank or duplicate CSV/Excel column names were automatically renamed without telling users, and helper/chart descriptions were smaller than the main interface text. The app now displays a translated note when it normalizes headers, and helper/chart text has been enlarged for easier reading. It also retains the named keyboard-scroll regions and live warning announcements added in the previous optimized copy.

## External dependency

CSV files work without a network connection. Excel workbook support loads SheetJS 0.20.3 from `https://cdn.sheetjs.com/xlsx-0.20.3/package/dist/xlsx.full.min.js` when an Excel file is selected, so Excel import requires internet access. Workbook contents are parsed in the browser and are not sent to the CDN. Local processing alone does not ensure HIPAA or institutional compliance.

## Included files

- `index.html`, `style.css`, `main.js`, `locales.js`: application entry point and runtime.
- `statistics-tests.cjs`: existing grouped-comparison statistical regression checks.
- `ux-smoke-tests.cjs`: structural checks plus language-invariance checks for table rows, notes, figures, and export text across en/fr/es/ru/ar.
- `AUDIT.md`: pre-change localization audit, root-cause finding, and test record.

Unused or unrelated project assets were not copied. The author page and its image are disabled in the source UI.

## Verification

Run these commands from this folder with Node.js installed:

```sh
node --check main.js
node --check locales.js
node ux-smoke-tests.cjs
node statistics-tests.cjs
```

The checks do not replace testing in a browser or screen reader. Interactive visual, keyboard-only, screen-reader, and cross-browser checks remain to be done manually; no accessibility-conformance claim is made.

## Figure customization upgrade

The Figures view keeps its existing page layout and SVG-based rendering while adding more controls:

- Histograms support Sturges, square-root, or Freedman–Diaconis automatic bin rules, plus a manual bin count.
- Box plots support per-category color pickers when grouped; each color is used for both the box and its legend item. With no group selected, a single-series color picker is available. The legend control stays disabled until a group is selected.
- Grouped box plots can display one overall exploratory p-value, using Welch or Kruskal–Wallis (Mann–Whitney U for two groups). Pairwise Holm p-values, group ordering, and individual-point controls are currently disabled.
- Bar charts support count or percentage display and category ordering. Scatter plots support grouping colors, point styling, and an optional linear trend line.
- Figure titles and axis titles, color palettes, opacity, gridlines, and legend visibility are customizable where they apply.

The standalone implementation copy remains at `C:\Users\16179\figure-studio-upgrade-v2`. The app changes copied into this repository are the production files listed above; figure-only test coverage has been added to the existing statistics and UX smoke-test scripts.
