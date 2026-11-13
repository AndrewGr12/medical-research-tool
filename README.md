# Medical Data Analysis AG

This is a copy of the existing clinical data analysis website with one concise introductory panel added to explain what the tool does. The analysis workflow, scripts, and other page content are preserved.

Open `index.html` in a browser, or serve this folder as a static website. No build step is required.

The browser-based tool supports descriptive summaries, group comparisons, tables, and figures. CSV data is processed locally in the browser. Excel import loads SheetJS from its CDN and requires internet access.

During upload, after the participant ID choice, a second dialog asks which columns are categorical or define groups. Automatically detected categorical variables start selected; numeric-looking flags such as CKD coded 0/1 can be selected to override detection. The selected columns are treated as categorical for summaries, Table 1, and figure grouping. This changes only the analysis type; uploaded values are not rewritten. The grouping choice applies to the current upload.

## Existing checks

Run with Node.js installed:

```sh
node --check main.js
node --check locales.js
node ux-smoke-tests.cjs
node statistics-tests.cjs
```
