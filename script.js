/* =========================================================================
   AI in Medical Research – beginner-friendly descriptive statistics
   -------------------------------------------------------------------------
   Everything runs in the browser. No data leave the page.

   Sections
     1. Constants and state
     2. Small utilities (DOM, formatting, downloads, messages)
     3. File import (CSV parser, Excel reader, dataset builder)
     4. Variable profiling and type detection
     5. Statistics (mean, SD, quartiles, box-plot statistics)
     6. Dataset-quality summary and variable-review table
     7. Table 1 (model, rendering)
     8. Charts (histogram, bar chart, custom box plot)
     9. Plain-English interpretation
    10. Exports (CSV, Excel, PNG, HTML report)
    11. Wiring: event listeners and initialisation
   ========================================================================= */
(function () {
  'use strict';

  /* ======================================================================
     1. CONSTANTS AND STATE
     ====================================================================== */
  const LIMITS = { fileBytes: 25 * 1024 * 1024, rows: 100000, cols: 300 };
  const PREVIEW_ROWS = 10;
  const MAX_BAR_CATEGORIES = 20; // more categories than this make a bar chart illegible
  const MAX_BOX_GROUPS = 12;
  const MIN_BOX_N = 3; // smallest group that gets a box plot

  // Text values treated as "missing" in addition to blank cells. Disclosed to the user.
  const MISSING_TOKENS = new Set(['na', 'n/a', '#n/a', 'nan', 'null']);
  // Numbers that are often used as missing-value codes. We only WARN about these; never auto-convert.
  const COMMON_MISSING_CODES = [-9999, -999, -99, -9, -1, 99, 999, 9999];

  const TYPE_LABELS = {
    continuous: 'Continuous numeric',
    binary: 'Binary categorical',
    multicategory: 'Multicategory categorical',
    identifier: 'Identifier',
    date: 'Date',
    text: 'Text / unclassified'
  };
  const TYPE_ORDER = ['continuous', 'binary', 'multicategory', 'identifier', 'date', 'text'];
  const ANALYZABLE = new Set(['continuous', 'binary', 'multicategory']);

  const DELIMITERS = { comma: ',', semicolon: ';', tab: '\t', pipe: '|' };

  const FOOTNOTE =
    'Continuous variables are presented as mean (SD) or median (IQR), as selected. Categorical variables are presented as n (%). ' +
    'Missing observations are excluded from variable-specific calculations unless otherwise specified.';
  const FOOTNOTE_EXTRA =
    'SD = sample standard deviation (n \u2212 1 denominator); IQR = interquartile range, shown as Q1\u2013Q3. Quartiles use linear interpolation ' +
    '(the same method as Excel QUARTILE.INC and R type 7). Percentages use the number of observations with a recorded value for that variable ' +
    'and may not sum to exactly 100% because of rounding. For continuous variables, \u201cMissing\u201d also counts entries that are not valid numbers.';

  const COLORS = {
    navy: '#13315C',
    teal: '#0E7C86',
    tealFill: 'rgba(14, 124, 134, 0.72)',
    tealLight: 'rgba(14, 124, 134, 0.22)',
    text: '#1B2733',
    grid: '#E3E8EE'
  };

  // Fictional educational data (26 patients). null = missing value.
  const SAMPLE_GRID = [
    ['Patient_ID', 'Age', 'Sex', 'BMI', 'Hemoglobin', 'Diabetes', 'Hypertension', 'Treatment_Group', 'Death_1_Year'],
    ['P001', 62, 'Male', 28.4, 13.8, 'Yes', 'Yes', 'Drug A', 'No'],
    ['P002', 55, 'Female', 31.2, 12.1, 'No', 'Yes', 'Drug B', 'No'],
    ['P003', 71, 'Male', 26.7, 11.9, 'Yes', 'Yes', 'Placebo', 'Yes'],
    ['P004', 48, 'Female', 23.5, 13.2, 'No', 'No', 'Drug A', 'No'],
    ['P005', 66, 'Female', null, 12.6, 'Yes', 'Yes', 'Drug B', 'No'],
    ['P006', 59, 'Male', 29.9, 14.5, 'No', 'Yes', 'Placebo', 'No'],
    ['P007', 74, 'Female', 24.8, 10.7, 'Yes', 'Yes', 'Drug A', 'Yes'],
    ['P008', 52, 'Male', 33.1, 14.9, 'Yes', 'No', 'Drug B', 'No'],
    ['P009', 45, 'Female', 22.3, 12.8, 'No', 'No', 'Placebo', 'No'],
    ['P010', 68, 'Male', 27.5, null, 'No', 'Yes', 'Drug A', 'No'],
    ['P011', 80, 'Female', 25.1, 10.2, 'Yes', 'Yes', 'Placebo', 'Yes'],
    ['P012', 57, 'Male', 30.6, 13.5, 'Yes', 'Yes', 'Drug B', 'No'],
    ['P013', 63, 'Female', 35.2, 11.4, 'Yes', 'Yes', 'Drug A', 'No'],
    ['P014', 41, 'Male', 24.0, 15.1, 'No', 'No', 'Placebo', 'No'],
    ['P015', 69, 'Female', 28.8, 12.0, 'No', 'Yes', 'Drug B', 'No'],
    ['P016', 77, 'Male', 21.9, 11.1, 'No', 'Yes', 'Drug A', 'Yes'],
    ['P017', 54, 'Female', 27.3, 12.9, 'Yes', 'No', 'Placebo', 'No'],
    ['P018', 60, 'Male', 32.4, 13.7, 'Yes', 'Yes', 'Drug B', null],
    ['P019', 50, 'Female', 26.2, 12.5, 'No', 'No', 'Drug A', 'No'],
    ['P020', 72, 'Male', 29.1, 10.9, 'Yes', 'Yes', 'Placebo', 'Yes'],
    ['P021', 58, 'Female', 30.0, null, 'No', 'Yes', 'Drug B', 'No'],
    ['P022', 46, 'Male', 25.6, 14.2, 'No', 'No', 'Drug A', 'No'],
    ['P023', 65, 'Female', 34.0, 11.8, 'Yes', 'Yes', 'Placebo', 'No'],
    ['P024', 53, 'Male', 27.9, 13.9, 'No', null, 'Drug B', 'No'],
    ['P025', 70, 'Female', 23.8, 11.5, 'No', 'Yes', 'Drug A', 'No'],
    ['P026', 61, 'Male', 31.5, 12.7, 'Yes', 'Yes', 'Placebo', 'No']
  ];

  const state = {
    source: null,        // { name, kind: 'csv' | 'excel' | 'sample' }
    csvText: null,       // kept so import options can re-parse
    workbook: null,      // SheetJS workbook for Excel files
    sheetName: null,
    delimiterChoice: 'auto',
    grid: null,          // normalized rows of cells (null = missing)
    parseNotes: [],      // notes produced while reading the file
    hasHeader: true,
    extraMissing: [],    // user-defined missing-value codes (strings)
    headers: [],
    rows: [],            // row-major, normalized, missing = null
    columns: [],         // see buildDataset()
    dupRows: 0,
    buildNotes: [],
    confirmed: false,
    stale: false,
    settings: { summary: 'mean', decimals: 1, showCounts: true },
    analysis: null,      // computed statistics for included variables (snapshot at "generate")
    table1: null,        // table model derived from analysis + settings
    charts: { hist: null, bar: null, box: null },
    chartText: { hist: '', bar: '', box: '' },
    chartNotes: { hist: '', bar: '', box: '' }
  };

  // Counters filled while cells are normalised during import.
  let importStats = { tokenMissing: 0, excelErrors: 0 };

  /* ======================================================================
     2. SMALL UTILITIES
     ====================================================================== */
  function $(id) { return document.getElementById(id); }

  /** Safe element builder. Text is always added through text nodes (never as HTML). */
  function el(tag, props, ...children) {
    const node = document.createElement(tag);
    if (props) {
      for (const [key, value] of Object.entries(props)) {
        if (value === false || value === null || value === undefined) continue;
        if (key.startsWith('on')) continue; // never allow inline handlers
        if (key === 'class') node.className = value;
        else if (key === 'text') node.textContent = value;
        else if (key === 'dataset') Object.assign(node.dataset, value);
        else if (value === true) node.setAttribute(key, '');
        else node.setAttribute(key, String(value));
      }
    }
    for (const child of children) {
      if (child === null || child === undefined) continue;
      node.append(child instanceof Node ? child : document.createTextNode(String(child)));
    }
    return node;
  }

  function clear(node) { while (node.firstChild) node.removeChild(node.firstChild); }
  function plural(n, one, many) { return n === 1 ? one : (many || one + 's'); }
  function fmtInt(n) { return Number(n).toLocaleString('en-US'); }
  function clamp(x, lo, hi) { return Math.min(hi, Math.max(lo, x)); }
  function prefersReducedMotion() { return window.matchMedia && window.matchMedia('(prefers-reduced-motion: reduce)').matches; }

  /** Round half away from zero with a tiny epsilon to dodge binary floating-point artefacts (e.g. 2.675). */
  function roundTo(x, d) {
    const f = Math.pow(10, d);
    const r = Math.round(Math.abs(x) * f + 1e-9) / f;
    return x < 0 ? -r : r;
  }
  /** Format a number with exactly d decimals; non-finite values show an em dash. */
  function fmt(x, d) {
    if (typeof x !== 'number' || !Number.isFinite(x)) return '\u2014';
    let r = roundTo(x, d);
    if (r === 0) r = 0; // avoid "-0"
    return r.toFixed(d);
  }
  function fmtPct(count, total, d) { return total > 0 ? fmt((count / total) * 100, d) + '%' : '\u2014'; }
  function fmtRange(a, b, d) {
    const joiner = (a < 0 || b < 0) ? ' to ' : '\u2013';
    return fmt(a, d) + joiner + fmt(b, d);
  }
  function displayName(name) { return String(name).replace(/_/g, ' '); }
  function truncate(s, n) { s = String(s); return s.length > n ? s.slice(0, n - 1) + '\u2026' : s; }
  function safeBaseName() {
    const raw = state.source ? state.source.name : 'dataset';
    return raw.replace(/\.[^.]+$/, '').replace(/[^A-Za-z0-9_-]+/g, '_').replace(/^_+|_+$/g, '').slice(0, 60) || 'dataset';
  }

  /* ---------- Messages ---------- */
  function messageNode(kind, text, closable) {
    const label = { error: 'Error', warn: 'Warning', info: 'Note', success: 'Done' }[kind] || 'Note';
    const node = el('div', { class: 'msg msg-' + kind, role: kind === 'error' ? 'alert' : 'status' },
      el('strong', { text: label + ': ' }), document.createTextNode(text));
    if (closable) {
      const btn = el('button', { type: 'button', class: 'msg-close', 'aria-label': 'Dismiss message', text: '\u00d7' });
      btn.addEventListener('click', () => node.remove());
      node.append(btn);
    }
    return node;
  }
  function setMessages(containerId, items) {
    const box = $(containerId);
    clear(box);
    items.forEach((m) => box.append(messageNode(m.kind, m.text)));
  }
  function addMessage(containerId, kind, text) { $(containerId).append(messageNode(kind, text)); }
  /** Global notification (export errors, library problems). Errors stay until dismissed. */
  function notify(kind, text) {
    const box = $('globalMessages');
    while (box.children.length >= 3) box.removeChild(box.firstChild);
    const node = messageNode(kind, text, true);
    box.append(node);
    if (kind !== 'error') setTimeout(() => node.remove(), 7000);
  }

  function setLoading(on, text) {
    $('loadingStatus').hidden = !on;
    if (text) $('loadingText').textContent = text;
    document.body.setAttribute('aria-busy', on ? 'true' : 'false');
  }

  /* ---------- Downloads ---------- */
  function triggerDownload(href, filename) {
    const a = document.createElement('a');
    a.href = href;
    a.download = filename;
    a.rel = 'noopener';
    document.body.append(a);
    a.click();
    a.remove();
  }
  function downloadBlob(blob, filename) {
    const url = URL.createObjectURL(blob);
    triggerDownload(url, filename);
    setTimeout(() => URL.revokeObjectURL(url), 2000);
  }

  /* ======================================================================
     3. FILE IMPORT
     ====================================================================== */

  /** Turn any raw cell into null (missing), a finite number, or a trimmed string. */
  function normalizeCell(v) {
    if (v === null || v === undefined) return null;
    if (v instanceof Date) return Number.isNaN(v.getTime()) ? null : dateToIso(v);
    if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE';
    if (typeof v === 'number') return Number.isFinite(v) ? v : String(v); // keep Infinity as text so it is flagged, not used
    const s = String(v).trim();
    if (s === '') return null;
    if (MISSING_TOKENS.has(s.toLowerCase())) { importStats.tokenMissing++; return null; }
    return s;
  }

  function dateToIso(d) {
    // SheetJS dates can be a few seconds off; nudge before reading local components.
    const t = new Date(d.getTime() + 30000);
    const p = (n) => String(n).padStart(2, '0');
    let s = t.getFullYear() + '-' + p(t.getMonth() + 1) + '-' + p(t.getDate());
    if (t.getHours() || t.getMinutes()) s += ' ' + p(t.getHours()) + ':' + p(t.getMinutes());
    return s;
  }

  function readFile(file, mode) {
    return new Promise((resolve, reject) => {
      const fr = new FileReader();
      fr.onload = () => resolve(fr.result);
      fr.onerror = () => reject(fr.error || new Error('The browser could not read this file.'));
      if (mode === 'text') fr.readAsText(file, 'UTF-8'); else fr.readAsArrayBuffer(file);
    });
  }

  /* ---------- CSV ---------- */
  function countOutsideQuotes(line, delim) {
    let inQ = false, count = 0;
    for (let i = 0; i < line.length; i++) {
      const ch = line[i];
      if (ch === '"') inQ = !inQ;
      else if (!inQ && ch === delim) count++;
    }
    return count;
  }

  /**
   * Guess the delimiter from the first ~20 logical lines. For each candidate we measure how many
   * lines have the same (non-zero) number of delimiters; consistency wins, then count.
   */
  function detectDelimiter(text) {
    const lines = [];
    let cur = '', inQ = false;
    const limit = Math.min(text.length, 200000);
    for (let i = 0; i < limit && lines.length < 20; i++) {
      const ch = text[i];
      if (ch === '"') inQ = !inQ;
      if (!inQ && ch === '\n') { if (cur.trim()) lines.push(cur); cur = ''; }
      else if (!(ch === '\r' && !inQ)) cur += ch;
    }
    if (cur.trim() && lines.length < 20) lines.push(cur);
    let best = ',', bestScore = 0;
    for (const d of Object.values(DELIMITERS)) {
      const counts = lines.map((l) => countOutsideQuotes(l, d));
      const first = counts[0] || 0;
      if (first === 0) continue;
      const consistency = counts.filter((c) => c === first).length / counts.length;
      const score = consistency * 1000 + first;
      if (score > bestScore) { best = d; bestScore = score; }
    }
    return best;
  }

  /** RFC-4180-style parser: quoted fields, escaped quotes (""), delimiters and newlines inside quotes. */
  function parseCsv(text, delim) {
    const rows = [];
    let row = [], field = '', inQ = false, i = 0;
    const n = text.length;
    const pushRow = () => {
      row.push(field); rows.push(row); row = []; field = '';
      if (rows.length > LIMITS.rows + 2) throw new Error(tooManyRowsMessage());
    };
    while (i < n) {
      const ch = text[i];
      if (inQ) {
        if (ch === '"') {
          if (text[i + 1] === '"') { field += '"'; i += 2; continue; }
          inQ = false; i++; continue;
        }
        field += ch; i++; continue;
      }
      if (ch === '"' && field === '') { inQ = true; i++; continue; }
      if (ch === delim) { row.push(field); field = ''; i++; continue; }
      if (ch === '\r') { if (text[i + 1] === '\n') i++; pushRow(); i++; continue; }
      if (ch === '\n') { pushRow(); i++; continue; }
      field += ch; i++;
    }
    const unterminated = inQ;
    if (field !== '' || row.length > 0) pushRow();
    return { rows, unterminated };
  }

  function tooManyRowsMessage() {
    return 'This file has more than ' + fmtInt(LIMITS.rows) + ' rows, which is more than this tool can analyze reliably in the browser. ' +
      'Nothing was truncated. Please split the file into smaller parts or remove rows you do not need, then upload it again.';
  }

  function importCsvText(text, delimChoice) {
    importStats = { tokenMissing: 0, excelErrors: 0 };
    const notes = [];
    if (text.charCodeAt(0) === 0xFEFF) text = text.slice(1); // strip BOM
    if (!text.trim()) throw new Error('The file is empty or contains only blank lines.');
    if (text.indexOf('\u0000') !== -1) {
      throw new Error('This file does not look like a text CSV (it contains binary data). If it is an Excel workbook, save it with an .xlsx extension; if it is a CSV, re-export it as “CSV UTF-8”.');
    }
    const delim = delimChoice === 'auto' ? detectDelimiter(text) : DELIMITERS[delimChoice];
    const delimName = Object.keys(DELIMITERS).find((k) => DELIMITERS[k] === delim);
    const { rows, unterminated } = parseCsv(text, delim);
    if (unterminated) {
      throw new Error('The file appears to be malformed: a quoted value is missing its closing quotation mark. Open the file in a spreadsheet program, check for stray quotation marks, and save it again as CSV.');
    }
    const grid = rows.map((r) => r.map(normalizeCell));
    if (delimChoice === 'auto') notes.push({ kind: 'info', text: 'Column separator detected automatically: ' + delimName + '. You can change this under Import options.' });
    const maxCols = grid.reduce((m, r) => Math.max(m, r.length), 0);
    if (maxCols === 1 && grid.length > 1) {
      notes.push({ kind: 'warn', text: 'Only one column was found. If your file uses a different separator (for example semicolons), choose it under Import options.' });
    }
    const bad = (text.match(/\uFFFD/g) || []).length;
    if (bad > 0) {
      notes.push({ kind: 'warn', text: bad + ' character(s) could not be decoded as UTF-8 and were replaced. If your file was saved in another encoding, re-save it as “CSV UTF-8” in your spreadsheet program.' });
    }
    appendImportStatNotes(notes);
    return { grid, notes };
  }

  function appendImportStatNotes(notes) {
    if (importStats.tokenMissing > 0) {
      notes.push({ kind: 'info', text: fmtInt(importStats.tokenMissing) + ' cell(s) containing NA, N/A, NaN or null were treated as missing (same as blank cells).' });
    }
    if (importStats.excelErrors > 0) {
      notes.push({ kind: 'warn', text: fmtInt(importStats.excelErrors) + ' Excel error cell(s) (such as #DIV/0!) were treated as missing.' });
    }
  }

  /* ---------- Excel ---------- */
  function excelCellValue(cell) {
    if (!cell || cell.t === 'z') return null;
    if (cell.t === 'e') { importStats.excelErrors++; return null; }
    return normalizeCell(cell.v);
  }

  /** Convert one worksheet to a grid, using only the area that actually contains values. */
  function sheetToGrid(ws) {
    if (!ws) return [];
    let minR = Infinity, minC = Infinity, maxR = -1, maxC = -1;
    for (const key in ws) {
      if (key[0] === '!') continue;
      const cell = ws[key];
      if (!cell || cell.t === 'z' || cell.v === undefined || cell.v === null || cell.v === '') continue;
      const a = XLSX.utils.decode_cell(key);
      if (a.r < minR) minR = a.r;
      if (a.c < minC) minC = a.c;
      if (a.r > maxR) maxR = a.r;
      if (a.c > maxC) maxC = a.c;
    }
    if (maxR < 0) return [];
    if (maxR - minR + 1 > LIMITS.rows + 1) throw new Error(tooManyRowsMessage());
    if (maxC - minC + 1 > LIMITS.cols) {
      throw new Error('This worksheet has more than ' + LIMITS.cols + ' columns. Nothing was truncated. Please remove columns you do not need and upload the file again.');
    }
    const grid = [];
    for (let r = minR; r <= maxR; r++) {
      const row = new Array(maxC - minC + 1);
      for (let c = minC; c <= maxC; c++) row[c - minC] = excelCellValue(ws[XLSX.utils.encode_cell({ r, c })]);
      grid.push(row);
    }
    return grid;
  }

  function importSheet(workbook, sheetName) {
    importStats = { tokenMissing: 0, excelErrors: 0 };
    const grid = sheetToGrid(workbook.Sheets[sheetName]);
    if (!grid.length || !grid.some((r) => r.some((c) => c !== null))) {
      throw new Error('The worksheet “' + sheetName + '” is empty. Choose a different worksheet or upload another file.');
    }
    const notes = [];
    if (workbook.SheetNames.length > 1) notes.push({ kind: 'info', text: 'Analyzing worksheet “' + sheetName + '” (this workbook has ' + workbook.SheetNames.length + ' worksheets).' });
    appendImportStatNotes(notes);
    return { grid, notes };
  }

  /* ---------- Header detection ---------- */
  /** Guess whether the first non-empty row holds column names: every non-empty cell is non-numeric text, mostly unique. */
  function guessHasHeader(grid) {
    const first = grid.find((r) => r.some((c) => c !== null));
    if (!first) return true;
    const cells = first.filter((c) => c !== null);
    if (!cells.length) return true;
    const allText = cells.every((c) => typeof c === 'string' && Number.isNaN(parseNumber(c)));
    if (!allText) return false;
    const unique = new Set(cells.map((c) => c.toLowerCase())).size;
    if (unique / cells.length >= 0.8) return true;
    // Duplicate-looking names are still a header if some column below holds only numbers
    // (a row of repeated words above numeric data is almost certainly column names).
    const below = grid.slice(grid.indexOf(first) + 1, grid.indexOf(first) + 21);
    for (let j = 0; j < first.length; j++) {
      if (first[j] === null) continue;
      const col = below.map((r) => r[j]).filter((v) => v !== null && v !== undefined);
      if (col.length && col.every((v) => !Number.isNaN(parseNumber(v)))) return true;
    }
    return false;
  }

  /* ---------- Build the dataset ---------- */
  /**
   * Pure function: turns a grid into headers/rows/columns, or throws a friendly Error.
   * Nothing is discarded silently: every removal/rename is reported in `notes`.
   */
  function buildDataset(grid, opts) {
    const notes = [];
    let width = 0;
    grid.forEach((r) => { if (r.length > width) width = r.length; });
    let ragged = 0;
    const padded = grid.map((r) => {
      if (r.length === width) return r.slice();
      ragged++;
      return r.concat(new Array(width - r.length).fill(null));
    });
    if (ragged > 0) {
      notes.push({ kind: 'warn', text: fmtInt(ragged) + ' row(s) had fewer fields than the widest row; the missing fields were filled with missing values. Check that your file is formatted consistently.' });
    }

    const nonEmpty = padded.filter((r) => r.some((c) => c !== null));
    const emptyRows = padded.length - nonEmpty.length;
    if (emptyRows > 0) notes.push({ kind: 'info', text: fmtInt(emptyRows) + ' completely empty row(s) were ignored.' });
    if (!nonEmpty.length) throw new Error('The file contains no data.');

    let headerRow = null;
    let dataRows;
    if (opts.hasHeader) { headerRow = nonEmpty[0]; dataRows = nonEmpty.slice(1); } else { dataRows = nonEmpty; }
    if (!dataRows.length) {
      throw new Error(opts.hasHeader
        ? 'The file has a header row but no data rows below it.'
        : 'The file contains no data rows.');
    }
    if (dataRows.length > LIMITS.rows) throw new Error(tooManyRowsMessage());

    // User-defined missing-value codes (explicit, applies to all columns).
    if (opts.extraMissing.length) {
      const codeNums = opts.extraMissing.map((c) => ({ raw: c, num: parseNumber(c) }));
      let applied = 0;
      dataRows.forEach((r) => {
        for (let j = 0; j < r.length; j++) {
          const v = r[j];
          if (v === null) continue;
          const s = String(v).trim();
          const num = parseNumber(v);
          if (codeNums.some((c) => c.raw === s || (!Number.isNaN(c.num) && !Number.isNaN(num) && c.num === num))) {
            r[j] = null; applied++;
          }
        }
      });
      notes.push({ kind: 'info', text: fmtInt(applied) + ' cell(s) matching your missing-value codes (' + opts.extraMissing.join(', ') + ') were treated as missing.' });
    }

    // Drop columns that have no header AND no data (nothing is lost); everything else is kept.
    const keep = [];
    let droppedCols = 0;
    for (let j = 0; j < width; j++) {
      const h = headerRow ? headerRow[j] : null;
      const blankHeader = h === null || String(h).trim() === '';
      const allEmpty = dataRows.every((r) => r[j] === null);
      if (blankHeader && allEmpty) droppedCols++; else keep.push(j);
    }
    if (droppedCols > 0) notes.push({ kind: 'info', text: fmtInt(droppedCols) + ' empty column(s) without a name were ignored.' });
    if (!keep.length) throw new Error('The file contains no usable columns.');
    if (keep.length > LIMITS.cols) throw new Error('This file has more than ' + LIMITS.cols + ' columns. Nothing was truncated. Please remove columns you do not need and upload it again.');

    // Names: fill blanks, de-duplicate.
    const used = new Set();
    const names = [];
    const blankRenamed = [];
    const dupRenamed = [];
    keep.forEach((j) => {
      let base = headerRow && headerRow[j] !== null ? String(headerRow[j]).trim() : '';
      if (!base) { base = 'Column_' + (j + 1); if (headerRow) blankRenamed.push(base); }
      let name = base, k = 2;
      while (used.has(name)) name = base + '_' + (k++);
      if (name !== base) dupRenamed.push(base + ' \u2192 ' + name);
      used.add(name);
      names.push(name);
    });
    if (!opts.hasHeader) notes.push({ kind: 'info', text: 'No header row is being used, so columns are named Column_1, Column_2, … You can change this under Import options.' });
    if (blankRenamed.length) notes.push({ kind: 'warn', text: 'Blank column name(s) were replaced with generated names: ' + blankRenamed.join(', ') + '.' });
    if (dupRenamed.length) notes.push({ kind: 'warn', text: 'Duplicate column names were made unique: ' + dupRenamed.join('; ') + '.' });

    const rows = dataRows.map((r) => keep.map((j) => r[j]));
    const columns = keep.map((j, idx) => {
      const values = rows.map((r) => r[idx]);
      const profile = profileColumn(values);
      const det = detectType(names[idx], profile);
      return {
        index: idx, name: names[idx], values, profile,
        detectedType: det.type, detectedReason: det.reason,
        type: det.type, include: ANALYZABLE.has(det.type) && profile.n > 0
      };
    });

    // Duplicate (fully identical) rows: counted, flagged, never removed.
    const seen = new Set();
    let dupRows = 0;
    rows.forEach((r) => {
      const key = r.map((v) => (v === null ? '\u0002' : String(v))).join('\u0001');
      if (seen.has(key)) dupRows++; else seen.add(key);
    });

    return { headers: names, rows, columns, notes, dupRows };
  }

  /* ---------- Orchestration of import ---------- */
  function fileExtension(name) {
    const m = /\.([^.]+)$/.exec(name || '');
    return m ? m[1].toLowerCase() : '';
  }

  async function handleFile(file) {
    setMessages('uploadMessages', []);
    if (!file) return;
    const ext = fileExtension(file.name);
    if (!['csv', 'xlsx', 'xls'].includes(ext)) {
      addMessage('uploadMessages', 'error', '“' + file.name + '” is not a supported file type. Please upload a .csv, .xlsx or .xls file.' + (state.grid ? ' Your previous dataset is still loaded.' : ''));
      return;
    }
    if (file.size === 0) {
      addMessage('uploadMessages', 'error', 'The file is empty (0 bytes). Please choose a file that contains data.' + (state.grid ? ' Your previous dataset is still loaded.' : ''));
      return;
    }
    if (file.size > LIMITS.fileBytes) {
      addMessage('uploadMessages', 'error', 'The file is ' + (file.size / 1048576).toFixed(1) + ' MB, which exceeds the ' + (LIMITS.fileBytes / 1048576) + ' MB limit. Nothing was truncated. Remove unneeded columns or split the file, then upload again.');
      return;
    }
    setLoading(true, 'Reading file…');
    try {
      let result;
      let extra = {};
      if (ext === 'csv') {
        const text = await readFile(file, 'text');
        result = importCsvText(text, 'auto');
        extra = { kind: 'csv', csvText: text, workbook: null, sheetName: null, delimiterChoice: 'auto' };
      } else {
        if (typeof XLSX === 'undefined') {
          throw new Error('The Excel reader (SheetJS) did not load, so Excel files cannot be opened right now. Check your internet connection or follow the offline instructions above. CSV files still work.');
        }
        const buf = await readFile(file, 'buffer');
        let wb;
        try { wb = XLSX.read(buf, { type: 'array', cellDates: true }); }
        catch (e) { throw new Error('This file could not be read as an Excel workbook. It may be corrupted, password-protected, or not a real Excel file.'); }
        if (!wb.SheetNames || !wb.SheetNames.length) throw new Error('This workbook contains no worksheets.');
        // Prefer the first worksheet that has data.
        let chosen = null;
        for (const name of wb.SheetNames) {
          try { result = importSheet(wb, name); chosen = name; break; } catch (e) { if (/more than/.test(e.message)) throw e; }
        }
        if (!chosen) throw new Error('Every worksheet in this workbook is empty.');
        extra = { kind: 'excel', csvText: null, workbook: wb, sheetName: chosen, delimiterChoice: 'auto' };
      }
      const hasHeader = guessHasHeader(result.grid);
      const ds = buildDataset(result.grid, { hasHeader, extraMissing: [] });
      commitDataset(ds, result, { name: file.name, kind: extra.kind }, extra, hasHeader, []);
    } catch (err) {
      addMessage('uploadMessages', 'error', err.message + (state.grid ? ' Your previous dataset is still loaded.' : ''));
    } finally {
      setLoading(false);
    }
  }

  function loadSample() {
    setMessages('uploadMessages', []);
    importStats = { tokenMissing: 0, excelErrors: 0 };
    const grid = SAMPLE_GRID.map((r) => r.map(normalizeCell));
    const result = { grid, notes: [{ kind: 'info', text: 'This is a FICTIONAL sample dataset created for teaching. It does not contain real patient data.' }] };
    try {
      const ds = buildDataset(grid, { hasHeader: true, extraMissing: [] });
      commitDataset(ds, result, { name: 'Sample dataset (fictional educational data)', kind: 'sample' },
        { kind: 'sample', csvText: null, workbook: null, sheetName: null, delimiterChoice: 'auto' }, true, []);
    } catch (err) {
      addMessage('uploadMessages', 'error', err.message);
    }
  }

  /** Replace the current dataset with a new one and reset everything derived from the old one. */
  function commitDataset(ds, parseResult, source, extra, hasHeader, extraMissing) {
    resetResults();
    state.source = source;
    state.csvText = extra.csvText;
    state.workbook = extra.workbook;
    state.sheetName = extra.sheetName;
    state.delimiterChoice = extra.delimiterChoice;
    state.grid = parseResult.grid;
    state.parseNotes = parseResult.notes;
    state.hasHeader = hasHeader;
    state.extraMissing = extraMissing;
    applyBuilt(ds);
    $('missingCodesInput').value = extraMissing.join(', ');
    renderDataset();
  }

  function applyBuilt(ds) {
    state.headers = ds.headers;
    state.rows = ds.rows;
    state.columns = ds.columns;
    state.dupRows = ds.dupRows;
    state.buildNotes = ds.notes;
  }

  /** Re-build from the current grid after an option change (header toggle, missing codes). */
  function rebuildFromGrid(hasHeader, extraMissing) {
    const ds = buildDataset(state.grid, { hasHeader, extraMissing }); // may throw
    resetResults();
    state.hasHeader = hasHeader;
    state.extraMissing = extraMissing;
    applyBuilt(ds);
    renderDataset();
  }

  function removeDataset() {
    resetResults();
    state.source = null; state.csvText = null; state.workbook = null; state.sheetName = null;
    state.grid = null; state.parseNotes = []; state.headers = []; state.rows = []; state.columns = [];
    state.extraMissing = []; state.dupRows = 0; state.buildNotes = [];
    setMessages('uploadMessages', []);
    setMessages('importNotes', []);
    setMessages('reviewMessages', []);
    $('fileInput').value = '';
    $('missingCodesInput').value = '';
    $('datasetInfo').hidden = true;
    $('step2').hidden = true;
    updateStepper();
    $('dropzone').scrollIntoView({ behavior: prefersReducedMotion() ? 'auto' : 'smooth', block: 'center' });
  }

  /** Clears Table 1, charts and interpretation (e.g. when a new dataset replaces the old one). */
  function resetResults() {
    ['hist', 'bar', 'box'].forEach(destroyChart);
    state.confirmed = false;
    state.stale = false;
    state.analysis = null;
    state.table1 = null;
    state.chartText = { hist: '', bar: '', box: '' };
    state.chartNotes = { hist: '', bar: '', box: '' };
    clear($('table1'));
    clear($('interpretation'));
    setMessages('table1Warnings', []);
    $('table1Footnote').textContent = '';
    $('step3').hidden = true;
    $('step4').hidden = true;
    $('staleBanner').hidden = true;
    $('confirmBtn').textContent = 'Confirm classifications and generate Table 1';
    updateStepper();
  }

  /* ======================================================================
     4. VARIABLE PROFILING AND TYPE DETECTION
     ====================================================================== */
  const NUM_RE = /^[-+]?(?:\d+\.?\d*|\.\d+)(?:[eE][-+]?\d+)?$/;

  /** Strict numeric parse: returns NaN (never 0) for anything that is not a finite number. */
  function parseNumber(v) {
    if (typeof v === 'number') return Number.isFinite(v) ? v : NaN;
    if (typeof v !== 'string') return NaN;
    const s = v.trim();
    if (!NUM_RE.test(s)) return NaN;
    const x = Number(s);
    return Number.isFinite(x) ? x : NaN;
  }

  function looksLikeDate(s) {
    s = String(s).trim();
    let m = /^(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})(?:[T ].*)?$/.exec(s);
    if (m) return +m[2] >= 1 && +m[2] <= 12 && +m[3] >= 1 && +m[3] <= 31;
    m = /^(\d{1,2})[-/.](\d{1,2})[-/.](\d{2}|\d{4})$/.exec(s);
    if (m) return +m[1] >= 1 && +m[1] <= 31 && +m[2] >= 1 && +m[2] <= 31 && (+m[1] <= 12 || +m[2] <= 12);
    m = /^\d{1,2}[ -](Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*[ -,]*\d{2,4}$/i.exec(s);
    return !!m;
  }

  /** Summarise a column's values once; reused by type detection, review table and analysis. */
  function profileColumn(values) {
    const counts = new Map();   // label -> frequency of non-missing values
    const numericValues = [];   // finite numbers only, in row order
    const lowerLabels = new Map();
    let n = 0, missing = 0, nonNumeric = 0, dateLike = 0, intAll = true, leadingZero = 0;
    const numSet = new Set();
    for (const v of values) {
      if (v === null) { missing++; continue; }
      n++;
      const label = typeof v === 'number' ? String(v) : v;
      counts.set(label, (counts.get(label) || 0) + 1);
      const num = parseNumber(v);
      if (!Number.isNaN(num)) {
        numericValues.push(num);
        numSet.add(num);
        if (!Number.isInteger(num)) intAll = false;
        if (typeof v === 'string' && /^[-+]?0\d/.test(v)) leadingZero++;
      } else {
        nonNumeric++;
        if (looksLikeDate(label)) dateLike++;
        const low = label.toLowerCase();
        if (!lowerLabels.has(low)) lowerLabels.set(low, label);
      }
    }
    // Case-variant labels such as "Yes" vs "yes".
    let caseExample = null;
    if (lowerLabels.size < [...counts.keys()].filter((k) => Number.isNaN(parseNumber(k))).length) {
      const byLower = new Map();
      for (const k of counts.keys()) {
        if (!Number.isNaN(parseNumber(k))) continue;
        const low = k.toLowerCase();
        if (byLower.has(low)) { caseExample = [byLower.get(low), k]; break; }
        byLower.set(low, k);
      }
    }
    const codes = [];
    COMMON_MISSING_CODES.forEach((c) => {
      if (numSet.has(c)) {
        const cnt = values.filter((v) => v !== null && parseNumber(v) === c).length;
        codes.push({ code: c, count: cnt });
      }
    });
    const samples = [];
    for (const k of counts.keys()) { samples.push(truncate(k, 18)); if (samples.length >= 4) break; }
    return {
      n, missing, unique: counts.size, counts, numericValues, uniqueNumeric: numSet.size, nonNumeric,
      dateLike, intAll, leadingZero, caseExample, codes, samples
    };
  }

  const ID_TOKENS = new Set(['id', 'ids', 'mrn', 'uuid', 'guid', 'identifier', 'pid', 'uid']);
  const CATEGORY_HINTS = new Set(['sex', 'gender', 'race', 'ethnicity', 'group', 'arm', 'status', 'stage', 'grade', 'class', 'category', 'type', 'code', 'smoker', 'smoking', 'alcohol']);

  function nameTokens(name) {
    return name.replace(/([a-z0-9])([A-Z])/g, '$1 $2').toLowerCase().split(/[^a-z0-9]+/).filter(Boolean);
  }
  function nameLooksLikeId(name) {
    const tokens = nameTokens(name);
    if (tokens.some((t) => ID_TOKENS.has(t))) return true;
    return /^(patient|subject|participant|record|case|study|person|member)(no|num|number|code)$/.test(tokens.join(''));
  }
  function nameHintsCategory(name) { return nameTokens(name).some((t) => CATEGORY_HINTS.has(t)); }

  function isSequentialFromZeroOrOne(nums) {
    const s = nums.slice().sort((a, b) => a - b);
    if (s[0] !== 0 && s[0] !== 1) return false;
    for (let i = 1; i < s.length; i++) if (s[i] - s[i - 1] !== 1) return false;
    return true;
  }

  /**
   * Transparent type-detection rules (checked in this order). Detection is a suggestion only.
   *  1. No non-missing values            -> text / unclassified
   *  2. >= 90% of values look like dates -> date
   *  3. Name contains an ID word         -> identifier
   *  4. >= 90% numeric (and no leading-zero codes):
   *       2 distinct values              -> binary (e.g. 0/1, 1/2)
   *       1..n sequence (n >= 10)        -> identifier (row number)
   *       all whole numbers, <= 7 distinct values with >= 3 obs per value, or <= 12 with a
   *       category-style name (sex, group, stage...)  -> multicategory (numeric codes)
   *       otherwise                      -> continuous
   *  5. Otherwise (text):
   *       2 distinct values              -> binary
   *       >= 90% unique, short, no spaces, n >= 5 -> identifier
   *       <= 5 distinct, or <= 20 distinct and <= half the number of observations -> multicategory
   *       otherwise                      -> text / unclassified
   */
  function detectType(name, p) {
    if (p.n === 0) return { type: 'text', reason: 'All values are missing.' };
    if (p.dateLike / p.n >= 0.9) return { type: 'date', reason: 'Values look like dates.' };
    if (nameLooksLikeId(name)) return { type: 'identifier', reason: 'The column name looks like an identifier.' };
    const numShare = p.numericValues.length / p.n;
    if (numShare >= 0.9 && p.leadingZero === 0) {
      const u = p.uniqueNumeric;
      if (u === 1) return { type: 'multicategory', reason: 'Only one distinct numeric value.' };
      if (u === 2) return { type: 'binary', reason: 'Numeric, with exactly two distinct values (a code, not a measurement).' };
      if (p.intAll && p.n >= 10 && u === p.n && isSequentialFromZeroOrOne(p.numericValues)) {
        return { type: 'identifier', reason: 'Numbers run 1, 2, 3, … once each (looks like a row number).' };
      }
      if (p.intAll && ((u <= 7 && p.n / u >= 3) || (u <= 12 && nameHintsCategory(name)))) {
        return { type: 'multicategory', reason: 'Whole numbers with few distinct values (likely category codes).' };
      }
      return { type: 'continuous', reason: 'Numeric with many distinct values.' };
    }
    const u = p.unique;
    if (u === 2) return { type: 'binary', reason: 'Exactly two distinct values.' };
    if (u === 1) return { type: 'multicategory', reason: 'Only one distinct value.' };
    if (p.n >= 5 && u / p.n >= 0.9) {
      const labels = [...p.counts.keys()];
      const avgLen = labels.reduce((s, l) => s + l.length, 0) / labels.length;
      if (avgLen <= 20 && labels.every((l) => !/\s/.test(l))) return { type: 'identifier', reason: 'Nearly every value is unique and short (looks like an identifier).' };
      return { type: 'text', reason: 'Nearly every value is unique (free text).' };
    }
    if (u <= 5 || (u <= 20 && u <= p.n * 0.5)) return { type: 'multicategory', reason: 'A small set of repeated labels.' };
    return { type: 'text', reason: 'Many distinct text values.' };
  }

  /** Per-column notes shown in the review table (reflect the type currently selected). */
  function getColumnNotes(col) {
    const p = col.profile, notes = [];
    const analyzable = ANALYZABLE.has(col.type);
    if (p.n === 0) notes.push({ kind: 'warn', text: 'All values are missing; this variable cannot be summarized.' });
    if (analyzable && p.codes.length) {
      notes.push({ kind: 'warn', text: 'Contains ' + p.codes.map((c) => c.code + ' (' + c.count + '\u00d7)').join(', ') + ', which may be missing-value codes. They are NOT treated as missing unless you list them above.' });
    }
    if (col.type === 'continuous') {
      if (p.nonNumeric > 0) notes.push({ kind: 'warn', text: fmtInt(p.nonNumeric) + ' non-numeric ' + plural(p.nonNumeric, 'entry', 'entries') + ' will be excluded from calculations (not converted to zero).' });
      if (p.numericValues.length > 0 && p.numericValues.length < 5) notes.push({ kind: 'warn', text: 'Only ' + p.numericValues.length + ' numeric value(s); too few for a meaningful summary.' });
    }
    if (analyzable && (col.type !== 'continuous') && p.unique > MAX_BAR_CATEGORIES) {
      notes.push({ kind: 'warn', text: p.unique + ' distinct categories: Table 1 will be long and this variable cannot be charted (limit ' + MAX_BAR_CATEGORIES + '). Is it really categorical?' });
    }
    if (p.caseExample && (col.type === 'binary' || col.type === 'multicategory')) {
      notes.push({ kind: 'info', text: 'Labels such as “' + p.caseExample[0] + '” and “' + p.caseExample[1] + '” differ only by capitalization and are counted as separate categories. Consider cleaning the data.' });
    }
    const share = p.n ? p.numericValues.length / p.n : 0;
    if (p.n && share > 0.1 && share < 0.9) notes.push({ kind: 'info', text: 'Mix of numbers and text (' + Math.round(share * 100) + '% numeric).' });
    if (p.leadingZero > 0) notes.push({ kind: 'info', text: 'Values with leading zeros (e.g. 007) are better treated as text codes than as numbers.' });
    if (col.type === 'identifier') {
      let rep = 0;
      p.counts.forEach((c) => { if (c > 1) rep++; });
      if (rep > 0) notes.push({ kind: 'warn', text: rep + ' identifier value(s) appear more than once (repeated measurements, or duplicate records?). Nothing was removed.' });
    }
    if (col.type === 'date') notes.push({ kind: 'info', text: 'Dates are not summarized in Table 1.' });
    return notes;
  }

  /* ======================================================================
     5. STATISTICS
     ====================================================================== */

  /**
   * Percentile by linear interpolation between order statistics (R type 7 / Excel PERCENTILE.INC):
   *   h = (n - 1) * p ;  value = x[floor(h)] + (h - floor(h)) * (x[ceil(h)] - x[floor(h)])
   * `sorted` must be ascending.
   */
  function quantile(sorted, p) {
    const n = sorted.length;
    if (n === 0) return NaN;
    if (n === 1) return sorted[0];
    const h = (n - 1) * p;
    const lo = Math.floor(h), hi = Math.ceil(h);
    return sorted[lo] + (h - lo) * (sorted[hi] - sorted[lo]);
  }

  /** Descriptive statistics for an array of finite numbers. SD uses the sample (n - 1) denominator. */
  function describe(values) {
    const n = values.length;
    if (n === 0) return { n: 0 };
    const sorted = values.slice().sort((a, b) => a - b);
    let sum = 0;
    for (const v of sorted) sum += v;
    const mean = sum / n;
    let ss = 0; // two-pass algorithm: sum of squared deviations from the mean
    for (const v of sorted) { const d = v - mean; ss += d * d; }
    const sd = n > 1 ? Math.sqrt(ss / (n - 1)) : NaN;
    const q1 = quantile(sorted, 0.25), median = quantile(sorted, 0.5), q3 = quantile(sorted, 0.75);
    return { n, mean, sd, median, q1, q3, iqr: q3 - q1, min: sorted[0], max: sorted[n - 1], sorted };
  }

  /** Sample skewness (population-moment form, g1). */
  function skewness(values) {
    const n = values.length;
    if (n < 3) return NaN;
    const mean = values.reduce((s, v) => s + v, 0) / n;
    let m2 = 0, m3 = 0;
    for (const v of values) { const d = v - mean; m2 += d * d; m3 += d * d * d; }
    m2 /= n; m3 /= n;
    return m2 > 0 ? m3 / Math.pow(m2, 1.5) : NaN;
  }

  /**
   * Box-plot statistics (Tukey convention):
   *   box = Q1..Q3, line = median; fences = Q1 - 1.5*IQR and Q3 + 1.5*IQR;
   *   whiskers reach the most extreme observations INSIDE the fences;
   *   observations outside the fences are "potential outliers".
   */
  function boxStats(values) {
    const d = describe(values);
    if (!d.n) return null;
    const lowFence = d.q1 - 1.5 * d.iqr, highFence = d.q3 + 1.5 * d.iqr;
    let whiskerLow = Infinity, whiskerHigh = -Infinity;
    const outliers = [];
    for (const v of d.sorted) {
      if (v < lowFence || v > highFence) outliers.push(v);
      else { if (v < whiskerLow) whiskerLow = v; if (v > whiskerHigh) whiskerHigh = v; }
    }
    return { n: d.n, q1: d.q1, median: d.median, q3: d.q3, iqr: d.iqr, whiskerLow, whiskerHigh, outliers, min: d.min, max: d.max };
  }

  const POSITIVE_LABELS = new Set(['yes', 'y', 'true', 'present', 'positive']);
  const NEGATIVE_LABELS = new Set(['no', 'n', 'false', 'absent', 'negative']);
  /** Natural order, except Yes/No-style pairs list the positive label first. */
  function orderLabels(labels) {
    const arr = labels.slice();
    if (arr.length === 2) {
      const low = arr.map((l) => l.toLowerCase());
      const pos = low.findIndex((l) => POSITIVE_LABELS.has(l));
      const neg = low.findIndex((l) => NEGATIVE_LABELS.has(l));
      if (pos !== -1 && neg !== -1 && pos !== neg) return [arr[pos], arr[neg]];
    }
    return arr.sort((a, b) => a.localeCompare(b, undefined, { numeric: true, sensitivity: 'base' }));
  }

  function numericAt(col, i) { return parseNumber(col.values[i]); }

  /* ======================================================================
     6. QUALITY SUMMARY AND VARIABLE REVIEW
     ====================================================================== */
  function computeQuality() {
    const N = state.rows.length, P = state.columns.length;
    const totalCells = N * P;
    const missingCells = state.columns.reduce((s, c) => s + c.profile.missing, 0);
    const dupIds = [];
    state.columns.forEach((c) => {
      if (c.type !== 'identifier') return;
      let ids = 0, rowsInvolved = 0;
      c.profile.counts.forEach((cnt) => { if (cnt > 1) { ids++; rowsInvolved += cnt; } });
      if (ids > 0) dupIds.push({ name: c.name, ids, rows: rowsInvolved });
    });
    return { N, P, totalCells, missingCells, pctMissing: totalCells ? (missingCells / totalCells) * 100 : 0, dupRows: state.dupRows, dupIds };
  }

  function renderQuality() {
    const q = computeQuality();
    const box = $('qualityCards');
    clear(box);
    const card = (value, label) => el('div', { class: 'qcard' }, el('div', { class: 'qvalue', text: value }), el('div', { class: 'qlabel', text: label }));
    box.append(
      card(fmtInt(q.N), 'Observations (rows)'),
      card(fmtInt(q.P), 'Variables (columns)'),
      card(fmtInt(q.dupRows), 'Duplicate rows'),
      card(fmtInt(q.missingCells), 'Missing cells'),
      card(fmt(q.pctMissing, 1) + '%', 'Cells missing overall')
    );
    const msgs = [];
    if (q.dupRows > 0) msgs.push({ kind: 'warn', text: fmtInt(q.dupRows) + ' row(s) are exact copies of an earlier row. They have been kept; please check whether they are true duplicates before analyzing.' });
    q.dupIds.forEach((d) => msgs.push({ kind: 'warn', text: 'Identifier column “' + d.name + '”: ' + fmtInt(d.ids) + ' value(s) appear more than once (' + fmtInt(d.rows) + ' rows involved). This can be legitimate (repeated visits) or a sign of duplicate patients. Nothing was removed.' }));
    if (q.dupRows === 0 && q.dupIds.length === 0) msgs.push({ kind: 'info', text: 'No duplicate rows or repeated identifiers were detected.' });
    setMessages('qualityMessages', msgs);
  }

  function buildVariableRow(col) {
    const p = col.profile;
    const analyzable = ANALYZABLE.has(col.type);
    const tr = el('tr', { class: col.include ? '' : 'is-excluded', dataset: { col: col.index } });
    tr.append(el('th', { scope: 'row', text: col.name }));
    tr.append(el('td', null, el('span', { class: 'badge', text: TYPE_LABELS[col.detectedType] }), el('div', { class: 'reason', text: col.detectedReason })));

    const sel = el('select', { 'data-field': 'type', 'aria-label': 'Treat ' + col.name + ' as' });
    TYPE_ORDER.forEach((t) => {
      const opt = el('option', { value: t, text: TYPE_LABELS[t] });
      if (t === col.type) opt.selected = true;
      sel.append(opt);
    });
    tr.append(el('td', null, sel));

    const cb = el('input', { type: 'checkbox', 'data-field': 'include', 'aria-label': 'Include ' + col.name + ' in the analysis', disabled: !analyzable || p.n === 0 });
    cb.checked = col.include;
    tr.append(el('td', { class: 'center' }, cb));

    tr.append(el('td', { class: 'num', text: fmtInt(p.n) }));
    tr.append(el('td', { class: 'num', text: fmtInt(p.missing) + ' (' + fmt(state.rows.length ? (p.missing / state.rows.length) * 100 : 0, 1) + '%)' }));
    tr.append(el('td', { class: 'num', text: fmtInt(p.unique) }));
    tr.append(el('td', { class: 'samples', text: p.samples.length ? p.samples.join(', ') : '—' }));

    const notes = getColumnNotes(col);
    const td = el('td');
    if (notes.length) {
      const ul = el('ul', { class: 'notes-list' });
      notes.forEach((n) => ul.append(el('li', { class: 'note-' + n.kind }, el('strong', { text: n.kind === 'warn' ? 'Warning: ' : 'Note: ' }), n.text)));
      td.append(ul);
    } else td.append(document.createTextNode('—'));
    tr.append(td);
    return tr;
  }

  function renderVariableTable() {
    const tb = $('variableTbody');
    clear(tb);
    state.columns.forEach((c) => tb.append(buildVariableRow(c)));
  }

  /** Reject classifications that cannot work (e.g. "continuous" for a text column). Returns an error string or null. */
  function validateTypeChange(col, newType) {
    const p = col.profile;
    if (ANALYZABLE.has(newType) && p.n === 0) return '“' + col.name + '” has no recorded values, so it cannot be analyzed.';
    if (newType === 'continuous' && p.numericValues.length === 0) return '“' + col.name + '” contains no valid numbers, so it cannot be treated as continuous. Choose a categorical type instead.';
    if (newType === 'binary' && p.unique > 2) return '“' + col.name + '” has ' + p.unique + ' distinct values, so it cannot be binary. Choose “Multicategory categorical” instead.';
    return null;
  }

  function onVariableChange(e) {
    const target = e.target;
    const field = target.getAttribute('data-field');
    if (!field) return;
    const tr = target.closest('tr');
    const col = state.columns[Number(tr.dataset.col)];
    if (!col) return;
    setMessages('reviewMessages', []);
    if (field === 'type') {
      const newType = target.value;
      const problem = validateTypeChange(col, newType);
      if (problem) {
        target.value = col.type; // revert
        addMessage('reviewMessages', 'error', problem);
        return;
      }
      const wasAnalyzable = ANALYZABLE.has(col.type);
      col.type = newType;
      if (!(wasAnalyzable && ANALYZABLE.has(newType))) col.include = ANALYZABLE.has(newType) && col.profile.n > 0;
      if (newType === 'continuous' && col.profile.nonNumeric > 0) {
        addMessage('reviewMessages', 'warn', '“' + col.name + '” has ' + col.profile.nonNumeric + ' non-numeric entries; they will be excluded from its calculations.');
      }
    } else if (field === 'include') {
      col.include = target.checked;
    }
    const fresh = buildVariableRow(col);
    tr.replaceWith(fresh);
    const refocus = fresh.querySelector('[data-field="' + field + '"]');
    if (refocus) refocus.focus();
    renderQuality();
    markStale();
  }

  function markStale() {
    if (state.confirmed) {
      state.stale = true;
      $('staleBanner').hidden = false;
      $('confirmBtn').textContent = 'Update Table 1 and charts';
    }
  }

  /* ---------- Dataset rendering (info, preview, notes) ---------- */
  function renderDataset() {
    $('datasetInfo').hidden = false;
    $('fileNameOut').textContent = state.source.name;
    $('dimsOut').textContent = fmtInt(state.rows.length) + ' rows \u00d7 ' + fmtInt(state.columns.length) + ' columns imported';
    $('headerCheck').checked = state.hasHeader;

    const isExcel = state.source.kind === 'excel';
    const isCsv = state.source.kind === 'csv';
    $('sheetField').hidden = !(isExcel && state.workbook && state.workbook.SheetNames.length > 1);
    $('delimiterField').hidden = !isCsv;
    if (isExcel) {
      const sel = $('sheetSelect');
      clear(sel);
      state.workbook.SheetNames.forEach((n) => { const o = el('option', { value: n, text: n }); if (n === state.sheetName) o.selected = true; sel.append(o); });
    }
    if (isCsv) $('delimiterSelect').value = state.delimiterChoice;

    setMessages('importNotes', []);
    [...state.parseNotes, ...state.buildNotes].forEach((n) => addMessage('importNotes', n.kind, n.text));

    renderPreview();
    renderQuality();
    renderVariableTable();
    $('step2').hidden = false;
    updateStepper();
  }

  function renderPreview() {
    const table = $('previewTable');
    clear(table);
    const thead = el('thead');
    const hr = el('tr');
    state.headers.forEach((h) => hr.append(el('th', { scope: 'col', text: h })));
    thead.append(hr);
    const tbody = el('tbody');
    state.rows.slice(0, PREVIEW_ROWS).forEach((r) => {
      const tr = el('tr');
      r.forEach((v) => {
        if (v === null) tr.append(el('td', { class: 'cell-na', text: '(missing)' }));
        else tr.append(el('td', { class: 'cell-clip', title: String(v), text: String(v) }));
      });
      tbody.append(tr);
    });
    table.append(thead, tbody);
    $('previewMeta').textContent = '(showing the first ' + Math.min(PREVIEW_ROWS, state.rows.length) + ' of ' + fmtInt(state.rows.length) + ' rows)';
  }

  function updateStepper() {
    const hasData = !!state.grid;
    const done = [hasData, state.confirmed, state.confirmed, state.confirmed];
    let current = 1;
    if (hasData) current = state.confirmed ? 3 : 2;
    document.querySelectorAll('#stepper li').forEach((li) => {
      const s = Number(li.dataset.step);
      li.classList.toggle('is-done', done[s - 1] && s !== current);
      li.classList.toggle('is-current', s === current);
      if (s === current) li.setAttribute('aria-current', 'step'); else li.removeAttribute('aria-current');
    });
  }

  /* ======================================================================
     7. TABLE 1
     ====================================================================== */
  const ORDER_SELECT = (cols) => cols; // placeholder for clarity: analysis columns keep dataset order

  /** Compute statistics for every included variable (snapshot of the current classifications). */
  function computeAnalysis() {
    const N = state.rows.length;
    const list = [];
    state.columns.filter((c) => c.include && ANALYZABLE.has(c.type)).forEach((c) => {
      if (c.type === 'continuous') {
        const d = describe(c.profile.numericValues);
        list.push({ col: c, type: 'continuous', n: d.n, missing: N - d.n, nonNumeric: c.profile.nonNumeric, desc: d });
      } else {
        const total = c.profile.n;
        const cats = orderLabels([...c.profile.counts.keys()]).map((label) => ({ label, count: c.profile.counts.get(label) }));
        list.push({ col: c, type: c.type, n: total, missing: c.profile.missing, cats });
      }
    });
    return { N, vars: list };
  }

  /** Build the Table 1 model from the analysis and current display settings. */
  function buildTable1() {
    const A = state.analysis;
    const d = state.settings.decimals;
    const N = A.N;
    const rows = [];
    const warnings = [];
    const missText = (m) => fmtInt(m) + ' (' + fmtPct(m, N, d) + ')';
    A.vars.forEach((v) => {
      const name = displayName(v.col.name);
      if (v.type === 'continuous') {
        const s = v.desc;
        const meanSd = s.n ? fmt(s.mean, d) + (Number.isFinite(s.sd) ? ' (' + fmt(s.sd, d) + ')' : ' (\u2014)') : '\u2014';
        const medIqr = s.n ? fmt(s.median, d) + ' (' + fmtRange(s.q1, s.q3, d) + ')' : '\u2014';
        const base = { varName: v.col.name, n: v.n, missingText: missText(v.missing) };
        if (state.settings.summary === 'both') {
          rows.push({ ...base, label: name, indent: 0, value: '', kind: 'var-header' });
          rows.push({ label: 'Mean (SD)', indent: 1, value: meanSd, kind: 'sub-row', varName: v.col.name, n: null, missingText: '' });
          rows.push({ label: 'Median (IQR)', indent: 1, value: medIqr, kind: 'sub-row', varName: v.col.name, n: null, missingText: '' });
        } else {
          const isMean = state.settings.summary === 'mean';
          rows.push({ ...base, label: name + ', ' + (isMean ? 'mean (SD)' : 'median (IQR)'), indent: 0, value: isMean ? meanSd : medIqr, kind: 'var-header' });
        }
        if (s.n === 0) warnings.push({ kind: 'warn', text: displayName(v.col.name) + ' has no valid numeric values to summarize.' });
        else if (s.n < 5) warnings.push({ kind: 'warn', text: displayName(v.col.name) + ' has only ' + s.n + ' valid observation(s); summary statistics from so few values may not be meaningful.' });
        if (v.nonNumeric > 0) warnings.push({ kind: 'info', text: displayName(v.col.name) + ': ' + v.nonNumeric + ' non-numeric ' + plural(v.nonNumeric, 'entry was', 'entries were') + ' excluded from calculations and counted as missing.' });
      } else {
        rows.push({ label: name, indent: 0, value: '', kind: 'var-header', varName: v.col.name, n: v.n, missingText: missText(v.missing) });
        v.cats.forEach((c) => {
          rows.push({ label: c.label, indent: 1, value: fmtInt(c.count) + ' (' + fmtPct(c.count, v.n, d) + ')', kind: 'cat-row', varName: v.col.name, n: null, missingText: '' });
        });
        if (v.n === 0) warnings.push({ kind: 'warn', text: displayName(v.col.name) + ' has no recorded values.' });
        if (v.cats.length > MAX_BAR_CATEGORIES) warnings.push({ kind: 'warn', text: displayName(v.col.name) + ' has ' + v.cats.length + ' categories, which makes the table long. Check that it is really categorical.' });
      }
    });
    return { N, rows, warnings };
  }

  function renderTable1() {
    state.table1 = buildTable1();
    const T = state.table1;
    const table = $('table1');
    clear(table);
    const showCounts = state.settings.showCounts;
    const hr = el('tr');
    hr.append(el('th', { scope: 'col', text: 'Variable' }), el('th', { scope: 'col', text: 'Overall (N = ' + fmtInt(T.N) + ')' }));
    if (showCounts) hr.append(el('th', { scope: 'col', text: 'Observations (n)' }), el('th', { scope: 'col', text: 'Missing, n (%)' }));
    table.append(el('caption', { class: 'visually-hidden', text: 'Table 1. Descriptive statistics for included variables' }), el('thead', null, hr));
    const tb = el('tbody');
    T.rows.forEach((r) => {
      const tr = el('tr', { class: r.kind });
      tr.append(el('th', { scope: 'row', text: r.label }), el('td', { text: r.value }));
      if (showCounts) tr.append(el('td', { text: r.n === null ? '' : fmtInt(r.n) }), el('td', { text: r.missingText }));
      tb.append(tr);
    });
    table.append(tb);
    $('table1Footnote').textContent = FOOTNOTE + ' ' + FOOTNOTE_EXTRA;
    setMessages('table1Warnings', T.warnings);
  }

  /** Rows (array of arrays) shared by the CSV, Excel and HTML-report exports of Table 1. */
  function table1Matrix() {
    const T = state.table1;
    const showCounts = state.settings.showCounts;
    const head = ['Variable', 'Overall (N = ' + T.N + ')'];
    if (showCounts) head.push('Observations (n)', 'Missing, n (%)');
    const body = T.rows.map((r) => {
      const row = ['    '.repeat(r.indent) + r.label, r.value];
      if (showCounts) row.push(r.n === null ? '' : r.n, r.missingText);
      return row;
    });
    return { head, body };
  }

  /* ======================================================================
     8. CHARTS
     ====================================================================== */
  const whiteBackgroundPlugin = {
    id: 'whiteBackground',
    beforeDraw(chart) {
      const ctx = chart.ctx;
      ctx.save();
      ctx.globalCompositeOperation = 'destination-over';
      ctx.fillStyle = '#ffffff';
      ctx.fillRect(0, 0, chart.width, chart.height);
      ctx.restore();
    }
  };

  function chartsAvailable() { return typeof Chart !== 'undefined'; }

  function destroyChart(key) {
    if (state.charts[key]) { state.charts[key].destroy(); state.charts[key] = null; }
  }

  function showChartMessage(prefix, text) {
    $(prefix + 'Wrap').hidden = true;
    const e = $(prefix + 'Empty');
    e.hidden = false;
    e.textContent = text;
  }
  function showChartCanvas(prefix) {
    $(prefix + 'Wrap').hidden = false;
    $(prefix + 'Empty').hidden = true;
  }

  function fillSelect(sel, options, preferred) {
    const prev = preferred !== undefined ? preferred : sel.value;
    clear(sel);
    options.forEach((o) => sel.append(el('option', { value: o.value, text: o.label, disabled: !!o.disabled })));
    const enabled = options.filter((o) => !o.disabled);
    const chosen = enabled.find((o) => String(o.value) === String(prev)) || enabled[0];
    if (chosen) sel.value = String(chosen.value);
    sel.disabled = enabled.length === 0;
  }

  function populateChartControls() {
    const included = state.analysis.vars.map((v) => v.col);
    const cont = included.filter((c) => state.analysis.vars.find((v) => v.col === c).type === 'continuous');
    const cats = included.filter((c) => state.analysis.vars.find((v) => v.col === c).type !== 'continuous');
    const contOpts = cont.map((c) => ({ value: c.index, label: c.name }));
    fillSelect($('histVar'), contOpts);
    fillSelect($('boxVar'), contOpts);
    fillSelect($('barVar'), cats.map((c) => ({
      value: c.index,
      label: c.name + (c.profile.unique > MAX_BAR_CATEGORIES ? ' (too many categories: ' + c.profile.unique + ')' : ''),
      disabled: c.profile.unique > MAX_BAR_CATEGORIES
    })));
    fillSelect($('boxGroup'), [{ value: '', label: 'None (single box plot)' }].concat(cats.map((c) => ({
      value: c.index,
      label: c.name + (c.profile.unique > MAX_BOX_GROUPS ? ' (too many categories: ' + c.profile.unique + ')' : ''),
      disabled: c.profile.unique > MAX_BOX_GROUPS
    }))));
  }

  function renderAllCharts() {
    renderHistogram();
    renderBar();
    renderBox();
  }

  function baseChartOptions(titleText) {
    return {
      responsive: true,
      maintainAspectRatio: false,
      animation: false,
      plugins: {
        legend: { display: false },
        title: { display: true, text: titleText, color: COLORS.text, font: { size: 14, weight: '600' } }
      }
    };
  }

  /* ---------- Histogram ---------- */
  /** "Nice" bin width: 1, 2, 2.5, 5 or 10 times a power of ten, never smaller than the raw width. */
  function niceStep(raw) {
    const exp = Math.floor(Math.log10(raw));
    const base = Math.pow(10, exp);
    const f = raw / base;
    const nice = f <= 1 ? 1 : f <= 2 ? 2 : f <= 2.5 ? 2.5 : f <= 5 ? 5 : 10;
    return nice * base;
  }
  function decimalsNeeded(w) {
    for (let d = 0; d <= 6; d++) { const x = w * Math.pow(10, d); if (Math.abs(x - Math.round(x)) < 1e-9) return d; }
    return 6;
  }

  /**
   * Bins: target count from the Freedman-Diaconis rule (width = 2*IQR*n^(-1/3)), falling back to Sturges
   * when IQR = 0, then limited to 5-20 bins. Width is rounded up to a "nice" value; bins are
   * [lower, upper) except the last, which includes its upper edge.
   */
  function computeBins(values) {
    const d = describe(values);
    const n = d.n, min = d.min, max = d.max;
    if (min === max) {
      return { single: true, width: null, edges: [min, min], labels: [fmt(min, Math.max(0, decimalsNeeded(min)))], counts: [n], dec: 0 };
    }
    let k;
    if (d.iqr > 0) k = Math.ceil((max - min) / (2 * d.iqr * Math.cbrt(1 / n)));
    else k = Math.ceil(Math.log2(n) + 1);
    k = clamp(k, 5, 20);
    const width = niceStep((max - min) / k);
    const start = Math.floor(min / width + 1e-12) * width;
    const nb = Math.max(1, Math.ceil((max - start) / width - 1e-9));
    const dec = decimalsNeeded(width);
    const edge = (i) => Number((start + i * width).toPrecision(12));
    const counts = new Array(nb).fill(0);
    for (const v of values) counts[Math.min(nb - 1, Math.max(0, Math.floor((v - start) / width + 1e-9)))]++;
    const labels = [], ranges = [];
    for (let i = 0; i < nb; i++) {
      const a = edge(i), b = edge(i + 1);
      ranges.push([a, b]);
      labels.push(fmtRange(a, b, dec));
    }
    return { single: false, width, edges: ranges, labels, counts, dec };
  }

  function renderHistogram() {
    destroyChart('hist');
    state.chartText.hist = ''; state.chartNotes.hist = '';
    $('histNote').textContent = ''; $('histInterp').textContent = '';
    const sel = $('histVar');
    if (!sel.value) { showChartMessage('hist', 'No continuous variables are included in the analysis. Change a variable type in step 2 to use this chart.'); return; }
    if (!chartsAvailable()) { showChartMessage('hist', 'The chart library (Chart.js) did not load, so charts are unavailable. Check your internet connection or follow the offline instructions in step 1.'); return; }
    const col = state.columns[Number(sel.value)];
    const vals = col.profile.numericValues;
    const excluded = state.rows.length - vals.length;
    if (vals.length < 2) { showChartMessage('hist', '“' + col.name + '” has fewer than 2 valid numeric values, so a histogram cannot be drawn.'); return; }
    showChartCanvas('hist');
    const bins = computeBins(vals);
    const opts = baseChartOptions('Distribution of ' + col.name);
    opts.plugins.tooltip = {
      callbacks: {
        title: (items) => {
          const i = items[0].dataIndex;
          if (bins.single) return 'Value: ' + bins.labels[0];
          const [a, b] = bins.edges[i];
          return (i === bins.counts.length - 1 ? 'From ' + fmt(a, bins.dec) + ' to ' + fmt(b, bins.dec) + ' (including upper edge)' : 'From ' + fmt(a, bins.dec) + ' up to (not including) ' + fmt(b, bins.dec));
        },
        label: (item) => 'Observations: ' + item.parsed.y
      }
    };
    opts.scales = {
      x: { title: { display: true, text: col.name, color: COLORS.text }, grid: { display: false }, ticks: { maxRotation: 60, autoSkip: true } },
      y: { beginAtZero: true, title: { display: true, text: 'Frequency (number of observations)', color: COLORS.text }, ticks: { precision: 0 }, grid: { color: COLORS.grid } }
    };
    state.charts.hist = new Chart($('histCanvas'), {
      type: 'bar',
      data: { labels: bins.labels, datasets: [{ label: 'Frequency', data: bins.counts, backgroundColor: COLORS.tealFill, borderColor: COLORS.navy, borderWidth: 1, barPercentage: 1, categoryPercentage: 1 }] },
      options: opts,
      plugins: [whiteBackgroundPlugin]
    });
    $('histCanvas').setAttribute('aria-label', 'Histogram of ' + col.name + ' with ' + bins.counts.length + ' bins. Details are described in text below.');
    const note = vals.length + ' value(s) plotted; ' + excluded + ' excluded (missing or not a valid number).' +
      (bins.single ? ' All values are identical.' : ' Bin width: ' + fmt(bins.width, Math.max(bins.dec, 0)) + '. Each bar covers values from its lower edge up to, but not including, its upper edge (the last bar includes its upper edge).');
    state.chartNotes.hist = note;
    $('histNote').textContent = note;
    state.chartText.hist = describeHistogram(col, vals, bins);
    $('histInterp').textContent = state.chartText.hist;
  }

  function describeHistogram(col, vals, bins) {
    const d = state.settings.decimals;
    const s = describe(vals);
    const name = displayName(col.name);
    const parts = [name + ': ' + s.n + ' observations are shown; values range from ' + fmt(s.min, d) + ' to ' + fmt(s.max, d) + '.'];
    if (!bins.single) {
      const maxC = Math.max(...bins.counts);
      const top = bins.counts.reduce((acc, c, i) => (c === maxC ? acc.concat(i) : acc), []);
      if (top.length === 1) parts.push('The most common range is ' + bins.labels[top[0]] + ' (' + maxC + ' observations).');
      else parts.push('Several ranges share the highest count (' + maxC + '), so no single peak stands out.');
      if (s.n < 20) parts.push('With fewer than 20 observations the overall shape cannot be described reliably.');
      else {
        const g1 = skewness(vals);
        if (Number.isFinite(g1) && g1 >= 1) parts.push('The distribution looks right-skewed (a longer tail of higher values; skewness ' + fmt(g1, 2) + '), and the mean (' + fmt(s.mean, d) + ') is higher than the median (' + fmt(s.median, d) + ').');
        else if (Number.isFinite(g1) && g1 <= -1) parts.push('The distribution looks left-skewed (a longer tail of lower values; skewness ' + fmt(g1, 2) + '), and the mean (' + fmt(s.mean, d) + ') is lower than the median (' + fmt(s.median, d) + ').');
        else parts.push('No strong skewness is suggested by a simple measure (skewness ' + fmt(g1, 2) + ').');
        parts.push('The apparent shape of a histogram can change with the bin width.');
      }
    }
    return parts.join(' ');
  }

  /* ---------- Bar chart ---------- */
  /** Draws each bar's value next to the bar (counts or percentages) without external plugins. */
  function makeValueLabelPlugin(texts) {
    return {
      id: 'valueLabels',
      afterDatasetsDraw(chart) {
        const ctx = chart.ctx;
        const horizontal = chart.options.indexAxis === 'y';
        ctx.save();
        ctx.fillStyle = COLORS.text;
        ctx.font = '12px system-ui, sans-serif';
        chart.getDatasetMeta(0).data.forEach((bar, i) => {
          const p = bar.getProps(['x', 'y'], true);
          if (horizontal) { ctx.textAlign = 'left'; ctx.textBaseline = 'middle'; ctx.fillText(texts[i], p.x + 5, p.y); }
          else { ctx.textAlign = 'center'; ctx.textBaseline = 'bottom'; ctx.fillText(texts[i], p.x, p.y - 4); }
        });
        ctx.restore();
      }
    };
  }

  function renderBar() {
    destroyChart('bar');
    state.chartText.bar = ''; state.chartNotes.bar = '';
    $('barNote').textContent = ''; $('barInterp').textContent = '';
    const sel = $('barVar');
    if (!sel.value) { showChartMessage('bar', 'No suitable categorical variable is available. Variables with more than ' + MAX_BAR_CATEGORIES + ' categories cannot be charted legibly; include a binary or multicategory variable in step 2.'); return; }
    if (!chartsAvailable()) { showChartMessage('bar', 'The chart library (Chart.js) did not load, so charts are unavailable. Check your internet connection or follow the offline instructions in step 1.'); return; }
    const col = state.columns[Number(sel.value)];
    const total = col.profile.n;
    const missing = col.profile.missing;
    if (total === 0) { showChartMessage('bar', '“' + col.name + '” has no recorded values to plot.'); return; }
    showChartCanvas('bar');
    const pctMode = $('barModePct').checked;
    const cats = orderLabels([...col.profile.counts.keys()]).map((l) => ({ label: l, count: col.profile.counts.get(l) }));
    const horizontal = cats.length > 8;
    const values = cats.map((c) => (pctMode ? (c.count / total) * 100 : c.count));
    const texts = cats.map((c) => (pctMode ? fmt((c.count / total) * 100, 1) + '% (n=' + c.count + ')' : String(c.count)));
    const opts = baseChartOptions((pctMode ? 'Percentage of observations by ' : 'Number of observations by ') + col.name);
    opts.indexAxis = horizontal ? 'y' : 'x';
    opts.layout = { padding: horizontal ? { right: 80 } : { top: 20 } };
    opts.plugins.tooltip = {
      callbacks: {
        title: (items) => cats[items[0].dataIndex].label,
        label: (item) => { const c = cats[item.dataIndex]; return c.count + ' observations (' + fmt((c.count / total) * 100, 1) + '% of ' + total + ' with recorded values)'; }
      }
    };
    const catAxis = { title: { display: true, text: col.name, color: COLORS.text }, grid: { display: false }, ticks: { callback(v) { return truncate(this.getLabelForValue(v), 24); } } };
    const valAxis = {
      beginAtZero: true, grid: { color: COLORS.grid },
      title: { display: true, text: pctMode ? 'Percentage of observations with a recorded value (%)' : 'Number of observations', color: COLORS.text },
      ticks: pctMode ? {} : { precision: 0 }
    };
    if (pctMode) valAxis.max = 100;
    opts.scales = horizontal ? { y: catAxis, x: valAxis } : { x: catAxis, y: valAxis };
    state.charts.bar = new Chart($('barCanvas'), {
      type: 'bar',
      data: { labels: cats.map((c) => c.label), datasets: [{ label: pctMode ? 'Percent' : 'Count', data: values, backgroundColor: COLORS.tealFill, borderColor: COLORS.navy, borderWidth: 1 }] },
      options: opts,
      plugins: [whiteBackgroundPlugin, makeValueLabelPlugin(texts)]
    });
    $('barCanvas').setAttribute('aria-label', 'Bar chart of ' + col.name + ' with ' + cats.length + ' categories. Details are described in text below.');
    const note = total + ' observation(s) with a recorded value are shown; ' + missing + ' missing ' + plural(missing, 'observation is', 'observations are') + ' not plotted' +
      (pctMode ? ', and percentages use ' + total + ' as the denominator.' : '.');
    state.chartNotes.bar = note;
    $('barNote').textContent = note;
    const top = cats.reduce((a, c) => (c.count > a.count ? c : a), cats[0]);
    const topTies = cats.filter((c) => c.count === top.count).length > 1;
    state.chartText.bar = 'Among ' + total + ' observations with a recorded ' + displayName(col.name) + ', ' +
      (topTies ? 'several categories tie for the highest count (' + top.count + ').' : 'the most frequent category is \u201c' + top.label + '\u201d (' + top.count + ', ' + fmt((top.count / total) * 100, 1) + '%).') +
      (missing > 0 ? ' ' + missing + ' ' + plural(missing, 'observation has', 'observations have') + ' no recorded value and ' + plural(missing, 'is', 'are') + ' not shown.' : '');
    $('barInterp').textContent = state.chartText.bar;
  }

  /* ---------- Box plot (custom, drawn on a real Chart.js canvas) ---------- */
  /**
   * Chart.js has no built-in box plot, so we build a genuine one: a scatter chart provides the axes,
   * scales, tooltips and PNG export, and this plugin draws the box (Q1-Q3), median line, whiskers with
   * caps, and outlier circles at pixel positions computed from the chart's own scales.
   */
  function makeBoxPlugin(groups) {
    return {
      id: 'boxPlotDrawing',
      afterDatasetsDraw(chart) {
        const ctx = chart.ctx;
        const xs = chart.scales.x, ys = chart.scales.y;
        const catWidth = xs.width / groups.length;
        const w = Math.min(70, catWidth * 0.45);
        ctx.save();
        groups.forEach((g, i) => {
          const px = xs.getPixelForValue(i);
          const yQ1 = ys.getPixelForValue(g.q1), yQ3 = ys.getPixelForValue(g.q3), yMed = ys.getPixelForValue(g.median);
          const yLo = ys.getPixelForValue(g.whiskerLow), yHi = ys.getPixelForValue(g.whiskerHigh);
          ctx.strokeStyle = COLORS.navy;
          ctx.lineWidth = 1.5;
          // whiskers and caps
          ctx.beginPath();
          ctx.moveTo(px, yHi); ctx.lineTo(px, yQ3);
          ctx.moveTo(px, yQ1); ctx.lineTo(px, yLo);
          ctx.moveTo(px - w / 4, yHi); ctx.lineTo(px + w / 4, yHi);
          ctx.moveTo(px - w / 4, yLo); ctx.lineTo(px + w / 4, yLo);
          ctx.stroke();
          // box
          ctx.fillStyle = COLORS.tealLight;
          ctx.fillRect(px - w / 2, yQ3, w, yQ1 - yQ3);
          ctx.strokeRect(px - w / 2, yQ3, w, yQ1 - yQ3);
          // median
          ctx.lineWidth = 3.5;
          ctx.beginPath(); ctx.moveTo(px - w / 2, yMed); ctx.lineTo(px + w / 2, yMed); ctx.stroke();
          // potential outliers (hollow circles)
          ctx.lineWidth = 1.5;
          ctx.fillStyle = '#ffffff';
          g.outliers.forEach((o) => {
            ctx.beginPath(); ctx.arc(px, ys.getPixelForValue(o), 3.8, 0, Math.PI * 2); ctx.fill(); ctx.stroke();
          });
        });
        ctx.restore();
      }
    };
  }

  function renderBox() {
    destroyChart('box');
    state.chartText.box = ''; state.chartNotes.box = '';
    $('boxNote').textContent = ''; $('boxInterp').textContent = '';
    setMessages('boxWarnings', []);
    const sel = $('boxVar');
    if (!sel.value) { showChartMessage('box', 'No continuous variables are included in the analysis. Change a variable type in step 2 to use this chart.'); return; }
    if (!chartsAvailable()) { showChartMessage('box', 'The chart library (Chart.js) did not load, so charts are unavailable. Check your internet connection or follow the offline instructions in step 1.'); return; }
    const valCol = state.columns[Number(sel.value)];
    const grpIdx = $('boxGroup').value;
    const grpCol = grpIdx === '' ? null : state.columns[Number(grpIdx)];
    const warnings = [];

    // Collect values per group (row-wise so each observation stays paired with its group).
    let rawGroups;
    let groupMissing = 0, valueMissing = 0;
    if (!grpCol) {
      rawGroups = [{ label: 'All observations', vals: valCol.profile.numericValues.slice() }];
    } else {
      const map = new Map();
      for (let i = 0; i < state.rows.length; i++) {
        const gv = grpCol.values[i];
        if (gv === null) { groupMissing++; continue; }
        const x = numericAt(valCol, i);
        if (Number.isNaN(x)) { valueMissing++; continue; }
        const label = typeof gv === 'number' ? String(gv) : gv;
        if (!map.has(label)) map.set(label, []);
        map.get(label).push(x);
      }
      rawGroups = orderLabels([...map.keys()]).map((l) => ({ label: l, vals: map.get(l) }));
    }

    const drawable = [], skipped = [];
    rawGroups.forEach((g) => {
      if (g.vals.length < MIN_BOX_N) { skipped.push(g); return; }
      const st = boxStats(g.vals);
      drawable.push({ label: g.label, ...st });
      if (g.vals.length < 10) warnings.push({ kind: 'warn', text: '“' + g.label + '” has only ' + g.vals.length + ' observations; its box plot may not be representative.' });
      if (st.iqr === 0) warnings.push({ kind: 'warn', text: 'In “' + g.label + '” the middle 50% of values are identical (IQR = 0), so there is little variation to display and most other values may appear as potential outliers.' });
    });
    if (skipped.length) {
      warnings.push({ kind: 'warn', text: 'Not drawn because they have fewer than ' + MIN_BOX_N + ' valid observations: ' + skipped.map((g) => '“' + g.label + '” (n = ' + g.vals.length + ')').join(', ') + '.' });
    }
    if (grpCol && groupMissing > 0) warnings.push({ kind: 'info', text: groupMissing + ' observation(s) with no recorded ' + displayName(grpCol.name) + ' are not included in any group.' });
    if (grpCol && valueMissing > 0) warnings.push({ kind: 'info', text: valueMissing + ' observation(s) have no valid ' + displayName(valCol.name) + ' value and are excluded.' });
    if (!grpCol) {
      const ex = state.rows.length - valCol.profile.numericValues.length;
      if (ex > 0) warnings.push({ kind: 'info', text: ex + ' observation(s) with a missing or non-numeric ' + displayName(valCol.name) + ' value are excluded.' });
    }
    setMessages('boxWarnings', warnings);

    if (!drawable.length) {
      showChartMessage('box', 'There are too few observations for a meaningful box plot (each group needs at least ' + MIN_BOX_N + ' valid values).');
      return;
    }
    const allMin = Math.min(...drawable.map((g) => g.min)), allMax = Math.max(...drawable.map((g) => g.max));
    if (allMin === allMax) warnings.push({ kind: 'warn', text: 'All plotted values are identical, so the box plot is a flat line.' });
    setMessages('boxWarnings', warnings);
    showChartCanvas('box');
    const pad = allMax > allMin ? (allMax - allMin) * 0.08 : 1;
    const labels = drawable.map((g) => truncate(g.label, 24) + ' (n=' + g.n + ')');
    const d = state.settings.decimals;
    const opts = baseChartOptions('Box plot of ' + valCol.name + (grpCol ? ' by ' + grpCol.name : ''));
    opts.interaction = { mode: 'nearest', axis: 'x', intersect: false };
    opts.plugins.tooltip = {
      callbacks: {
        title: (items) => drawable[items[0].dataIndex].label,
        label: (item) => {
          const g = drawable[item.dataIndex];
          const lines = ['n = ' + g.n, 'Median: ' + fmt(g.median, d + 1), 'Q1\u2013Q3: ' + fmt(g.q1, d + 1) + ' to ' + fmt(g.q3, d + 1), 'Whiskers: ' + fmt(g.whiskerLow, d + 1) + ' to ' + fmt(g.whiskerHigh, d + 1)];
          lines.push('Potential outliers: ' + g.outliers.length + (g.outliers.length ? ' (' + g.outliers.slice(0, 5).map((o) => fmt(o, d + 1)).join(', ') + (g.outliers.length > 5 ? ', …' : '') + ')' : ''));
          return lines;
        }
      }
    };
    opts.scales = {
      x: { type: 'category', labels, offset: true, title: { display: true, text: grpCol ? grpCol.name : 'All observations', color: COLORS.text }, grid: { display: false } },
      y: { min: allMin - pad, max: allMax + pad, title: { display: true, text: valCol.name, color: COLORS.text }, grid: { color: COLORS.grid } }
    };
    state.charts.box = new Chart($('boxCanvas'), {
      type: 'scatter',
      // One invisible point per group (at the median) gives Chart.js something to attach tooltips to.
      data: { datasets: [{ label: 'Group summary', data: drawable.map((g, i) => ({ x: labels[i], y: g.median })), pointRadius: 0, pointHoverRadius: 0, pointHitRadius: 60, showLine: false }] },
      options: opts,
      plugins: [whiteBackgroundPlugin, makeBoxPlugin(drawable)]
    });
    $('boxCanvas').setAttribute('aria-label', 'Box plot of ' + valCol.name + (grpCol ? ' grouped by ' + grpCol.name : '') + ' with ' + drawable.length + ' box(es). Details are described in text below.');
    const note = 'Convention: box = Q1 to Q3 (linear-interpolation quartiles), line = median, whiskers reach the most extreme values within 1.5 \u00d7 IQR of the box, hollow circles = potential outliers beyond that.';
    state.chartNotes.box = note;
    $('boxNote').textContent = note;
    state.chartText.box = describeBox(valCol, grpCol, drawable);
    $('boxInterp').textContent = state.chartText.box;
  }

  function describeBox(valCol, grpCol, groups) {
    const d = state.settings.decimals;
    const name = displayName(valCol.name);
    const one = (g, prefix) => prefix + 'median ' + fmt(g.median, d) + ' (Q1\u2013Q3 ' + fmtRange(g.q1, g.q3, d) + ', n = ' + g.n + ')' +
      (g.outliers.length ? '; ' + g.outliers.length + ' ' + plural(g.outliers.length, 'value falls', 'values fall') + ' outside the whiskers' : '') + '.';
    let text;
    if (!grpCol) {
      text = 'For ' + name + ', the ' + one(groups[0], '');
      if (groups[0].outliers.length) text += ' Potential outliers are not necessarily errors and should be checked rather than removed automatically.';
    } else {
      text = 'Median ' + name + ' by ' + displayName(grpCol.name) + ': ' + groups.map((g) => g.label + ' ' + fmt(g.median, d) + ' (Q1\u2013Q3 ' + fmtRange(g.q1, g.q3, d) + ', n = ' + g.n + ')').join('; ') + '.';
      const withOut = groups.filter((g) => g.outliers.length);
      if (withOut.length) text += ' Potential outliers appear in ' + withOut.map((g) => g.label).join(', ') + '; they should be checked rather than removed automatically.';
      text += ' These are descriptive summaries only. No statistical test was performed, so they do not show whether groups differ significantly.';
    }
    return text;
  }

  /* ======================================================================
     9. PLAIN-ENGLISH INTERPRETATION (deterministic templates)
     ====================================================================== */
  function interpretationSentences(v) {
    const d = state.settings.decimals;
    const name = displayName(v.col.name);
    const N = state.analysis.N;
    const out = [];
    if (v.type === 'continuous') {
      const s = v.desc;
      if (s.n === 0) return ['There are no valid numeric values for ' + name + '.'];
      const mode = state.settings.summary;
      if (mode === 'mean' || mode === 'both') {
        out.push('The average ' + name + ' is ' + fmt(s.mean, d) + (Number.isFinite(s.sd) ? ' (standard deviation ' + fmt(s.sd, d) + ')' : '') + ', based on ' + s.n + ' ' + plural(s.n, 'observation') + ' with a recorded value.');
      }
      if (mode === 'median' || mode === 'both') {
        out.push('The median ' + name + ' is ' + fmt(s.median, d) + ', with an interquartile range of ' + fmtRange(s.q1, s.q3, d) + '.');
      }
      out.push('Recorded values range from ' + fmt(s.min, d) + ' to ' + fmt(s.max, d) + '.');
      if (s.n >= 10 && Number.isFinite(s.sd) && s.sd > 0 && Math.abs(s.mean - s.median) / s.sd >= 0.3) {
        out.push('The mean (' + fmt(s.mean, d) + ') and median (' + fmt(s.median, d) + ') differ noticeably, which can happen when values are skewed or include unusual values. Check the histogram before choosing how to report this variable.');
      }
    } else if (v.type === 'binary' && v.cats.length === 2) {
      const [a, b] = v.cats;
      if (v.n > 0) out.push('Among the ' + v.n + ' observations with a recorded ' + name + ', ' + a.count + ' (' + fmtPct(a.count, v.n, d) + ') are \u201c' + a.label + '\u201d and ' + b.count + ' (' + fmtPct(b.count, v.n, d) + ') are \u201c' + b.label + '\u201d.');
    } else if (v.n > 0) {
      const top = v.cats.reduce((acc, c) => (c.count > acc.count ? c : acc), v.cats[0]);
      out.push(name + ' has ' + v.cats.length + ' ' + plural(v.cats.length, 'category', 'categories') + ' among ' + v.n + ' observations with a recorded value. The most common is \u201c' + top.label + '\u201d (' + top.count + ', ' + fmtPct(top.count, v.n, d) + ').');
      if (v.cats.length <= 6) out.push('Percentages: ' + v.cats.map((c) => c.label + ' ' + fmtPct(c.count, v.n, d)).join('; ') + '.');
    }
    if (v.type !== 'continuous' && v.n === 0) out.push('No values are recorded for ' + name + '.');
    if (v.missing > 0) out.push(name + ' has ' + v.missing + ' missing ' + plural(v.missing, 'observation') + ' (' + fmtPct(v.missing, N, d) + ').');
    else out.push(name + ' has no missing observations.');
    return out;
  }

  /** Build DOM nodes for the "Understand your results" section (also reused by the HTML report). */
  function buildInterpretationNodes() {
    const nodes = [];
    if (!state.analysis) return nodes;
    const q = computeQuality();
    const d = state.settings.decimals;
    nodes.push(el('h4', { text: 'Dataset overview' }));
    const ov = el('ul');
    ov.append(
      el('li', { text: 'The dataset has ' + fmtInt(q.N) + ' observations and ' + fmtInt(q.P) + ' variables; ' + state.analysis.vars.length + ' ' + plural(state.analysis.vars.length, 'variable is', 'variables are') + ' summarized in Table 1.' }),
      el('li', { text: fmtInt(q.missingCells) + ' of ' + fmtInt(q.totalCells) + ' cells (' + fmt(q.pctMissing, d) + '%) are missing across all variables. Missing values are left out of each variable\u2019s calculations.' })
    );
    nodes.push(ov);
    nodes.push(el('h4', { text: 'Variable summaries' }));
    state.analysis.vars.forEach((v) => {
      nodes.push(el('h4', { text: displayName(v.col.name) }));
      const ul = el('ul');
      interpretationSentences(v).forEach((s) => ul.append(el('li', { text: s })));
      nodes.push(ul);
    });
    nodes.push(el('p', { class: 'interp-disclaimer', text: 'These statements describe only the uploaded dataset. Descriptive statistics cannot show causes, make a diagnosis, or establish clinical implications, and no statistical tests were performed.' }));
    return nodes;
  }

  function renderInterpretation() {
    const box = $('interpretation');
    clear(box);
    buildInterpretationNodes().forEach((n) => box.append(n));
  }

  /* ======================================================================
     Generate results (step 2 -> steps 3 and 4)
     ====================================================================== */
  function generateResults() {
    setMessages('reviewMessages', []);
    if (!state.columns.some((c) => c.include && ANALYZABLE.has(c.type))) {
      addMessage('reviewMessages', 'error', 'No variables are selected for analysis. Tick “Include in analysis” for at least one continuous or categorical variable.');
      return;
    }
    try {
      state.analysis = computeAnalysis();
      state.confirmed = true;
      state.stale = false;
      $('staleBanner').hidden = true;
      $('confirmBtn').textContent = 'Update Table 1 and charts';
      renderTable1();
      populateChartControls();
      $('step3').hidden = false;
      $('step4').hidden = false;
      renderAllCharts();
      renderInterpretation();
      updateStepper();
      $('step3').scrollIntoView({ behavior: prefersReducedMotion() ? 'auto' : 'smooth', block: 'start' });
    } catch (err) {
      addMessage('reviewMessages', 'error', 'Something went wrong while generating the results: ' + err.message);
    }
  }

  function onSettingsChange() {
    state.settings.summary = $('summarySelect').value;
    state.settings.decimals = Number($('decimalsSelect').value);
    state.settings.showCounts = $('showCountsCheck').checked;
    if (!state.analysis) return;
    renderTable1();
    renderInterpretation();
    // chart descriptions use the decimals setting too
    if (state.charts.hist) renderHistogram();
    if (state.charts.box) renderBox();
    if (state.charts.bar) renderBar();
  }

  /* ======================================================================
     10. EXPORTS
     ====================================================================== */
  /** Guards against spreadsheet formula injection when a CSV is opened in Excel. */
  function spreadsheetSafe(s) {
    if (/^[=+@\t\r]/.test(s)) return "'" + s;
    if (/^-/.test(s) && !/^-[\d.]+$/.test(s)) {
      // a leading minus before digits (e.g. "-3.2 (1.1)") is replaced by a true minus sign; other leading "-" gets an apostrophe
      return /^-\d/.test(s) ? '\u2212' + s.slice(1) : "'" + s;
    }
    return s;
  }
  function csvCell(v) {
    let s = v === null || v === undefined ? '' : String(v);
    s = spreadsheetSafe(s);
    return /[",\r\n]/.test(s) || /^\s|\s$/.test(s) ? '"' + s.replace(/"/g, '""') + '"' : s;
  }
  function toCsv(rows) { return '\uFEFF' + rows.map((r) => r.map(csvCell).join(',')).join('\r\n'); }

  function guardResults() {
    if (!state.table1) { notify('error', 'Generate Table 1 first (step 2).'); return false; }
    return true;
  }

  function exportTable1Csv() {
    if (!guardResults()) return;
    try {
      const { head, body } = table1Matrix();
      const rows = [head, ...body, [], ['Notes'], [FOOTNOTE], [FOOTNOTE_EXTRA]];
      downloadBlob(new Blob([toCsv(rows)], { type: 'text/csv;charset=utf-8' }), 'table1_' + safeBaseName() + '.csv');
    } catch (e) { notify('error', 'The CSV file could not be created: ' + e.message); }
  }

  function exportTable1Xlsx() {
    if (!guardResults()) return;
    if (typeof XLSX === 'undefined') { notify('error', 'The Excel library (SheetJS) did not load, so .xlsx export is unavailable. Use the CSV export, or check your connection / offline setup.'); return; }
    try {
      const { head, body } = table1Matrix();
      const aoa = [head, ...body, [], ['Notes'], [FOOTNOTE], [FOOTNOTE_EXTRA]];
      const ws = XLSX.utils.aoa_to_sheet(aoa);
      ws['!cols'] = [{ wch: 40 }, { wch: 22 }, { wch: 18 }, { wch: 18 }];
      const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, ws, 'Table 1');
      XLSX.writeFile(wb, 'table1_' + safeBaseName() + '.xlsx');
    } catch (e) { notify('error', 'The Excel file could not be created: ' + e.message); }
  }

  function qualityMatrix() {
    const q = computeQuality();
    const N = q.N;
    const rows = [
      ['Dataset-quality summary'],
      ['Metric', 'Value'],
      ['Observations (rows)', q.N],
      ['Variables (columns)', q.P],
      ['Duplicate rows (exact copies of an earlier row)', q.dupRows],
      ['Total cells', q.totalCells],
      ['Missing cells', q.missingCells],
      ['Missing cells (%)', fmt(q.pctMissing, 2)]
    ];
    q.dupIds.forEach((d) => rows.push(['Repeated identifier values in "' + d.name + '"', d.ids + ' values (' + d.rows + ' rows)']));
    rows.push([], ['Per-variable missing data'],
      ['Variable', 'Type used', 'Total rows', 'Non-missing', 'Missing', 'Missing (%)', 'Non-numeric entries (continuous only)', 'Unique values']);
    state.columns.forEach((c) => {
      rows.push([c.name, TYPE_LABELS[c.type], N, c.profile.n, c.profile.missing, fmt(N ? (c.profile.missing / N) * 100 : 0, 2), c.type === 'continuous' ? c.profile.nonNumeric : '', c.profile.unique]);
    });
    rows.push([], ['Note: missing values are excluded from calculations for the corresponding variable only; nothing was imputed or deleted.']);
    return rows;
  }

  function exportQualityCsv() {
    if (!state.grid) { notify('error', 'Upload a dataset first.'); return; }
    try {
      downloadBlob(new Blob([toCsv(qualityMatrix())], { type: 'text/csv;charset=utf-8' }), 'data_quality_' + safeBaseName() + '.csv');
    } catch (e) { notify('error', 'The CSV file could not be created: ' + e.message); }
  }

  function exportChartPng(key, filePrefix) {
    const chart = state.charts[key];
    if (!chart) { notify('error', 'There is no chart to download. Choose a variable with enough valid data first.'); return; }
    try {
      triggerDownload(chart.canvas.toDataURL('image/png'), filePrefix + '_' + safeBaseName() + '.png');
    } catch (e) { notify('error', 'The PNG could not be created: ' + e.message); }
  }

  const REPORT_CSS =
    'body{font-family:system-ui,-apple-system,"Segoe UI",Roboto,Arial,sans-serif;color:#1b2733;max-width:960px;margin:2rem auto;padding:0 1rem;line-height:1.5}' +
    'h1{color:#0b2545;margin-bottom:.2rem}h2{color:#0b2545;border-bottom:2px solid #0e7c86;padding-bottom:.25rem;margin-top:2rem}h3,h4{color:#13315c}' +
    'table{border-collapse:collapse;width:100%;font-variant-numeric:tabular-nums;margin:.5rem 0}' +
    'th,td{padding:.35rem .7rem;border-bottom:1px solid #dde3ea;text-align:right}th:first-child,td:first-child{text-align:left}' +
    'thead th{border-top:2px solid #0b2545;border-bottom:1px solid #0b2545}tbody tr:last-child td,tbody tr:last-child th{border-bottom:2px solid #0b2545}' +
    'tr.cat-row th,tr.sub-row th{padding-left:2rem;font-weight:400}tr.var-header th{font-weight:700}' +
    '.fn,.meta{color:#536171;font-size:.9rem}img{max-width:100%;height:auto;border:1px solid #dde3ea;border-radius:8px}figure{margin:1rem 0}';

  function buildReportHtml() {
    const doc = document.implementation.createHTMLDocument('Data Analysis Assistant \u2013 Results');
    const charset = doc.createElement('meta');
    charset.setAttribute('charset', 'utf-8');
    doc.head.insertBefore(charset, doc.head.firstChild);
    const style = doc.createElement('style');
    style.textContent = REPORT_CSS;
    doc.head.append(style);

    const body = doc.body;
    body.append(el('h1', { text: 'Data Analysis Assistant: results' }));
    body.append(el('p', { class: 'meta', text: 'Dataset: ' + state.source.name + ' \u2022 ' + fmtInt(state.rows.length) + ' rows, ' + fmtInt(state.columns.length) + ' columns \u2022 Generated ' + new Date().toLocaleString() + ' \u2022 Processed locally in the browser.' }));
    if (state.source.kind === 'sample') body.append(el('p', { class: 'meta', text: 'The sample dataset is fictional and for education only.' }));

    body.append(el('h2', { text: 'Table 1' }));
    const { head, body: rowsData } = table1Matrix();
    const table = el('table');
    table.append(el('thead', null, el('tr', null, ...head.map((h) => el('th', { scope: 'col', text: h })))));
    const tb = el('tbody');
    state.table1.rows.forEach((r, i) => {
      const tr = el('tr', { class: r.kind });
      tr.append(el('th', { scope: 'row', text: r.label }));
      rowsData[i].slice(1).forEach((c) => tr.append(el('td', { text: String(c) })));
      tb.append(tr);
    });
    table.append(tb);
    body.append(table, el('p', { class: 'fn', text: FOOTNOTE + ' ' + FOOTNOTE_EXTRA }));

    body.append(el('h2', { text: 'Dataset quality and missing data' }));
    const qm = qualityMatrix();
    const qTable = el('table');
    const sepIdx = qm.findIndex((r) => r[0] === 'Variable');
    qTable.append(el('thead', null, el('tr', null, ...qm[sepIdx].map((h) => el('th', { scope: 'col', text: h })))));
    const qb = el('tbody');
    qm.slice(sepIdx + 1).filter((r) => r.length > 1).forEach((r) => qb.append(el('tr', null, ...r.map((c, i) => (i === 0 ? el('th', { scope: 'row', text: String(c) }) : el('td', { text: String(c) }))))));
    qTable.append(qb);
    const qSum = computeQuality();
    body.append(el('p', { text: fmtInt(qSum.N) + ' observations; ' + fmtInt(qSum.P) + ' variables; ' + fmtInt(qSum.dupRows) + ' duplicate rows; ' + fmtInt(qSum.missingCells) + ' missing cells (' + fmt(qSum.pctMissing, 1) + '%).' }), qTable);

    body.append(el('h2', { text: 'Charts' }));
    [['hist', 'Histogram'], ['bar', 'Bar chart'], ['box', 'Box plot']].forEach(([key, title]) => {
      const chart = state.charts[key];
      if (!chart) return;
      const fig = el('figure');
      fig.append(el('h3', { text: title }));
      const img = doc.createElement('img');
      img.setAttribute('src', chart.canvas.toDataURL('image/png'));
      img.setAttribute('alt', title + ': ' + (state.chartText[key] || ''));
      fig.append(img);
      if (state.chartNotes[key]) fig.append(el('p', { class: 'fn', text: state.chartNotes[key] }));
      if (state.chartText[key]) fig.append(el('p', { text: state.chartText[key] }));
      body.append(fig);
    });

    body.append(el('h2', { text: 'Understand your results' }));
    buildInterpretationNodes().forEach((n) => body.append(n));
    return '<!DOCTYPE html>\n' + doc.documentElement.outerHTML;
  }

  function exportReportHtml() {
    if (!guardResults()) return;
    try {
      downloadBlob(new Blob([buildReportHtml()], { type: 'text/html;charset=utf-8' }), 'report_' + safeBaseName() + '.html');
    } catch (e) { notify('error', 'The report could not be created: ' + e.message); }
  }

  /* ======================================================================
     11. EVENT WIRING AND INITIALISATION (listeners are attached exactly once)
     ====================================================================== */
  function onSheetChange() {
    const name = $('sheetSelect').value;
    try {
      const result = importSheet(state.workbook, name);
      const hasHeader = guessHasHeader(result.grid);
      const ds = buildDataset(result.grid, { hasHeader, extraMissing: [] });
      state.sheetName = name;
      commitDataset(ds, result, state.source, { csvText: null, workbook: state.workbook, sheetName: name, delimiterChoice: 'auto', kind: 'excel' }, hasHeader, []);
      setMessages('uploadMessages', []);
    } catch (err) {
      $('sheetSelect').value = state.sheetName;
      setMessages('uploadMessages', [{ kind: 'error', text: err.message + ' The previous worksheet is still loaded.' }]);
    }
  }

  function onDelimiterChange() {
    const choice = $('delimiterSelect').value;
    try {
      const result = importCsvText(state.csvText, choice);
      const hasHeader = guessHasHeader(result.grid);
      const ds = buildDataset(result.grid, { hasHeader, extraMissing: [] });
      commitDataset(ds, result, state.source, { csvText: state.csvText, workbook: null, sheetName: null, delimiterChoice: choice, kind: 'csv' }, hasHeader, []);
      setMessages('uploadMessages', []);
    } catch (err) {
      $('delimiterSelect').value = state.delimiterChoice;
      setMessages('uploadMessages', [{ kind: 'error', text: err.message + ' The previous setting is still in use.' }]);
    }
  }

  function onHeaderToggle() {
    const wanted = $('headerCheck').checked;
    try {
      rebuildFromGrid(wanted, state.extraMissing);
      setMessages('uploadMessages', []);
    } catch (err) {
      $('headerCheck').checked = state.hasHeader;
      setMessages('uploadMessages', [{ kind: 'error', text: err.message }]);
    }
  }

  function onApplyMissingCodes() {
    const codes = $('missingCodesInput').value.split(/[,;\s]+/).map((s) => s.trim()).filter(Boolean);
    const unique = [...new Set(codes)];
    try {
      rebuildFromGrid(state.hasHeader, unique);
      addMessage('reviewMessages', 'info', unique.length
        ? 'Missing-value codes applied. Variable types were re-detected, so please review them again and re-generate Table 1.'
        : 'Missing-value codes cleared. Variable types were re-detected; please review them again.');
    } catch (err) {
      setMessages('reviewMessages', [{ kind: 'error', text: err.message }]);
    }
  }

  function wireDropzone() {
    const dz = $('dropzone');
    ['dragenter', 'dragover'].forEach((t) => dz.addEventListener(t, (e) => { e.preventDefault(); dz.classList.add('is-dragover'); }));
    ['dragleave', 'dragend'].forEach((t) => dz.addEventListener(t, () => dz.classList.remove('is-dragover')));
    dz.addEventListener('drop', (e) => {
      e.preventDefault();
      dz.classList.remove('is-dragover');
      const files = e.dataTransfer && e.dataTransfer.files;
      if (!files || !files.length) return;
      handleFile(files[0]).then(() => {
        if (files.length > 1) addMessage('uploadMessages', 'info', 'You dropped ' + files.length + ' files; only the first (' + files[0].name + ') was used.');
      });
    });
    // Stop the browser from opening a file that is dropped outside the drop zone.
    window.addEventListener('dragover', (e) => e.preventDefault());
    window.addEventListener('drop', (e) => e.preventDefault());
    $('fileInput').addEventListener('change', (e) => {
      const f = e.target.files && e.target.files[0];
      handleFile(f).finally(() => { e.target.value = ''; });
    });
  }

  function init() {
    if (!chartsAvailable()) {
      notify('warn', 'The chart library (Chart.js) could not be loaded, so charts will not work. Check your internet connection or use the offline instructions in step 1.');
    } else {
      Chart.defaults.font.family = 'system-ui, -apple-system, "Segoe UI", Roboto, Arial, sans-serif';
      Chart.defaults.color = COLORS.text;
    }
    if (typeof XLSX === 'undefined') {
      notify('warn', 'The Excel library (SheetJS) could not be loaded. CSV files still work, but Excel import and .xlsx export are unavailable until it loads.');
    }

    wireDropzone();
    $('sampleBtn').addEventListener('click', loadSample);
    $('removeBtn').addEventListener('click', removeDataset);
    $('sheetSelect').addEventListener('change', onSheetChange);
    $('delimiterSelect').addEventListener('change', onDelimiterChange);
    $('headerCheck').addEventListener('change', onHeaderToggle);
    $('applyMissingCodesBtn').addEventListener('click', onApplyMissingCodes);
    $('missingCodesInput').addEventListener('keydown', (e) => { if (e.key === 'Enter') { e.preventDefault(); onApplyMissingCodes(); } });
    $('variableTbody').addEventListener('change', onVariableChange); // delegated: survives re-rendering
    $('confirmBtn').addEventListener('click', generateResults);
    $('dlQualityCsv').addEventListener('click', exportQualityCsv);

    ['summarySelect', 'decimalsSelect', 'showCountsCheck'].forEach((id) => $(id).addEventListener('change', onSettingsChange));
    $('dlTableCsv').addEventListener('click', exportTable1Csv);
    $('dlTableXlsx').addEventListener('click', exportTable1Xlsx);

    $('histVar').addEventListener('change', renderHistogram);
    $('barVar').addEventListener('change', renderBar);
    $('barModeCount').addEventListener('change', renderBar);
    $('barModePct').addEventListener('change', renderBar);
    $('boxVar').addEventListener('change', renderBox);
    $('boxGroup').addEventListener('change', renderBox);
    $('dlHistPng').addEventListener('click', () => exportChartPng('hist', 'histogram'));
    $('dlBarPng').addEventListener('click', () => exportChartPng('bar', 'bar_chart'));
    $('dlBoxPng').addEventListener('click', () => exportChartPng('box', 'box_plot'));
    $('dlReportHtml').addEventListener('click', exportReportHtml);

    updateStepper();
  }

  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', init);
  else init();
})();
