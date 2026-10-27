'use strict';
/* Descriptive Statistics — runs entirely in the browser. No data leave the page. */

// ---------- Settings (easy to modify) ----------
const CONFIG = {
  SKEW_THRESHOLD: 1,        // |skewness| below this => approximately symmetric
  MIN_N_FOR_MEAN: 10,       // below this n, default to Median (IQR)
  NUMERIC_FRACTION: 0.8,    // share of non-missing values that must be numeric
  PREVIEW_ROWS: 8,
  HIGH_MISSING: 0.2,        // warn at >= 20% missing
  FREQ_LIMIT: 20,           // frequency-table levels shown before 'Show all'
  T1_MAX_LEVELS: 20,        // categorical levels listed in Table 1
  CHUNK: 10                 // columns analyzed between UI yields
};
const MISSING = new Set(['', 'na', 'n/a', 'nan', 'null', 'missing', '.', '-']);
const $ = id => document.getElementById(id);
const esc = s => String(s).replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
const displayVariableName = name => window.AppI18n?.text(name) || name;
let sheetJsPromise = null;
function loadSheetJS() {
  if (window.XLSX) return Promise.resolve(window.XLSX);
  if (!sheetJsPromise) sheetJsPromise = new Promise((resolve, reject) => {
    const script = document.createElement('script');
    script.src = 'https://cdn.sheetjs.com/xlsx-0.20.3/package/dist/xlsx.full.min.js';
    script.onload = () => window.XLSX ? resolve(window.XLSX) : reject(new Error('The Excel parser did not load.'));
    script.onerror = () => { sheetJsPromise = null; reject(new Error('Excel files need an internet connection to load the spreadsheet parser. Your file has not been uploaded.')); };
    document.head.appendChild(script);
  });
  return sheetJsPromise;
}
const state = { view: 'home', fig: null, patientId: '', types: {}, choices: {}, t1: { group: '', overall: true, testMode: 'parametric', showPValues: false, skip: new Set(), selected: new Set(), catChoices: {} }, expandedCategoricals: new Set(), excluded: [], cols: [], baseWarnings: [], fileName: '', headers: [], rows: [], results: [], categorical: [], sortKey: null, sortDir: 1 };
// About page state retained for future use: let aboutReturnState = null;

// ---------- CSV parsing (RFC 4180-style state machine) ----------
function parseCSV(text) {
  if (text.charCodeAt(0) === 0xFEFF) text = text.slice(1);
  const rows = [];
  let row = [], field = '', inQuotes = false;
  const endRow = () => {
    row.push(field); field = '';
    if (row.length > 1 || row[0].trim() !== '') rows.push(row); // skip blank rows
    row = [];
  };
  for (let i = 0; i < text.length; i++) {
    const c = text[i];
    if (inQuotes) {
      if (c === '"') { if (text[i + 1] === '"') { field += '"'; i++; } else inQuotes = false; }
      else field += c;
    } else if (c === '"' && field.trim() === '') { field = ''; inQuotes = true; }
    else if (c === ',') { row.push(field); field = ''; }
    else if (c === '\n') endRow();
    else if (c === '\r') { if (text[i + 1] === '\n') i++; endRow(); }
    else field += c;
  }
  if (field !== '' || row.length) endRow();
  return { rows, unterminatedQuote: inQuotes };
}

// ---------- Numeric handling ----------
function parseNumber(raw) {
  const s = raw.trim();
  if (/^[+-]?(\d+\.?\d*|\.\d+)([eE][+-]?\d+)?$/.test(s)) { const v = Number(s); return isFinite(v) ? v : null; }
  if (/^[+-]?\d{1,3}(,\d{3})+(\.\d+)?$/.test(s)) return Number(s.replace(/,/g, '')); // 1,234
  return null; // text, "Inf", etc. are never coerced
}

// Unique internal/display names: Age, Age (2), Age (3)
function uniqueNames(headers) {
  const used = new Set();
  return headers.map((h, i) => {
    const base = h.trim() || `Column ${i + 1}`;
    let name = base, k = 2;
    while (used.has(name)) name = `${base} (${k++})`;
    used.add(name);
    return name;
  });
}

function headerNameWarning(headers) {
  const raw = headers.map(h => String(h ?? '').trim());
  const normalized = uniqueNames(raw);
  return raw.some((name, i) => !name || name !== normalized[i])
    ? 'Some column names were blank or duplicated, so the app assigned unique names. Review them before interpreting results.'
    : '';
}

// ---------- Statistics ----------
// Quantile: linear interpolation between order statistics (R type 7 / Excel QUARTILE.INC)
function quantile(sorted, p) {
  const n = sorted.length;
  if (!n) return null;
  const h = (n - 1) * p, lo = Math.floor(h), hi = Math.ceil(h);
  return sorted[lo] + (h - lo) * (sorted[hi] - sorted[lo]);
}

// Skewness: adjusted Fisher-Pearson G1 = n/((n-1)(n-2)) * sum(((x-mean)/s)^3); needs n>=3, s>0
function computeStats(vals) {
  const sorted = Float64Array.from(vals).sort();
  const n = sorted.length;
  let sum = 0;
  for (let i = 0; i < n; i++) sum += sorted[i];
  const mean = sum / n;
  let ss = 0;
  for (let i = 0; i < n; i++) ss += (sorted[i] - mean) ** 2;
  const sd = n > 1 ? Math.sqrt(ss / (n - 1)) : null; // sample SD (n-1)
  let skew = null;
  if (n >= 3 && sd > 0) {
    let c = 0;
    for (let i = 0; i < n; i++) c += ((sorted[i] - mean) / sd) ** 3;
    skew = (n / ((n - 1) * (n - 2))) * c;
  }
  const q1 = quantile(sorted, 0.25), median = quantile(sorted, 0.5), q3 = quantile(sorted, 0.75);
  return { n, mean, sd, median, q1, q3, iqr: q3 - q1, min: sorted[0], max: sorted[n - 1], skew };
}

// Heuristic only — NOT a normality test
function chooseMethod(s) {
  if (s.sd === 0) return 'Mean (SD)';                  // constant variable
  if (s.n < CONFIG.MIN_N_FOR_MEAN) return 'Median (IQR)';
  return s.skew !== null && Math.abs(s.skew) < CONFIG.SKEW_THRESHOLD ? 'Mean (SD)' : 'Median (IQR)';
}

// ---------- Formatting (single central place; rounding only at display time) ----------
function decimalsFor(r) {
  const ref = [r.sd, r.iqr, Math.abs(r.mean)].find(v => v > 0) || 1;
  if (ref >= 100) return 0;
  if (ref >= 1) return 1;
  if (ref >= 0.1) return 2;
  return Math.min(6, Math.ceil(-Math.log10(ref)) + 1);
}
function fmt(x, d) {
  if (x === null || x === undefined || !isFinite(x)) return '—';
  if (Math.abs(x) >= 1e15) return x.toExponential(2);
  const value = Number(x.toFixed(d)) === 0 ? 0 : x;
  return new Intl.NumberFormat(window.AppI18n?.locale || 'en-US', { minimumFractionDigits: d, maximumFractionDigits: d, useGrouping: false }).format(value);
}
const localFixed = (x, d) => new Intl.NumberFormat(window.AppI18n?.locale || 'en-US', { minimumFractionDigits: d, maximumFractionDigits: d, useGrouping: false }).format(x);
function fmtSummary(s, method, d) {
  return method === 'Mean (SD)' ? `${fmt(s.mean, d)} (${fmt(s.sd, d)})` : `${fmt(s.median, d)} (${fmt(s.q1, d)}–${fmt(s.q3, d)})`;
}
function summaryText(r) { return fmtSummary(r, r.method, decimalsFor(r)); }
const missingText = r => `${r.missing.toLocaleString(window.AppI18n?.locale || 'en-US')} (${localFixed(r.missingPct, 1)}%)`;

// ---------- Column analysis ----------
function analyzeColumn(name, idx, rows, warnings, forced) {
  const total = rows.length, raw = new Array(total);
  let missing = 0;
  for (let i = 0; i < total; i++) {
    const s = (rows[i][idx] ?? '').trim();
    raw[i] = MISSING.has(s.toLowerCase()) ? null : s;
    if (raw[i] === null) missing++;
  }
  const rawNonMissing = total - missing;
  if (rawNonMissing === 0) { warnings.push(`“${name}” is empty and was skipped.`); return { type: 'excluded', name, empty: true, missing: total, missingPct: 100 }; }
  if (forced === 'exclude') return { type: 'excluded', name };
  const idLikeName = /(^|[^a-z])(id|identifier|mrn|medical\s*record|subject\s*id|participant\s*id|record\s*id|case\s*id)([^a-z]|$)/i.test(name);
  if (idLikeName) {
    const distinctRaw = new Set(raw.filter(value => value !== null)).size;
    if (distinctRaw / rawNonMissing >= 0.85) warnings.push(`“${name}” looks like a possible identifier because most non-missing values are unique; review its variable type or exclude it if appropriate.`);
  }
  const nums = raw.map(s => s === null ? null : parseNumber(s));   // null = missing or non-numeric
  const vals = nums.filter(v => v !== null);
  if (forced === 'numeric' && !vals.length) warnings.push(`“${name}” has no numeric values and was kept categorical.`);
  const asNum = forced === 'numeric' ? vals.length > 0 : forced === 'categorical' ? false : vals.length / rawNonMissing >= CONFIG.NUMERIC_FRACTION;
  if (asNum) {
    const nonNum = rawNonMissing - vals.length;
    missing += nonNum;
    const nonMissing = total - missing, missingPct = (missing / total) * 100;
    if (missingPct >= CONFIG.HIGH_MISSING * 100) warnings.push(`“${name}” has ${localFixed(missingPct, 0)}% missing values, including non-numeric entries.`);
    if (nonNum) warnings.push(`“${name}”: ${nonNum} non-numeric value(s) were counted as missing.`);
    const s = computeStats(vals), auto = chooseMethod(s);
    if (s.sd === 0) warnings.push(`“${name}” is constant.`);
    const distinct = new Set(vals).size;
    if (!forced && vals.length > 1 && distinct <= 5 && vals.every(Number.isInteger)) warnings.push(`“${name}” has only ${distinct} distinct integer values; it may be categorical.`);
    return { type: 'numeric', name, nums, missing, missingPct, autoMethod: auto, choice: 'auto', method: auto, ...s };
  }
  const nonMissing = rawNonMissing, missingPct = (missing / total) * 100;
  if (missingPct >= CONFIG.HIGH_MISSING * 100) warnings.push(`“${name}” has ${localFixed(missingPct, 0)}% missing values.`);
  const freq = new Map();   // exact (trimmed, case-sensitive) labels as strings
  raw.forEach(s => { if (s !== null) freq.set(s, (freq.get(s) || 0) + 1); });
  if (freq.size === nonMissing && nonMissing > 20) warnings.push(`“${name}” has all-distinct values (possible identifier).`);
  if (new Set([...freq.keys()].map(k => k.toLowerCase())).size < freq.size) warnings.push(`“${name}” has labels differing only by letter case; they are kept separate.`);
  return { type: 'categorical', name, labels: raw, freq, unique: freq.size, nonMissing, missing, missingPct };
}

// Process columns in chunks so the UI stays responsive on wide files
async function analyzeAll(headers, rows) {
  const warnings = [], results = [], categorical = [], excluded = [], cols = [];
  for (let c = 0; c < headers.length; c++) {
    try {
      const r = analyzeColumn(headers[c], c, rows, warnings, state.types[headers[c]]);
      cols.push(r); (r.type === 'numeric' ? results : r.type === 'categorical' ? categorical : excluded).push(r);
    } catch (e) { warnings.push(`“${headers[c]}” could not be analyzed and was skipped.`); }
    if (c % CONFIG.CHUNK === CONFIG.CHUNK - 1) {
      setStatus(`Analyzing variables… ${c + 1} of ${headers.length}`);
      await new Promise(r => setTimeout(r));
    }
  }
  return { results, categorical, excluded, cols, warnings };
}

// ---------- Loading ----------
function setStatus(msg) { $('status').textContent = window.AppI18n?.text(msg) || msg; }
function showError(msg) { const e = $('error'); e.textContent = window.AppI18n?.text(msg) || msg; e.hidden = !msg; }

function suggestPatientId(headers, rows) {
  const namePattern = /(^|[^a-z])(patient\s*id|patientid|mrn|medical\s*record|subject\s*id|participant\s*id|record\s*id|case\s*id)([^a-z]|$)/i;
  for (let i = 0; i < headers.length; i++) {
    if (!namePattern.test(headers[i])) continue;
    const values = rows.map(row => (row[i] ?? '').trim()).filter(value => value && !MISSING.has(value.toLowerCase()));
    if (values.length && new Set(values).size / values.length >= 0.85) return { name: headers[i], unique: new Set(values).size, count: values.length };
  }
  return null;
}

function askPatientId(headers, rows) {
  const dialog = $('patientPrompt'), select = $('patientColumn');
  select.innerHTML = '<option value="">Choose a column…</option>' + headers.map(h => `<option value="${esc(h)}">${esc(h)}</option>`).join('');
  const suggestion = suggestPatientId(headers, rows);
  $('patientSuggestion').hidden = !suggestion;
  $('patientSuggestion').textContent = suggestion ? `Possible patient ID: “${suggestion.name}” (${suggestion.unique} unique non-missing values). Select it if appropriate. Nothing is excluded until you choose the column and confirm.` : '';
  $('patientYes').disabled = true;
  dialog.showModal();
  return new Promise(resolve => dialog.addEventListener('close', () => resolve(dialog.returnValue === 'yes' ? select.value : ''), { once: true }));
}

function chooseWorksheet(sheetNames) {
  if (sheetNames.length === 1) return Promise.resolve(sheetNames[0]);
  const dialog = $('sheetPrompt'), select = $('sheetChoice');
  select.innerHTML = sheetNames.map(name => `<option value="${esc(name)}">${esc(name)}</option>`).join('');
  dialog.showModal();
  return new Promise(resolve => dialog.addEventListener('close', () => resolve(dialog.returnValue === 'choose' ? select.value : null), { once: true }));
}

async function loadText(text, name, options = {}) {
  showError(''); setStatus('Reading file…');
  await new Promise(r => setTimeout(r));
  try {
    const { rows: all, unterminatedQuote } = parseCSV(text);
    if (!all.length) throw new Error('The file is empty.');
    if (all.length === 1) throw new Error('The file contains only a header row and no data.');
    const headers = uniqueNames(all[0]);
    const rows = all.slice(1);
    const warnings = [];
    const headerWarning = headerNameWarning(all[0]);
    if (headerWarning) warnings.push(headerWarning);
    if (unterminatedQuote) warnings.push('The file ends inside an unclosed quote; the last row may be malformed.');
    const ragged = rows.filter(r => r.length !== headers.length).length;
    if (ragged) warnings.push(`${ragged} row(s) had a different number of cells than the header; missing cells were treated as empty and extras ignored.`);
    const dup = rows.length - new Set(rows.map(r => r.join('\u0001'))).size;
    if (dup) warnings.push(`${dup} duplicate row(s) found; none were removed.`);
    const patientId = options.skipPatientPrompt ? '' : await askPatientId(headers, rows);
    if (!options.skipPatientPrompt) setStatus('Choose whether to exclude a patient ID column…');
    state.patientId = patientId;
    state.types = {}; state.choices = {};
    if (patientId) {
      state.types[patientId] = 'exclude';
      // Patient ID exclusion warning box disabled by request:
      // warnings.push(`“${patientId}” was identified as a patient ID and excluded from analysis.`);
    }
    state.t1 = { group: '', overall: true, testMode: 'parametric', showPValues: false, skip: new Set(), selected: new Set(headers.filter(h => h !== patientId)), catChoices: {} };
    state.expandedCategoricals = new Set();
    const a = await analyzeAll(headers, rows);
    Object.assign(state, { baseWarnings: warnings, excluded: a.excluded, cols: a.cols, fileName: name, headers, rows, results: a.results, categorical: a.categorical, sortKey: null, sortDir: 1 });
    state.fig = null; state.view = 'home';
    renderAll([...warnings, ...a.warnings]);
    /* Patient ID warning banner disabled by request:
    if (patientId) {
      $('patientNotice').textContent = `Patient ID column “${patientId}” is excluded from summaries and figures. It is hidden in the data preview and is not included in Table 1.`;
      $('patientNotice').hidden = false;
    } else { $('patientNotice').hidden = true; $('patientNotice').textContent = ''; }
    */
    showView('home');
    setStatus('');
  } catch (e) {
    setStatus(''); showError(`Could not read this file: ${e.message}`);
  }
}

async function handleFile(file) {
  if (!file) return;
  if (/\.csv$/i.test(file.name) || /csv/i.test(file.type)) {
    try { await loadText(await file.text(), file.name); }
    catch (e) { showError('The file could not be read.'); }
    return;
  }
  if (/\.(xlsx|xls|xlsm|xlsb)$/i.test(file.name)) {
    setStatus('Opening Excel workbook…'); showError('');
    try {
      const XLSX = await loadSheetJS();
      const workbook = XLSX.read(await file.arrayBuffer());
      if (!workbook.SheetNames.length) throw new Error('The workbook does not contain a worksheet.');
      const sheetName = await chooseWorksheet(workbook.SheetNames);
      if (!sheetName) { setStatus(''); return; }
      const matrix = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], { header: 1, raw: false, defval: '', blankrows: false });
      const csv = matrix.map(row => row.map(value => {
        const cell = String(value ?? '');
        return /[",\r\n]/.test(cell) ? `"${cell.replace(/"/g, '""')}"` : cell;
      }).join(',')).join('\r\n');
      await loadText(csv, file.name);
    } catch (e) {
      setStatus(''); showError(`Could not read this Excel file: ${e.message}`);
    }
    return;
  }
  showError('Please choose a CSV file or an Excel workbook (.xlsx, .xls, .xlsm, or .xlsb).');
}

// ---------- Rendering ----------
function renderAll(warnings) {
  if ($('results').hidden) fadeHide($('landing'), () => { fadeShow($('results')); const h = $('home').querySelector('[data-h]'); if (h) h.focus(); });
  $('fname').textContent = state.fileName;
  $('fname').setAttribute('data-raw', '');
  $('fstats').innerHTML = [['Rows', state.rows.length], ['Columns', state.headers.length], ['Numeric', state.results.length], ['Categorical', state.categorical.length], ['Excluded', state.excluded.length]]
    .map(([k, v]) => {
      const count = v.toLocaleString(window.AppI18n?.locale || 'en-US');
      if (window.AppI18n?.language === 'fr') {
        const singular = v === 1;
        const labels = { Rows: singular ? 'ligne' : 'lignes', Columns: singular ? 'colonne' : 'colonnes', Numeric: `variable${singular ? '' : 's'} numérique${singular ? '' : 's'}`, Categorical: `variable${singular ? '' : 's'} catégorielle${singular ? '' : 's'}`, Excluded: `variable${singular ? '' : 's'} exclue${singular ? '' : 's'}` };
        return `<div class="count-first"><dt>${esc(labels[k])}</dt><dd>${count}</dd></div>`;
      }
      return `<div><dt>${k}</dt><dd>${count}</dd></div>`;
    }).join('');
  const w = $('warnings');
  w.hidden = !warnings.length;
  w.innerHTML = warnings.length ? `<strong>Notes</strong><ul>${warnings.slice(0, 30).map(x => `<li>${esc(x)}</li>`).join('')}${warnings.length > 30 ? `<li>…and ${warnings.length - 30} more.</li>` : ''}</ul>` : '';
  /* renderPreview(); renderTypes(); */
  renderTable(); renderCategorical(); fillGroupSelect(); renderT1(); figRefresh();
}

/* Data preview renderer retained for future use.
function renderPreview() {
  const rows = state.rows.slice(0, CONFIG.PREVIEW_ROWS);
  const cols = state.headers.map((h, i) => [h, i]).filter(([h]) => h !== state.patientId);
  $('preview').innerHTML = `<thead><tr>${cols.map(([h]) => `<th scope="col">${esc(h)}</th>`).join('')}</tr></thead><tbody>${
    rows.map(r => `<tr>${cols.map(([, i]) => `<td>${esc(r[i] ?? '')}</td>`).join('')}</tr>`).join('')}</tbody>`;
}
*/

function getView() {
  let v = state.cols.filter(r => r.type !== 'excluded');
  const k = state.sortKey;
  if (k) v = [...v].sort((a, b) => {
    const av = k === 'name' ? a.name : k === 'n' ? (a.type === 'numeric' ? a.n : a.nonMissing) : k === 'missingPct' ? a.missingPct : a.type === 'numeric' ? a.method : (state.t1.catChoices[a.name] || 'auto');
    const bv = k === 'name' ? b.name : k === 'n' ? (b.type === 'numeric' ? b.n : b.nonMissing) : k === 'missingPct' ? b.missingPct : b.type === 'numeric' ? b.method : (state.t1.catChoices[b.name] || 'auto');
    return (k === 'name' || k === 'method' ? String(av).localeCompare(String(bv), undefined, { numeric: true }) : av - bv) * state.sortDir;
  });
  return v;
}

// Apply a user choice ('auto' | 'Mean (SD)' | 'Median (IQR)') to one variable
function setMethod(r, choice) {
  r.choice = choice; state.choices[r.name] = choice;
  r.method = choice === 'auto' ? r.autoMethod : choice;
}
function methodSelect(r) {
  const opt = (v, label) => `<option value="${v}"${r.choice === v ? ' selected' : ''}>${label}</option>`;
  return `<select class="method" data-name="${esc(r.name)}" aria-label="Summary method for ${esc(r.name)}">
    ${opt('auto', `Auto: ${r.autoMethod}`)}${opt('Mean (SD)', 'Mean (SD)')}${opt('Median (IQR)', 'Median (IQR)')}</select>`;
}

function categoricalMode(r) { const choice = state.t1.catChoices[r.name] || 'auto'; return choice === 'auto' ? 'counts_percentage' : choice; }
function categoricalLevels(r) {
  const mode = categoricalMode(r), levels = [...r.freq].sort((a, b) => b[1] - a[1] || a[0].localeCompare(b[0], undefined, { numeric: true }));
  return levels.map(([label, n]) => {
    const pct = r.nonMissing ? localFixed(100 * n / r.nonMissing, 1) + '%' : '—';
    const value = mode === 'counts' ? String(n) : mode === 'percentage' ? pct : `${n} (${pct})`;
    return `${label}: ${value}`;
  });
}
function categoricalValue(r) { return categoricalLevels(r).join('; ') || '—'; }
function summaryValue(r) { return r.type === 'numeric' ? summaryText(r) : categoricalValue(r); }
function summaryCell(r) {
  if (r.type === 'numeric') return esc(summaryText(r));
  const levels = categoricalLevels(r);
  if (!levels.length) return '—';
  const expanded = state.expandedCategoricals.has(r.name), shown = expanded ? levels : levels.slice(0, 8);
  const toggle = levels.length > 8 ? `<button type="button" class="btn cat-expand" data-expand-cat="${esc(r.name)}" aria-expanded="${expanded}">${expanded ? 'Show fewer' : `Show all (${levels.length})`}</button>` : '';
  return `${shown.map(level => `<span data-raw>${esc(level)}</span>`).join('<br>')}${toggle}`;
}
function categoryMethodSelect(r) {
  const choice = state.t1.catChoices[r.name] || 'auto';
  const options = [['auto', 'Auto: Counts (percentage)'], ['counts', 'Counts'], ['percentage', 'Percentage'], ['counts_percentage', 'Counts (percentage)']];
  return `<select class="method" data-cat-method="${esc(r.name)}" aria-label="Summary format for ${esc(r.name)}">${options.map(([v, label]) => `<option value="${v}"${choice === v ? ' selected' : ''}>${label}</option>`).join('')}</select>`;
}

function renderTable() {
  const view = getView();
  $('stats').querySelector('tbody').innerHTML = view.length ? view.map(r => `<tr${r.type === 'numeric' ? ' tabindex="0"' : ''} data-name="${esc(r.name)}" data-type="${r.type}"${r.type === 'numeric' ? ` aria-label="Details for ${esc(r.name)}"` : ''}>
    <td>${esc(r.name)}</td><td class="num">${(r.type === 'numeric' ? r.n : r.nonMissing).toLocaleString(window.AppI18n?.locale || 'en-US')}</td><td class="num">${missingText(r)}</td><td class="summary-cell">${summaryCell(r)}</td><td>${r.type === 'numeric' ? methodSelect(r) : categoryMethodSelect(r)}</td></tr>`).join('')
    : `<tr><td colspan="5">No summary variables are available.</td></tr>`;
  document.querySelectorAll('#stats th[aria-sort]').forEach(th => {
    const b = th.querySelector('button');
    th.setAttribute('aria-sort', b && b.dataset.sort === state.sortKey ? (state.sortDir === 1 ? 'ascending' : 'descending') : 'none');
  });
  updateScope(view.length);
}

// Make the scope of global settings and exports explicit when a filter is active
function updateScope(shown) {
  const n = state.results.length, c = state.categorical.length;
  $('scope').textContent = `Continuous settings apply to all ${n} numeric variable${n === 1 ? '' : 's'}; categorical settings apply to all ${c} categorical variable${c === 1 ? '' : 's'}.`;
  $('exportScope').textContent = '';
}



function showDetail(name) {
  const r = state.results.find(x => x.name === name);
  if (!r) return;

  const d = Math.max(decimalsFor(r), 1);
  const f = x => fmt(x, d);

  const rowsData = [
    ['N', r.n.toLocaleString(window.AppI18n?.locale || 'en-US')],
    ['Missing', missingText(r)],
    ['Mean', f(r.mean)],
    ['SD', f(r.sd)],
    ['Median', f(r.median)],
    ['Q1', f(r.q1)],
    ['Q3', f(r.q3)],
    ['IQR', f(r.iqr)],
    ['Minimum', f(r.min)],
    ['Maximum', f(r.max)],
    ['Skewness', fmt(r.skew, 2)],
    ['Selected summary', r.method + (r.choice === 'auto' ? ' (auto)' : ' (manual)')],
    ['Auto-suggested', r.autoMethod]
  ];

  $('dtitle').textContent = r.name;
  $('dtitle').dataset.variable = r.name;
  $('dtitle').setAttribute('data-raw', '');

  $('dbody').innerHTML = rowsData
    .map(([k, v]) => `<dt>${k}</dt><dd>${esc(v)}</dd>`)
    .join('');

  $('detail').showModal();

  // Draw after dialog becomes visible
  requestAnimationFrame(() => drawBoxPlot(r));
}

// ---------- Box & Whisker Plot ----------

function drawBoxPlot(r) {
  const canvas = $('boxplot');
  if (!canvas) return;

  const ctx = canvas.getContext('2d');

  // Account for high-DPI displays
  const cssWidth = 700;
  const cssHeight = 280;
  const dpr = window.devicePixelRatio || 1;

  canvas.width = cssWidth * dpr;
  canvas.height = cssHeight * dpr;
  canvas.style.width = '100%';
  canvas.style.height = 'auto';

  ctx.setTransform(dpr, 0, 0, dpr, 0, 0);
  ctx.clearRect(0, 0, cssWidth, cssHeight);

  const left = 55;
  const right = 25;
  const top = 35;
  const bottom = 55;
  const plotWidth = cssWidth - left - right;
  const y = 125;
  const boxHeight = 70;

  const values = [
    r.min,
    r.q1,
    r.median,
    r.q3,
    r.max
  ];

  const min = r.min;
  const max = r.max;

  // Handle constant variables
  let range = max - min;
  if (range === 0) range = 1;

  const x = value =>
    left + ((value - min) / range) * plotWidth;

  // Background
  ctx.fillStyle = '#ffffff';
  ctx.fillRect(0, 0, cssWidth, cssHeight);

  // Title
  ctx.fillStyle = '#1d2430';
  ctx.font = '600 15px system-ui, -apple-system, "Segoe UI", Roboto, sans-serif';
  ctx.textAlign = 'center';
  ctx.fillText(window.AppI18n?.text('Box plot (whiskers span min to max; no outlier rule)') || 'Box plot (whiskers span min to max; no outlier rule)', cssWidth / 2, 20);

  // Axis
  ctx.strokeStyle = '#aab3bf';
  ctx.lineWidth = 1;

  ctx.beginPath();
  ctx.moveTo(left, y + boxHeight / 2);
  ctx.lineTo(left + plotWidth, y + boxHeight / 2);
  ctx.stroke();

  // Whisker line
  const xMin = x(r.min);
  const xQ1 = x(r.q1);
  const xMed = x(r.median);
  const xQ3 = x(r.q3);
  const xMax = x(r.max);

  ctx.strokeStyle = '#1f4e79';
  ctx.lineWidth = 2;

  // Main whisker
  ctx.beginPath();
  ctx.moveTo(xMin, y + boxHeight / 2);
  ctx.lineTo(xQ1, y + boxHeight / 2);
  ctx.moveTo(xQ3, y + boxHeight / 2);
  ctx.lineTo(xMax, y + boxHeight / 2);
  ctx.stroke();

  // Minimum whisker cap
  ctx.beginPath();
  ctx.moveTo(xMin, y + 15);
  ctx.lineTo(xMin, y + boxHeight - 15);
  ctx.stroke();

  // Maximum whisker cap
  ctx.beginPath();
  ctx.moveTo(xMax, y + 15);
  ctx.lineTo(xMax, y + boxHeight - 15);
  ctx.stroke();

  // IQR box
  ctx.fillStyle = '#eef3f8';
  ctx.strokeStyle = '#1f4e79';
  ctx.lineWidth = 2;

  ctx.fillRect(
    xQ1,
    y,
    Math.max(xQ3 - xQ1, 1),
    boxHeight
  );

  ctx.strokeRect(
    xQ1,
    y,
    Math.max(xQ3 - xQ1, 1),
    boxHeight
  );

  // Median
  ctx.strokeStyle = '#1d2430';
  ctx.lineWidth = 3;

  ctx.beginPath();
  ctx.moveTo(xMed, y);
  ctx.lineTo(xMed, y + boxHeight);
  ctx.stroke();

  // Determine sensible axis tick values
  const tickCount = 5;
  const d = decimalsFor(r);

  ctx.font = '12px system-ui, -apple-system, "Segoe UI", Roboto, sans-serif';
  ctx.fillStyle = '#5b6573';
  ctx.textAlign = 'center';

  for (let i = 0; i <= tickCount; i++) {
    const value = min + (max - min) * (i / tickCount);
    const px = x(value);

    // Tick
    ctx.strokeStyle = '#d9dee5';
    ctx.lineWidth = 1;

    ctx.beginPath();
    ctx.moveTo(px, y + boxHeight + 5);
    ctx.lineTo(px, y + boxHeight + 11);
    ctx.stroke();

    // Label
    ctx.fillStyle = '#5b6573';
    ctx.fillText(fmt(value, d), px, y + boxHeight + 29);
  }

  // Labels beneath the plot
  ctx.font = '600 11px system-ui, -apple-system, "Segoe UI", Roboto, sans-serif';

  const labelY = y + boxHeight + 48;

  ctx.fillStyle = '#5b6573';

  ctx.fillText(window.AppI18n?.text('Min') || 'Min', xMin, labelY);
  ctx.fillText(window.AppI18n?.text('Q1') || 'Q1', xQ1, labelY);
  ctx.fillText(window.AppI18n?.text('Median') || 'Median', xMed, labelY);
  ctx.fillText(window.AppI18n?.text('Q3') || 'Q3', xQ3, labelY);
  ctx.fillText(window.AppI18n?.text('Max') || 'Max', xMax, labelY);
}

// ---------- Export ----------
const tableRows = () => [['Variable', 'n', 'Missing data, n (%)', 'Value', 'Statistic'].map(x => window.AppI18n?.text(x) || x),
  ...getView().map(r => [safe(displayVariableName(r.name)), r.type === 'numeric' ? r.n : r.nonMissing, missingText(r), safe(summaryValue(r)), window.AppI18n?.text(r.type === 'numeric' ? (r.choice === 'auto' ? `Auto: ${r.autoMethod}` : r.method) : (({ auto: 'Auto: Counts (percentage)', counts: 'Counts', percentage: 'Percentage', counts_percentage: 'Counts (percentage)' })[state.t1.catChoices[r.name] || 'auto'])) || (r.type === 'numeric' ? r.method : state.t1.catChoices[r.name] || 'auto')])];
const csvCell = v => /[",\n\r]/.test(String(v)) ? `"${String(v).replace(/"/g, '""')}"` : String(v);

async function copyText(text, okMsg) {
  if (!text) { setStatus('Nothing to copy.'); return; }
  let ok = false;
  try {
    if (navigator.clipboard && window.isSecureContext) { await navigator.clipboard.writeText(text); ok = true; }
  } catch (_) { ok = false; }
  if (!ok) { // fallback for insecure contexts or denied permission
    const t = document.createElement('textarea');
    t.value = text; t.setAttribute('readonly', ''); t.style.cssText = 'position:fixed;opacity:0';
    document.body.appendChild(t); t.select();
    try { ok = document.execCommand('copy'); } catch (_) { ok = false; }
    t.remove();
  }
  setStatus(ok ? okMsg : 'Copy failed. Use Export → Download CSV instead.');
}
function download(text, filename, type) {
  const a = document.createElement('a');
  a.href = URL.createObjectURL(new Blob([text], { type }));
  a.download = filename; document.body.appendChild(a); a.click(); a.remove();
  setTimeout(() => URL.revokeObjectURL(a.href), 1000);
}
const baseName = () => (state.fileName.replace(/\.(csv|xlsx|xls|xlsm|xlsb)$/i, '') || 'data') + '_descriptive_stats';

// ---------- Sample data (entirely synthetic) ----------
function sampleCSV() {
  let seed = 20240607;
  const rand = () => { seed |= 0; seed = seed + 0x6D2B79F5 | 0; let t = Math.imul(seed ^ seed >>> 15, 1 | seed); t = t + Math.imul(t ^ t >>> 7, 61 | t) ^ t; return ((t ^ t >>> 14) >>> 0) / 4294967296; };
  const norm = (m, s) => m + s * Math.sqrt(-2 * Math.log(1 - rand())) * Math.cos(2 * Math.PI * rand());
  const maybe = (v, p) => rand() < p ? (rand() < 0.5 ? '' : 'NA') : v;
  const lines = ['Age,BMI,Hemoglobin,CRP,Albumin,Length_of_stay,Disease_duration,Sex,Diagnosis,Disease_activity,Smoking_status,Diabetes,Hypertension,Treatment'];
  for (let i = 0; i < 240; i++) {
    lines.push([
      Math.max(18, Math.round(norm(54, 12))),
      maybe(Math.max(14, norm(27, 4.5)).toFixed(1), 0.03),
      maybe(norm(13.1, 1.6).toFixed(1), 0.02),
      maybe(Math.exp(norm(1.2, 1.1)).toFixed(1), 0.08),
      maybe(Math.min(5, norm(3.8, 0.45)).toFixed(1), 0.05),
      Math.max(1, Math.round(Math.exp(norm(1.5, 0.8)))),
      maybe(Math.exp(norm(1.6, 0.9)).toFixed(1), 0.04),
      rand() < 0.48 ? 'Male' : 'Female',
      rand() < 0.55 ? "Crohn's disease" : 'Ulcerative colitis',
      (() => { const p = rand(); return p < 0.2 ? 'Remission' : p < 0.45 ? 'Mild' : p < 0.8 ? 'Moderate' : 'Severe'; })(),
      (() => { const p = rand(); return p < 0.62 ? 'Never' : p < 0.82 ? 'Former' : 'Current'; })(),
      rand() < 0.14 ? 'Yes' : 'No',
      rand() < 0.28 ? 'Yes' : 'No',
      (() => { const p = rand(); return p < 0.42 ? 'Standard therapy' : p < 0.82 ? 'Biologic' : 'Combination therapy'; })()
    ].join(','));
  }
  return lines.join('\n');
}


// ---------- Variable types, categorical frequencies, Table 1 ----------
const safe = s => /^\s*[=+\-@]/.test(String(s)) ? `'${s}` : String(s);   // export-only guard against spreadsheet formula injection
const tsvCell = v => String(v).replace(/[\t\r\n]+/g, ' ');

/* Variable type override workflow retained for future use.
async function reanalyze() {
  setStatus('Recalculating…');
  const a = await analyzeAll(state.headers, state.rows);
  Object.assign(state, { results: a.results, categorical: a.categorical, excluded: a.excluded, cols: a.cols });
  a.results.forEach(r => setMethod(r, state.choices[r.name] || 'auto'));
  renderAll([...state.baseWarnings, ...a.warnings]);
  setStatus('Variable types updated.');
}

function renderTypes() {
  const f = document.activeElement && document.activeElement.dataset ? document.activeElement.dataset.tname : null;
  const opts = c => [['auto', 'Auto'], ['numeric', 'Numeric'], ['categorical', 'Categorical'], ['exclude', 'Exclude']]
    .map(([v, l]) => `<option value="${v}"${(state.types[c.name] || 'auto') === v ? ' selected' : ''}>${l}</option>`).join('');
  $('types').innerHTML = `<thead><tr><th scope="col">Variable</th><th scope="col">Current</th><th scope="col">Type</th><th scope="col">In Table 1</th></tr></thead><tbody>${
    state.cols.map(c => `<tr><td data-raw>${esc(c.name)}</td><td>${c.empty ? 'Empty (skipped)' : c.type === 'excluded' ? 'Excluded' : c.type}</td><td><select class="method" data-tname="${esc(c.name)}" aria-label="Type of ${esc(c.name)}">${opts(c)}</select></td><td><input type="checkbox" data-t1="${esc(c.name)}" aria-label="Include ${esc(c.name)} in Table 1"${state.t1.skip.has(c.name) ? '' : ' checked'}${c.type === 'excluded' ? ' disabled' : ''}></td></tr>`).join('')}</tbody>`;
  if (f) { const el = [...$('types').querySelectorAll('select')].find(e => e.dataset.tname === f); if (el) el.focus(); }
}
*/

function freqTable(v) {
  const rows = [...v.freq].sort((a, b) => b[1] - a[1] || a[0].localeCompare(b[0], undefined, { numeric: true }));
  const wrap = document.createElement('div'), box = document.createElement('div'), tbl = document.createElement('table'), note = document.createElement('p'), btn = document.createElement('button');
  box.className = 'tablewrap'; note.className = 'note'; btn.type = 'button'; btn.className = 'btn'; box.append(tbl); wrap.append(box, note, btn);
  const draw = limit => {
    tbl.innerHTML = `<thead><tr><th scope="col">Category</th><th scope="col" class="num">n (%)</th></tr></thead><tbody>${rows.slice(0, limit).map(([k, n]) =>
      `<tr><td data-raw>${esc(k)}</td><td class="num">${n.toLocaleString(window.AppI18n?.locale || 'en-US')} (${localFixed(100 * n / v.nonMissing, 1)}%)</td></tr>`).join('')}</tbody>`;
    const hidden = Math.max(0, rows.length - limit);
    note.textContent = `Denominator: ${v.nonMissing.toLocaleString(window.AppI18n?.locale || 'en-US')} non-missing; ${v.missing.toLocaleString(window.AppI18n?.locale || 'en-US')} missing excluded.${hidden ? ` Showing ${limit} of ${rows.length} levels.` : ''}`;
    btn.hidden = !hidden; btn.textContent = `Show all ${rows.length} levels`;
  };
  btn.addEventListener('click', () => draw(Infinity)); draw(CONFIG.FREQ_LIMIT);
  return wrap;
}

function renderCategorical() {
  if (!$('catsearch') || !$('catinfo') || !$('cat')) return;
  const c = state.categorical, q = $('catsearch').value.trim().toLowerCase(), box = $('cat');
  $('catinfo').textContent = `${c.length} categorical, ${state.results.length} numeric, ${state.excluded.length} excluded. Labels are trimmed and case-sensitive; missing values are shown separately.`;
  box.replaceChildren();
  c.filter(v => !q || v.name.toLowerCase().includes(q)).forEach(v => {
    const d = document.createElement('details'), s = document.createElement('summary');
    s.textContent = `${v.name}: ${v.unique.toLocaleString(window.AppI18n?.locale || 'en-US')} level${v.unique === 1 ? '' : 's'}, ${v.missing.toLocaleString(window.AppI18n?.locale || 'en-US')} missing (${localFixed(v.missingPct, 1)}%)`;
    d.append(s);
    d.addEventListener('toggle', () => { if (d.open && !d.dataset.done) { d.dataset.done = '1'; d.append(freqTable(v)); } });
    box.append(d);
  });
  if (!box.children.length) box.textContent = c.length ? 'No variables match your search.' : 'No categorical variables.';
}

function fillGroupSelect() {
  if (!state.categorical.some(c => c.name === state.t1.group)) state.t1.group = '';
  $('t1group').innerHTML = '<option value="">None</option>' + state.categorical.map(c => `<option data-raw value="${esc(c.name)}"${c.name === state.t1.group ? ' selected' : ''}>${esc(c.name)}</option>`).join('');
  const selectable = state.headers.filter(name => name !== state.patientId);
  $('t1variables').innerHTML = selectable.map(name => `<label class="t1-variable"><input type="checkbox" data-t1-variable="${esc(name)}"${state.t1.selected.has(name) ? ' checked' : ''}><span data-raw>${esc(name)}</span></label>`).join('');
  $('t1variablesLabel').textContent = `Select variables (${state.t1.selected.size}/${selectable.length})`;
  renderCategoricalSettings();
}

function renderCategoricalSettings() {
  const options = [['auto', 'Auto: Counts (percentage)'], ['counts', 'Counts'], ['percentage', 'Percentage'], ['counts_percentage', 'Counts (percentage)']];
  $('catSettings').innerHTML = state.categorical.map(r => '<label class="cat-setting"><span data-raw>' + esc(r.name) + '</span><select class="method" data-cat-method="' + esc(r.name) + '" aria-label="Summary format for ' + esc(r.name) + '">' + options.map(([v, label]) => '<option value="' + v + '"' + ((state.t1.catChoices[r.name] || 'auto') === v ? ' selected' : '') + '>' + label + '</option>').join('') + '</select></label>').join('') || '<p class="note">No categorical variables were detected.</p>';
}

// ---------- Group comparison tests (kept independent of Table 1 rendering) ----------
function logGamma(z) {
  const p = [676.5203681218851, -1259.1392167224028, 771.32342877765313, -176.61502916214059, 12.507343278686905, -0.13857109526572012, 9.9843695780195716e-6, 1.5056327351493116e-7];
  if (z < 0.5) return Math.log(Math.PI) - Math.log(Math.sin(Math.PI * z)) - logGamma(1 - z);
  z -= 1;
  let x = 0.99999999999980993;
  for (let i = 0; i < p.length; i++) x += p[i] / (z + i + 1);
  const t = z + p.length - 0.5;
  return 0.9189385332046727 + (z + 0.5) * Math.log(t) - t + Math.log(x);
}
function betaFraction(a, b, x) {
  const max = 240, eps = 3e-14, tiny = 1e-300;
  const qab = a + b, qap = a + 1, qam = a - 1;
  let c = 1, d = 1 - qab * x / qap;
  if (Math.abs(d) < tiny) d = tiny;
  d = 1 / d;
  let h = d;
  for (let m = 1; m <= max; m++) {
    const m2 = 2 * m;
    let aa = m * (b - m) * x / ((qam + m2) * (a + m2));
    d = 1 + aa * d; if (Math.abs(d) < tiny) d = tiny;
    c = 1 + aa / c; if (Math.abs(c) < tiny) c = tiny;
    d = 1 / d; h *= d * c;
    aa = -(a + m) * (qab + m) * x / ((a + m2) * (qap + m2));
    d = 1 + aa * d; if (Math.abs(d) < tiny) d = tiny;
    c = 1 + aa / c; if (Math.abs(c) < tiny) c = tiny;
    d = 1 / d;
    const delta = d * c; h *= delta;
    if (Math.abs(delta - 1) < eps) break;
  }
  return h;
}
function regularizedBeta(x, a, b) {
  if (x <= 0) return 0; if (x >= 1) return 1;
  const bt = Math.exp(logGamma(a + b) - logGamma(a) - logGamma(b) + a * Math.log(x) + b * Math.log1p(-x));
  return x < (a + 1) / (a + b + 2) ? bt * betaFraction(a, b, x) / a : 1 - bt * betaFraction(b, a, 1 - x) / b;
}
function regularizedGammaQ(a, x) {
  if (x <= 0) return 1;
  const gln = logGamma(a), eps = 3e-14, tiny = 1e-300;
  if (x < a + 1) {
    let ap = a, sum = 1 / a, del = sum;
    for (let n = 1; n <= 1000; n++) { ap++; del *= x / ap; sum += del; if (Math.abs(del) < Math.abs(sum) * eps) break; }
    const p = sum * Math.exp(-x + a * Math.log(x) - gln);
    return Math.max(0, Math.min(1, 1 - p));
  }
  let b = x + 1 - a, c = 1 / tiny, d = 1 / Math.max(Math.abs(b), tiny) * Math.sign(b || 1), h = d;
  for (let i = 1; i <= 1000; i++) {
    const an = -i * (i - a); b += 2;
    d = an * d + b; if (Math.abs(d) < tiny) d = tiny;
    c = b + an / c; if (Math.abs(c) < tiny) c = tiny;
    d = 1 / d; const delta = d * c; h *= delta;
    if (Math.abs(delta - 1) < eps) break;
  }
  return Math.max(0, Math.min(1, Math.exp(-x + a * Math.log(x) - gln) * h));
}
function chiSquareTail(x, df) { return x < 0 || df <= 0 ? NaN : regularizedGammaQ(df / 2, x / 2); }
function fTail(x, df1, df2) { return x < 0 || df1 <= 0 || df2 <= 0 ? NaN : regularizedBeta(df2 / (df2 + df1 * x), df2 / 2, df1 / 2); }
function finiteTest(p, method, reason = '') {
  return { p: Number.isFinite(p) && p >= 0 && p <= 1 ? p : null, method, reason: Number.isFinite(p) && p >= 0 && p <= 1 ? '' : (reason || 'The test returned a non-finite or out-of-range result.') };
}
function pText(result) { return result.p === null ? 'NA' : result.p < 0.001 ? `<${localFixed(0.001, 3)}` : localFixed(result.p, 3); }
function pTitle(result) { return result.p === null ? `${result.method}: ${result.reason}` : `${result.method}; complete observations only.`; }
function meanVariance(x) {
  const n = x.length, mean = x.reduce((a, b) => a + b, 0) / n;
  return { n, mean, variance: n > 1 ? x.reduce((s, v) => s + (v - mean) ** 2, 0) / (n - 1) : NaN };
}
function welchT(groups) {
  const [a, b] = groups.map(meanVariance);
  if (a.n < 2 || b.n < 2) return finiteTest(NaN, "Welch's t-test", 'At least two non-missing observations are required in each group.');
  const va = a.variance / a.n, vb = b.variance / b.n, se2 = va + vb;
  if (se2 === 0) return finiteTest(a.mean === b.mean ? 1 : NaN, "Welch's t-test", 'Both groups have zero within-group variance; the test is undefined.');
  const t = (a.mean - b.mean) / Math.sqrt(se2), df = se2 ** 2 / (va ** 2 / (a.n - 1) + vb ** 2 / (b.n - 1));
  return finiteTest(regularizedBeta(df / (df + t * t), df / 2, 0.5), "Welch's t-test");
}
function welchAnova(groups) {
  const s = groups.map(meanVariance), k = s.length;
  if (k < 3) return welchT(groups);
  if (s.some(x => x.n < 2)) return finiteTest(NaN, "Welch's ANOVA", 'At least two non-missing observations are required in every group.');
  if (s.every(x => x.variance === 0)) return finiteTest(s.every(x => x.mean === s[0].mean) ? 1 : NaN, "Welch's ANOVA", 'All groups have zero within-group variance; the test is undefined.');
  if (s.some(x => !(x.variance > 0))) return finiteTest(NaN, "Welch's ANOVA", 'A group has zero within-group variance; Welch’s ANOVA is undefined for this configuration.');
  const w = s.map(x => x.n / x.variance), W = w.reduce((a, b) => a + b, 0), avg = w.reduce((a, wi, i) => a + wi * s[i].mean, 0) / W;
  const numerator = w.reduce((sum, wi, i) => sum + wi * (s[i].mean - avg) ** 2, 0) / (k - 1);
  const term = w.reduce((sum, wi, i) => sum + (1 - wi / W) ** 2 / (s[i].n - 1), 0);
  if (!(term > 0)) return finiteTest(NaN, "Welch's ANOVA", 'The denominator degrees of freedom could not be estimated.');
  const correction = 1 + (2 * (k - 2) / (k * k - 1)) * term, df2 = (k * k - 1) / (3 * term), f = numerator / correction;
  return finiteTest(fTail(f, k - 1, df2), "Welch's ANOVA");
}
function ranked(values) {
  const sorted = values.map((value, i) => ({ value, i })).sort((a, b) => a.value - b.value), ranks = new Array(values.length), ties = [];
  for (let i = 0; i < sorted.length;) {
    let j = i + 1; while (j < sorted.length && sorted[j].value === sorted[i].value) j++;
    const rank = (i + 1 + j) / 2;
    for (let k = i; k < j; k++) ranks[sorted[k].i] = rank;
    if (j - i > 1) ties.push(j - i);
    i = j;
  }
  return { ranks, ties };
}
function combinationsCapped(n, k, limit = 50000) {
  k = Math.min(k, n - k); let x = 1;
  for (let i = 1; i <= k; i++) { x = x * (n - k + i) / i; if (x > limit) return limit + 1; }
  return Math.round(x);
}
function mannWhitney(groups) {
  if (groups.length !== 2 || groups.some(x => x.length < 2)) return finiteTest(NaN, 'Mann–Whitney U', 'At least two non-missing observations are required in each of two groups.');
  const all = [...groups[0], ...groups[1]], { ranks, ties } = ranked(all), n1 = groups[0].length, n2 = groups[1].length, N = n1 + n2;
  const rankSum = ranks.slice(0, n1).reduce((a, b) => a + b, 0), u = rankSum - n1 * (n1 + 1) / 2, mean = n1 * n2 / 2;
  const ways = combinationsCapped(N, n1);
  if (ways <= 50000) {
    let extreme = 0;
    function enumerate(start, selected, sum) {
      if (selected === n1) { const permU = sum - n1 * (n1 + 1) / 2; if (Math.abs(permU - mean) + 1e-12 >= Math.abs(u - mean)) extreme++; return; }
      for (let i = start; i <= N - (n1 - selected); i++) enumerate(i + 1, selected + 1, sum + ranks[i]);
    }
    enumerate(0, 0, 0);
    return finiteTest(extreme / ways, 'Mann–Whitney U (exact permutation)');
  }
  const tie = ties.reduce((sum, t) => sum + t ** 3 - t, 0), variance = n1 * n2 / 12 * (N + 1 - tie / (N * (N - 1)));
  if (variance === 0) return finiteTest(1, 'Mann–Whitney U (tie-corrected asymptotic)');
  const z = Math.max(0, Math.abs(u - mean) - 0.5) / Math.sqrt(variance);
  return finiteTest(regularizedGammaQ(0.5, z * z / 2), 'Mann–Whitney U (tie-corrected asymptotic)');
}
function kruskalWallis(groups) {
  if (groups.length < 2 || groups.some(x => !x.length)) return finiteTest(NaN, 'Kruskal–Wallis', 'At least one non-missing observation is required in every group.');
  const all = groups.flat(), { ranks, ties } = ranked(all), N = all.length;
  const correction = 1 - ties.reduce((a, t) => a + t ** 3 - t, 0) / (N ** 3 - N);
  if (correction === 0) return finiteTest(1, 'Kruskal–Wallis (tie-corrected)');
  const hFor = sums => Math.max(0, (12 * sums.reduce((v, rankSum, i) => v + rankSum * rankSum / groups[i].length, 0) / (N * (N + 1)) - 3 * (N + 1)) / correction);
  const observedSums = []; let offset = 0;
  for (const g of groups) { observedSums.push(ranks.slice(offset, offset + g.length).reduce((a, b) => a + b, 0)); offset += g.length; }
  const observed = hFor(observedSums);
  let ways = 1, remaining = N;
  for (const g of groups) { ways *= combinationsCapped(remaining, g.length); remaining -= g.length; if (ways > 50000) break; }
  if (ways <= 50000) {
    const counts = groups.map(() => 0), sums = groups.map(() => 0); let extreme = 0;
    function assign(pos) {
      if (pos === N) { if (hFor(sums) + 1e-12 >= observed) extreme++; return; }
      for (let i = 0; i < groups.length; i++) if (counts[i] < groups[i].length) { counts[i]++; sums[i] += ranks[pos]; assign(pos + 1); sums[i] -= ranks[pos]; counts[i]--; }
    }
    assign(0);
    return finiteTest(extreme / ways, 'Kruskal–Wallis (exact permutation)');
  }
  return finiteTest(chiSquareTail(observed, groups.length - 1), 'Kruskal–Wallis (tie-corrected asymptotic)');
}
const logChoose = (n, k) => k < 0 || k > n ? -Infinity : logGamma(n + 1) - logGamma(k + 1) - logGamma(n - k + 1);
function fisher2x2(table) {
  const [r1, r2] = table, n1 = r1[0] + r1[1], n2 = r2[0] + r2[1], c1 = r1[0] + r2[0], N = n1 + n2, obs = r1[0];
  const lo = Math.max(0, c1 - n2), hi = Math.min(n1, c1), logDen = logChoose(N, c1);
  const lp = x => logChoose(n1, x) + logChoose(n2, c1 - x) - logDen, observed = lp(obs), probs = [];
  for (let x = lo; x <= hi; x++) probs.push({ x, logp: lp(x) });
  const max = Math.max(...probs.map(x => x.logp)), den = probs.reduce((s, x) => s + Math.exp(x.logp - max), 0);
  const num = probs.filter(x => x.logp <= observed + 1e-10).reduce((s, x) => s + Math.exp(x.logp - max), 0);
  return finiteTest(Math.min(1, num / den), "Fisher's exact test (two-sided probability ordering)");
}
function tablePearson(table) {
  const rows = table.map(r => r.reduce((a, b) => a + b, 0)), cols = table[0].map((_, j) => table.reduce((s, r) => s + r[j], 0)), N = rows.reduce((a, b) => a + b, 0);
  let x2 = 0, expected = [];
  for (let i = 0; i < rows.length; i++) for (let j = 0; j < cols.length; j++) { const e = rows[i] * cols[j] / N; expected.push(e); x2 += (table[i][j] - e) ** 2 / e; }
  return { x2, df: (rows.length - 1) * (cols.length - 1), expected, rowTotals: rows };
}
function fisherFreemanHalton(table, limit = 40000) {
  const nr = table.length, nc = table[0].length, rowMargins = table.map(r => r.reduce((a, b) => a + b, 0));
  const colRemain = table[0].map((_, j) => table.reduce((s, r) => s + r[j], 0));
  const observedLogWeight = -table.flat().reduce((s, n) => s + logGamma(n + 1), 0), weights = [], current = Array.from({ length: nr }, () => Array(nc).fill(0));
  let visited = 0, exceeded = false;
  function nextRow(i) {
    if (exceeded) return;
    if (i === nr - 1) {
      if (colRemain.reduce((a, b) => a + b, 0) !== rowMargins[i]) return;
      current[i] = colRemain.slice(); visited++;
      if (visited > limit) { exceeded = true; return; }
      weights.push({ logw: -current.flat().reduce((s, n) => s + logGamma(n + 1), 0) });
      return;
    }
    function allocate(j, left) {
      if (exceeded) return;
      if (j === nc - 1) {
        if (left < 0 || left > colRemain[j]) return;
        current[i][j] = left; colRemain[j] -= left; nextRow(i + 1); colRemain[j] += left; return;
      }
      const future = colRemain.slice(j + 1).reduce((a, b) => a + b, 0), min = Math.max(0, left - future), max = Math.min(left, colRemain[j]);
      for (let x = min; x <= max; x++) { current[i][j] = x; colRemain[j] -= x; allocate(j + 1, left - x); colRemain[j] += x; if (exceeded) return; }
    }
    allocate(0, rowMargins[i]);
  }
  nextRow(0);
  if (exceeded || !weights.length) return finiteTest(NaN, 'Fisher–Freeman–Halton exact test', `The conditional table space exceeds ${limit.toLocaleString(window.AppI18n?.locale || 'en-US')} tables, so a reliable exact p-value was not computed.`);
  const max = Math.max(...weights.map(w => w.logw)), denominator = weights.reduce((s, w) => s + Math.exp(w.logw - max), 0);
  const numerator = weights.filter(w => w.logw <= observedLogWeight + 1e-10).reduce((s, w) => s + Math.exp(w.logw - max), 0);
  return finiteTest(Math.min(1, numerator / denominator), 'Fisher–Freeman–Halton exact test (conditional probability ordering)');
}
function categoricalAssociation(groups) {
  const levels = [...new Set(groups.flat())].sort((a, b) => String(a).localeCompare(String(b), undefined, { numeric: true }));
  const table = groups.map(group => { const counts = new Map(); group.forEach(x => counts.set(x, (counts.get(x) || 0) + 1)); return levels.map(x => counts.get(x) || 0); });
  if (table.length < 2 || table.some(r => r.reduce((a, b) => a + b, 0) === 0)) return finiteTest(NaN, 'Categorical association test', 'At least one complete observation is required in every group.');
  const activeCols = levels.map((_, j) => j).filter(j => table.some(r => r[j] > 0)), activeRows = table.filter(r => r.reduce((a, b) => a + b, 0) > 0);
  const reduced = activeRows.map(r => activeCols.map(j => r[j]));
  if (reduced[0].length < 2) return finiteTest(1, 'Categorical association test', 'Only one observed outcome category remains among complete observations.');
  const stat = tablePearson(reduced), small = stat.expected.some(e => e < 1) || stat.expected.filter(e => e < 5).length > stat.expected.length * 0.2;
  if (!small) return finiteTest(chiSquareTail(stat.x2, stat.df), "Pearson's chi-square test");
  if (reduced.length === 2 && reduced[0].length === 2) return fisher2x2(reduced);
  return fisherFreemanHalton(reduced);
}
function selectComparisonTest(type, groups, mode) {
  if (type === 'categorical') return categoricalAssociation(groups);
  if (groups.length < 2) return finiteTest(NaN, 'Group comparison', 'At least two non-missing groups are required.');
  if (mode === 'nonparametric') return groups.length === 2 ? mannWhitney(groups) : kruskalWallis(groups);
  return groups.length === 2 ? welchT(groups) : welchAnova(groups);
}

function buildTable1() {
  const all = [...Array(state.rows.length).keys()], g = state.categorical.find(c => c.name === state.t1.group);
  const nat = (a, b) => a.localeCompare(b, undefined, { numeric: true });
  const french = window.AppI18n?.language === 'fr';
  const sets = [], notes = [`N: total number of patients (rows in the file); n: number of values summarized${g ? `; columns are levels of “${displayVariableName(g.name)}”` : ''}.`], pvalueInfo = [];
  if (!g || state.t1.overall) sets.push({ label: 'Overall', idx: all });
  let comparisonSets = [];
  if (g) {
    comparisonSets = [...g.freq.keys()].sort(nat).map(k => ({ label: k, idx: all.filter(i => g.labels[i] === k) }));
    comparisonSets.forEach(s => sets.push(s));
    const miss = all.filter(i => g.labels[i] === null);
    if (miss.length) { sets.push({ label: 'Missing group', idx: miss }); notes.push(`${miss.length} row(s) have a missing “${g.name}” value; they appear under “Missing group” and are included in Overall.`); }
  }
  const includeTests = Boolean(g && state.t1.showPValues);
  const p = test => includeTests ? pText(test) : '';
  const rows = [['N', ...sets.map(s => s.idx.length.toLocaleString(window.AppI18n?.locale || 'en-US')), ...(includeTests ? [''] : [])]], rawRowIndices = new Set();
  pvalueInfo.push('');
  let capped = 0;
  for (const c of state.cols) {
    if (c.type === 'excluded' || state.t1.skip.has(c.name) || !state.t1.selected.has(c.name) || c === g) continue;
    if (c.type === 'numeric') {
      const method = g ? (state.t1.testMode === 'nonparametric' ? 'Median (IQR)' : 'Mean (SD)') : c.method;
      const d = decimalsFor(c), test = includeTests ? selectComparisonTest('numeric', comparisonSets.map(s => s.idx.map(i => c.nums[i]).filter(x => x !== null)), state.t1.testMode) : null;
      rows.push([`${displayVariableName(c.name)}, ${method === 'Mean (SD)' ? 'mean (SD)' : 'median (IQR)'}`, ...sets.map(s => {
        const v = s.idx.map(i => c.nums[i]).filter(x => x !== null);
        return v.length ? fmtSummary(computeStats(v), method, d) : '—'; }), ...(includeTests ? [p(test)] : [])]);
      pvalueInfo.push(test ? pTitle(test) : '');
      if (!french) { rows.push(['  Non-missing n', ...sets.map(s => s.idx.filter(i => c.nums[i] !== null).length.toLocaleString(window.AppI18n?.locale || 'en-US')), ...(includeTests ? [''] : [])]); pvalueInfo.push(''); }
    } else {
      const mode = state.t1.catChoices[c.name] === 'counts' || state.t1.catChoices[c.name] === 'percentage' || state.t1.catChoices[c.name] === 'counts_percentage' ? state.t1.catChoices[c.name] : 'counts_percentage';
      const label = mode === 'counts' ? 'counts' : mode === 'percentage' ? 'percentage' : 'counts (percentage)';
      const test = includeTests ? selectComparisonTest('categorical', comparisonSets.map(s => s.idx.map(i => c.labels[i]).filter(x => x !== null)), state.t1.testMode) : null;
      rows.push([`${displayVariableName(c.name)}, ${label}`, ...sets.map(() => ''), ...(includeTests ? [p(test)] : [])]);
      pvalueInfo.push(test ? pTitle(test) : '');
      const cnt = sets.map(s => { const m = new Map(); let den = 0; for (const i of s.idx) { const l = c.labels[i]; if (l !== null) { m.set(l, (m.get(l) || 0) + 1); den++; } } return { m, den }; });
      const lv = [...c.freq].sort((a, b) => b[1] - a[1] || nat(a[0], b[0]));
      if (lv.length > CONFIG.T1_MAX_LEVELS) { capped++; rows.push([`  ${lv.length} levels (not listed)`, ...sets.map(() => ''), ...(includeTests ? [''] : [])]); pvalueInfo.push(''); }
      else lv.forEach(([k]) => { rows.push([`  ${k}`, ...cnt.map(x => { const n = x.m.get(k) || 0, pct = x.den ? localFixed(100 * n / x.den, 1) + '%' : '—'; return mode === 'counts' ? n.toLocaleString(window.AppI18n?.locale || 'en-US') : mode === 'percentage' ? pct : x.den ? `${n.toLocaleString(window.AppI18n?.locale || 'en-US')} (${pct})` : '—'; }), ...(includeTests ? [''] : [])]); rawRowIndices.add(rows.length - 1); pvalueInfo.push(''); });
      if (!french) { rows.push(['  Non-missing n', ...cnt.map(x => x.den.toLocaleString(window.AppI18n?.locale || 'en-US')), ...(includeTests ? [''] : [])]); pvalueInfo.push(''); }
    }
    const col = c.type === 'numeric' ? c.nums : c.labels;
    const ms = sets.map(s => s.idx.reduce((k, i) => k + (col[i] === null ? 1 : 0), 0));
    if (ms.some(Boolean)) {
      const missingCells = french ? ms.map((n, i) => `${n.toLocaleString(window.AppI18n?.locale || 'fr-FR')} (${localFixed(sets[i].idx.length ? 100 * n / sets[i].idx.length : 0, 1)}%)`) : ms.map(String);
      rows.push([french ? '  Missing data, n (%)' : '  Missing', ...missingCells, ...(includeTests ? [''] : [])]); pvalueInfo.push('');
    }
  }
  if (g) notes.push(includeTests ? `Exploratory, unadjusted p-values compare complete observations across ${comparisonSets.length} non-missing groups; missing group labels are excluded from tests. They are not a measure of baseline balance or clinical importance, and no multiplicity adjustment is applied. Continuous summaries use ${state.t1.testMode === 'nonparametric' ? 'median (IQR)' : 'mean (SD)'}; Welch tests compare means, while Mann–Whitney/Kruskal–Wallis test rank distributions and do not specifically test medians. Categorical variables use one overall association test. Pearson chi-square requires no expected count below 1 and at least 80% of expected counts ≥5; sparse tables use exact tests when feasible. Hover or focus a p-value for the test and any limitation.` : 'Grouped summaries are descriptive. P-values are hidden by default; enable “Show exploratory p-values” only when an unadjusted group comparison is appropriate. These tests do not measure baseline balance or clinical importance.');
  else notes.push('Continuous: mean (SD) or median (IQR) as selected under Summary (Auto is a heuristic, not a normality test). Categorical: counts, percentages of non-missing values, or both as selected under Summary settings (Auto defaults to counts and percentages). Statistics use available observations; missing values are never treated as zero. Non-numeric entries in numeric variables are counted as missing.');
  if (capped) notes.push(`${capped} categorical variable(s) with more than ${CONFIG.T1_MAX_LEVELS} levels are not listed in this table.`);
  return { header: ['Characteristic', ...sets.map(s => s.label), ...(includeTests ? ['Exploratory p-value'] : [])], rows, notes, pvalueInfo, rawRowIndices };
}

function renderT1() {
  $('t1overall').disabled = !state.t1.group;
  $('t1showPValues').disabled = !state.t1.group;
  $('t1showPValues').checked = state.t1.showPValues;
  $('t1variablesLabel').textContent = window.AppI18n?.text(`Select variables (${state.t1.selected.size}/${state.headers.filter(name => name !== state.patientId).length})`) || `Select variables (${state.t1.selected.size}/${state.headers.filter(name => name !== state.patientId).length})`;
  $('t1testMode').value = state.t1.testMode;
  const t = localeTable1();
  const groupLevels = new Set(state.categorical.find(c => c.name === state.t1.group)?.freq.keys() || []);
  const pIndex = t.header.length - 1, hasP = Boolean(state.t1.group && state.t1.showPValues);
  $('t1').innerHTML = `<thead><tr>${t.header.map((h, i) => `<th scope="col"${i ? ' class="num"' : ''}${groupLevels.has(h) ? ' data-raw' : ''}>${esc(h)}</th>`).join('')}</tr></thead><tbody>${
    t.rows.map((r, ri) => `<tr>${r.map((v, i) => { const rawLevel = i === 0 && r[0].startsWith('  ') && !['Missing', 'Non-missing n', 'Missing data, n (%)'].includes(r[0].trim()); return `<td class="${i ? 'num' : r[0].startsWith('  ') ? 'ind' : ''}"${rawLevel ? ' data-raw' : ''}${hasP && i === pIndex && t.pvalueInfo[ri] ? ` title="${esc(t.pvalueInfo[ri])}" tabindex="0"` : ''}>${esc(v.trim())}</td>`; }).join('')}</tr>`).join('')}</tbody>`;
  $('t1notes').innerHTML = t.notes.map(x => `<li>${esc(x)}</li>`).join('');
}
const localeTable1 = () => {
  const t = buildTable1(), tx = x => window.AppI18n?.text(x) || x;
  const rows = t.rows.map((r, ri) => {
    let first = r[0];
    if (t.rawRowIndices.has(ri)) { /* category values stay unchanged */ }
    else if (['Missing', 'Non-missing n', 'Missing data, n (%)'].includes(first.trim())) first = first.startsWith('  ') ? `  ${tx(first.trim())}` : tx(first);
    else first = tx(first);
    return [first, ...r.slice(1)];
  });
  const notes = window.AppI18n?.language === 'fr'
    ? ['N : nombre total de patients (lignes du fichier) ; n : nombre de valeurs concernées.',
       'Variables continues : moyenne (ET) ou médiane (IIQ), selon la statistique choisie dans les statistiques descriptives. Le mode Auto repose sur une règle empirique et ne constitue pas un test de normalité.',
       'Variables catégorielles : effectifs et/ou pourcentages calculés sur les valeurs renseignées (Auto : les deux).',
       'Les statistiques portent sur les données disponibles ; les valeurs manquantes ne sont jamais remplacées par zéro. Toute valeur non numérique dans une variable numérique est traitée comme manquante.',
       ...(state.t1.group && t.notes[1] ? [tx(t.notes[1])] : []), ...t.notes.slice(2).map(tx)]
    : t.notes.map(tx);
  return { header: t.header.map(h => ['Characteristic', 'Exploratory p-value', 'Overall'].includes(h) ? tx(h) : h), rows, notes, pvalueInfo: t.pvalueInfo };
};
const t1Matrix = () => { const t = localeTable1(); return [t.header.map(safe), ...t.rows.map(r => [safe(r[0]), ...r.slice(1)]), [''], ...t.notes.map(x => [x])]; };

function initExtra() {
  document.addEventListener('click', e => {
    document.querySelectorAll('details.menu[open]').forEach(menu => { if (!menu.contains(e.target)) menu.open = false; });
    const v = e.target.closest('[data-view]'), t = e.target.closest('[data-fig]');
    if (v) showView(v.dataset.view);
    else if (t) { state.fig.type = t.dataset.fig; figRefresh(); if (state.view !== 'figures') return; }
  });
  $('figctl').addEventListener('input', e => { const k = e.target.dataset.f; if (!k) return; state.fig[k] = k === 'bins' ? (Math.floor(+e.target.value) || 0) : e.target.value; figDraw(); });
  $('figreset').addEventListener('click', () => { state.fig = null; figRefresh(); });
  $('figsvg').addEventListener('click', () => figCur && download(figCur.svg, figName('svg'), 'image/svg+xml'));
  $('figpng').addEventListener('click', figPng);
  $('figpdf').addEventListener('click', figPdf);
  $('dlDocx').addEventListener('click', () => exportDocx(summaryDocx, baseName() + '.docx'));
  $('t1docx').addEventListener('click', () => exportDocx(table1Docx, baseName().replace('_descriptive_stats', '_table1') + '.docx'));
  /* Variable types controls are temporarily disabled; preserve handler for future use.
  $('types').addEventListener('change', async e => {
    const s = e.target.closest('select[data-tname]'), cb = e.target.closest('input[data-t1]');
    if (s) { if (s.value === 'auto') delete state.types[s.dataset.tname]; else state.types[s.dataset.tname] = s.value; await reanalyze(); }
    else if (cb) { if (cb.checked) state.t1.skip.delete(cb.dataset.t1); else state.t1.skip.add(cb.dataset.t1); renderT1(); }
  });
  */
  $('t1group').addEventListener('change', e => { state.t1.group = e.target.value; renderT1(); });
  $('t1testMode').addEventListener('change', e => { state.t1.testMode = e.target.value === 'nonparametric' ? 'nonparametric' : 'parametric'; renderT1(); });
  $('t1variables').addEventListener('change', e => {
    const input = e.target.closest('[data-t1-variable]'); if (!input) return;
    if (input.checked) state.t1.selected.add(input.dataset.t1Variable); else state.t1.selected.delete(input.dataset.t1Variable);
    renderT1();
  });
  $('catSettings').addEventListener('change', e => {
    const sel = e.target.closest('[data-cat-method]'); if (!sel) return;
    state.t1.catChoices[sel.dataset.catMethod] = sel.value;
    renderCategoricalSettings(); renderTable(); renderT1();
  });
  $('t1selectAll').addEventListener('click', () => { state.t1.selected = new Set(state.headers.filter(name => name !== state.patientId)); fillGroupSelect(); renderT1(); });
  $('t1selectNone').addEventListener('click', () => { state.t1.selected.clear(); fillGroupSelect(); renderT1(); });
  $('t1overall').addEventListener('change', e => { state.t1.overall = e.target.checked; renderT1(); });
  $('t1showPValues').addEventListener('change', e => { state.t1.showPValues = e.target.checked; renderT1(); });
  if ($('catsearch')) $('catsearch').addEventListener('input', renderCategorical);
  /* Table 1 Copy action temporarily disabled; retained for future use.
  $('t1copy').addEventListener('click', () => copyText(t1Matrix().map(r => r.map(tsvCell).join('\t')).join('\n'), 'Table 1 copied (tab-separated).'));
  */
  $('t1csv').addEventListener('click', () => download(t1Matrix().map(r => r.map(csvCell).join(',')).join('\r\n'), baseName().replace('_descriptive_stats', '_table1') + '.csv', 'text/csv'));
  $('stats').querySelector('tbody').addEventListener('change', renderT1);
  document.querySelectorAll('[data-setall]').forEach(b => b.addEventListener('click', renderT1));
}


// ---------- Views & Figures (SVG generated from the loaded data; no charting library) ----------
let figCur = null;
const PAL = ['#1f4e79', '#E69F00', '#009E73', '#CC79A7', '#56B4E9', '#D55E00', '#999999'];   // Okabe-Ito based
const FIGS = {
  hist: { name: 'Histogram', info: 'Distribution of a numeric variable', needs: 'Numeric', icon: '<path d="M3 20V12h4v8M9 20V5h4v15M15 20V9h4v11"/>' },
  box: { name: 'Box plot', info: 'Median, IQR, Tukey outliers; optional groups', needs: 'Numeric + optional categorical', icon: '<path d="M6 8h12v8H6zM12 8V3M12 16v5M6 12h12"/>' },
  bar: { name: 'Bar chart', info: 'Category counts or percentages', needs: 'Categorical', icon: '<path d="M4 6h12M4 12h16M4 18h8"/>' },
  scatter: { name: 'Scatter plot', info: 'Two numeric variables', needs: 'Two numeric + optional categorical', icon: '<circle cx="6" cy="17" r="1.5"/><circle cx="11" cy="11" r="1.5"/><circle cx="17" cy="14" r="1.5"/><circle cx="19" cy="6" r="1.5"/>' },
  miss: { name: 'Missingness', info: 'Missing values per variable', needs: 'All variables', icon: '<path d="M4 5h16M4 10h10M4 15h6M4 20h13"/>' }
};
const tl = v => new Intl.NumberFormat(window.AppI18n?.locale || 'en-US', { maximumSignificantDigits: 6, useGrouping: false }).format(Number(v));
const cut = (s, n) => s.length > n ? s.slice(0, n - 1) + '…' : s;
const lin = (d0, d1, r0, r1) => v => r0 + (v - d0) / ((d1 - d0) || 1) * (r1 - r0);
const numVar = n => state.results.find(r => r.name === n), catVar = n => state.categorical.find(r => r.name === n);
const byFreq = c => [...c.freq].sort((a, b) => b[1] - a[1] || a[0].localeCompare(b[0], undefined, { numeric: true }));
function niceTicks(lo, hi, n = 5) {
  if (!(hi > lo)) return [lo];
  const raw = (hi - lo) / n, p = 10 ** Math.floor(Math.log10(raw)), f = raw / p, step = (f < 1.5 ? 1 : f < 3 ? 2 : f < 7 ? 5 : 10) * p, out = [];
  for (let v = Math.ceil(lo / step - 1e-9) * step; v <= hi + step * 1e-9; v += step) out.push(+v.toPrecision(12));
  return out;
}
const T = (x, y, s, a = 'middle', fill = '#1d2430', sz = 12) => `<text x="${x}" y="${y}" text-anchor="${a}" font-size="${sz}" fill="${fill}">${esc(s)}</text>`;
const yAx = (d, sy, tk) => tk.map(v => `<line x1="${d.l}" x2="${d.W - d.r}" y1="${sy(v)}" y2="${sy(v)}" stroke="#e3e7ec"/>${T(d.l - 6, sy(v) + 4, tl(v), 'end', '#5b6573', 11)}`).join('') + `<line x1="${d.l}" x2="${d.l}" y1="${d.t}" y2="${d.H - d.b}" stroke="#5b6573"/>`;
const xAx = (d, sx, tk) => tk.map(v => `<line x1="${sx(v)}" x2="${sx(v)}" y1="${d.H - d.b}" y2="${d.H - d.b + 5}" stroke="#5b6573"/>${T(sx(v), d.H - d.b + 18, tl(v), 'middle', '#5b6573', 11)}`).join('') + `<line x1="${d.l}" x2="${d.W - d.r}" y1="${d.H - d.b}" y2="${d.H - d.b}" stroke="#5b6573"/>`;
function frame(d, inner, title, xl, yl, extra = '') {
  return `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 ${d.W} ${d.H}" width="${d.W}" height="${d.H}" font-family="system-ui, Arial, sans-serif" role="img" aria-label="${esc(title)}"><title>${esc(title)}</title><rect width="${d.W}" height="${d.H}" fill="#fff"/><text x="${d.W / 2}" y="22" text-anchor="middle" font-size="15" font-weight="600" fill="#1d2430">${esc(title)}</text>${extra}${inner}${T((d.l + d.W - d.r) / 2, d.H - 10, xl)}${yl ? `<text transform="translate(16 ${(d.t + d.H - d.b) / 2}) rotate(-90)" text-anchor="middle" font-size="12" fill="#1d2430">${esc(yl)}</text>` : ''}</svg>`;
}
const need = kind => ({ msg: (kind === 'cat' ? state.categorical : state.results).length ? 'Choose a variable.' : (kind === 'cat' ? 'No categorical variables were detected. Try a histogram, box plot or the missingness plot.' : 'No numeric variables were detected. Try a bar chart or the missingness plot.') });
const done = (svg, d, desc) => ({ svg, w: d.W, h: d.H, desc });

function figHist(f) {
  const r = numVar(f.x); if (!r) return need('num');
  const v = r.nums.filter(x => x !== null), constant = r.min === r.max;
  const k = constant ? 1 : f.bins > 0 ? Math.min(100, f.bins) : Math.max(5, Math.min(40, Math.ceil(Math.log2(v.length) + 1)));   // Sturges default
  const a = constant ? r.min - 0.5 : r.min, b = constant ? r.max + 0.5 : r.max, w = (b - a) / k, c = new Array(k).fill(0);
  v.forEach(x => c[Math.min(k - 1, Math.floor((x - a) / w))]++);
  const d = { W: 640, H: 400, l: 62, r: 20, t: 40, b: 56 }, ymax = Math.max(...c), sx = lin(a, b, d.l, d.W - d.r), sy = lin(0, ymax, d.H - d.b, d.t);
  const bars = c.map((n, i) => `<rect x="${sx(a + i * w) + 0.5}" y="${sy(n)}" width="${Math.max(0, sx(a + (i + 1) * w) - sx(a + i * w) - 1)}" height="${sy(0) - sy(n)}" fill="${PAL[0]}"/>`).join('');
  return done(frame(d, yAx(d, sy, niceTicks(0, ymax)) + bars + xAx(d, sx, niceTicks(a, b)), `Distribution of ${cut(r.name, 50)}`, r.name, 'Frequency'), d,
    `Histogram of ${r.name}: n = ${v.length} plotted; ${state.rows.length - v.length} missing or non-numeric excluded; ${k} bin${k === 1 ? '' : 's'} of width ${tl(w)}.`);
}

function figBox(f) {
  const r = numVar(f.x); if (!r) return need('num');
  const g = f.g ? catVar(f.g) : null; let excl = 0, hiddenObs = 0, hiddenLv = 0, groups;
  if (g) {
    const shown = byFreq(g).map(x => x[0]); hiddenLv = Math.max(0, shown.length - 12); shown.length = Math.min(12, shown.length);
    groups = shown.map(k => ({ label: k, v: [] })); const ix = new Map(shown.map((k, i) => [k, i]));
    r.nums.forEach((x, i) => { if (x === null) return; const l = g.labels[i]; if (l === null) excl++; else if (ix.has(l)) groups[ix.get(l)].v.push(x); else hiddenObs++; });
  } else groups = [{ label: 'All', v: r.nums.filter(x => x !== null) }];
  groups = groups.filter(x => x.v.length);
  if (!groups.length) return { msg: 'No valid observations for this selection.' };
  const all = groups.flatMap(x => x.v), lo = Math.min(...all), hi = Math.max(...all), pad = (hi - lo || 1) * 0.06;
  const d = { W: 640, H: 400, l: 62, r: 20, t: 40, b: 70 }, sy = lin(lo - pad, hi + pad, d.H - d.b, d.t), bw = (d.W - d.l - d.r) / groups.length, bx = Math.min(60, bw * 0.5);
  const body = groups.map((G, i) => {
    const s = computeStats(G.v), lf = s.q1 - 1.5 * s.iqr, hf = s.q3 + 1.5 * s.iqr, inside = G.v.filter(x => x >= lf && x <= hf), cx = d.l + (i + 0.5) * bw, col = PAL[g ? i % PAL.length : 0];
    const wl = Math.min(...inside), wh = Math.max(...inside);
    return `<g stroke="#1d2430" stroke-width="1.5"><line x1="${cx}" x2="${cx}" y1="${sy(wl)}" y2="${sy(s.q1)}"/><line x1="${cx}" x2="${cx}" y1="${sy(s.q3)}" y2="${sy(wh)}"/><line x1="${cx - bx / 4}" x2="${cx + bx / 4}" y1="${sy(wl)}" y2="${sy(wl)}"/><line x1="${cx - bx / 4}" x2="${cx + bx / 4}" y1="${sy(wh)}" y2="${sy(wh)}"/><rect x="${cx - bx / 2}" y="${sy(s.q3)}" width="${bx}" height="${Math.max(1, sy(s.q1) - sy(s.q3))}" fill="${col}" fill-opacity=".35"/><line x1="${cx - bx / 2}" x2="${cx + bx / 2}" y1="${sy(s.median)}" y2="${sy(s.median)}" stroke-width="3"/></g>`
      + G.v.filter(x => x < lf || x > hf).map(x => `<circle cx="${cx}" cy="${sy(x)}" r="3" fill="none" stroke="#1d2430"/>`).join('')
      + (g ? T(cx, d.H - d.b + 16, cut(G.label, 12), 'middle', '#1d2430', 11) : '') + T(cx, d.H - d.b + (g ? 30 : 16), `n=${G.v.length}`, 'middle', '#5b6573', 10);
  }).join('');
  const ax = yAx(d, sy, niceTicks(lo - pad, hi + pad)) + `<line x1="${d.l}" x2="${d.W - d.r}" y1="${d.H - d.b}" y2="${d.H - d.b}" stroke="#5b6573"/>`;
  return done(frame(d, ax + body, g ? `${cut(r.name, 30)} by ${cut(g.name, 25)}` : cut(r.name, 60), g ? g.name : '', r.name), d,
    `Box plot of ${r.name}${g ? ` by ${g.name}` : ''}: n = ${all.length} plotted. Box = Q1–Q3 (R type 7) with median line; whiskers reach the most extreme values within 1.5×IQR; circles beyond are outliers and are not removed. Excluded: ${r.nums.length - r.nums.filter(x => x !== null).length} missing/non-numeric${g ? `; ${excl} with missing group; ${hiddenObs} in ${hiddenLv} less frequent level(s) not shown (max 12 groups)` : ''}.`);
}

function figBar(f) {
  const c = catVar(f.c); if (!c) return need('cat');
  const lv = byFreq(c), shown = lv.slice(0, 15), hid = lv.length - shown.length, pct = f.mode === 'pct';
  const d = { W: 640, H: 104 + 26 * shown.length, l: 140, r: 90, t: 40, b: 56 }, vals = shown.map(([, n]) => pct ? 100 * n / c.nonMissing : n), mx = Math.max(...vals), sx = lin(0, mx, d.l, d.W - d.r);
  const bars = shown.map(([k, n], i) => { const y = d.t + 8 + i * 26, p = 100 * n / c.nonMissing; return T(d.l - 6, y + 14, cut(k, 20), 'end', '#1d2430', 11) + `<rect x="${d.l}" y="${y}" width="${Math.max(0, sx(vals[i]) - d.l)}" height="20" fill="${PAL[0]}"/>` + T(sx(vals[i]) + 4, y + 14, `${n.toLocaleString(window.AppI18n?.locale || 'en-US')} (${localFixed(p, 1)}%)`, 'start', '#1d2430', 11); }).join('');
  return done(frame(d, xAx(d, sx, niceTicks(0, mx)) + bars, cut(c.name, 60), pct ? `Percent of non-missing (n = ${c.nonMissing})` : 'Count', ''), d,
    `Bar chart of ${c.name}: ${c.nonMissing} non-missing observations; ${c.missing} missing excluded from percentages.${hid ? ` Showing the 15 most frequent of ${lv.length} levels; ${hid} not shown.` : ''}`);
}

function figMiss() {
  const n = state.rows.length;
  const items = state.cols.filter(c => (c.type !== 'excluded' || c.empty) && c.name !== state.patientId).map(c => [c.name, c.missing]).sort((a, b) => b[1] - a[1] || a[0].localeCompare(b[0]));
  const d = { W: 640, H: 104 + 22 * items.length, l: 150, r: 90, t: 40, b: 56 }, sx = lin(0, 100, d.l, d.W - d.r);
  const bars = items.map(([h, m], i) => { const y = d.t + 8 + i * 22, p = 100 * m / n; return T(d.l - 6, y + 13, cut(h, 22), 'end', '#1d2430', 11) + `<rect x="${d.l}" y="${y}" width="${Math.max(0, sx(p) - d.l)}" height="16" fill="${PAL[0]}"/>` + T(sx(p) + 4, y + 12, `${m.toLocaleString(window.AppI18n?.locale || 'en-US')} (${localFixed(p, 1)}%)`, 'start', '#1d2430', 11); }).join('');
  return done(frame(d, xAx(d, sx, [0, 25, 50, 75, 100]) + bars, 'Missing values by variable', `Percent of ${n} rows missing`, ''), d,
    `Missingness for ${items.length} analyzed variables (${n} rows), sorted by percent missing. Numeric columns also count non-numeric entries as missing. Missing = blank, NA, N/A, NaN, null, missing, "." or "-".`);
}

function figScatter(f) {
  const X = numVar(f.x), Y = numVar(f.y); if (!X || !Y) return need('num');
  let g = f.g ? catVar(f.g) : null, note = '';
  if (g && g.unique > 7) { note = ` Color-by “${g.name}” ignored (more than 7 levels).`; g = null; }
  const lv = g ? byFreq(g).map(x => x[0]) : [], pts = []; let exclG = 0;
  state.rows.forEach((_, i) => { const x = X.nums[i], y = Y.nums[i]; if (x === null || y === null) return; if (g && g.labels[i] === null) { exclG++; return; } pts.push([x, y, g ? lv.indexOf(g.labels[i]) : 0]); });
  if (!pts.length) return { msg: 'No observations have valid values for both variables.' };
  const rng = (a, b) => a === b ? [a - 1, b + 1] : [a - (b - a) * 0.04, b + (b - a) * 0.04];
  const [x0, x1] = rng(Math.min(...pts.map(p => p[0])), Math.max(...pts.map(p => p[0]))), [y0, y1] = rng(Math.min(...pts.map(p => p[1])), Math.max(...pts.map(p => p[1])));
  const d = { W: 640, H: 400, l: 62, r: 20, t: g ? 52 : 40, b: 56 }, sx = lin(x0, x1, d.l, d.W - d.r), sy = lin(y0, y1, d.H - d.b, d.t);
  let lx = d.l; const leg = g ? lv.map((k, i) => { const s = `<circle cx="${lx + 5}" cy="36" r="4" fill="${PAL[i]}"/>${T(lx + 13, 40, cut(k, 14), 'start', '#1d2430', 11)}`; lx += 28 + Math.min(14, k.length) * 6; return s; }).join('') : '';
  const dots = pts.map(p => `<circle cx="${sx(p[0])}" cy="${sy(p[1])}" r="3" fill="${PAL[p[2]]}" fill-opacity=".65"/>`).join('');
  return done(frame(d, yAx(d, sy, niceTicks(y0, y1)) + xAx(d, sx, niceTicks(x0, x1)) + dots, `${cut(Y.name, 28)} vs ${cut(X.name, 28)}`, X.name, Y.name, leg), d,
    `Scatter plot: ${pts.length} of ${state.rows.length} observations shown (pairs with a missing or non-numeric X or Y excluded${g ? `; ${exclG} excluded for missing ${g.name}` : ''}). No trend line is fitted.${note}`);
}

function figRefresh() {
  const num = state.results.map(r => r.name), cat = state.categorical.map(r => r.name);
  const f = state.fig || (state.fig = { type: 'hist', x: '', y: '', c: '', g: '', bins: 0, mode: 'count' });
  if (!num.includes(f.x)) f.x = num[0] || '';
  if (!num.includes(f.y)) f.y = num[1] || num[0] || '';
  if (!cat.includes(f.c)) f.c = cat[0] || '';
  if (f.g && !cat.includes(f.g)) f.g = '';
  $('figtypes').innerHTML = Object.entries(FIGS).map(([k, v]) => `<button type="button" class="figtype" data-fig="${k}" aria-pressed="${f.type === k}"><svg viewBox="0 0 24 24" aria-hidden="true">${v.icon}</svg><strong>${v.name}</strong><span>${v.info}</span><em>${v.needs}</em></button>`).join('');
  const S = (id, label, list, val, none) => `<label for="fig${id}">${label}</label><select id="fig${id}" class="method" data-f="${id}">${none ? '<option value="">None</option>' : ''}${list.map(n => `<option data-raw value="${esc(n)}"${n === val ? ' selected' : ''}>${esc(n)}</option>`).join('')}</select>`;
  $('figctl').innerHTML = { hist: () => S('x', 'Variable', num, f.x) + `<label for="figbins">Bins</label><input id="figbins" data-f="bins" type="number" min="1" max="100" placeholder="Auto" value="${f.bins || ''}">`,
    box: () => S('x', 'Variable', num, f.x) + S('g', 'Group by', cat, f.g, true), bar: () => S('c', 'Variable', cat, f.c) + `<label for="figmode">Show</label><select id="figmode" class="method" data-f="mode"><option value="count"${f.mode === 'count' ? ' selected' : ''}>Counts</option><option value="pct"${f.mode === 'pct' ? ' selected' : ''}>Percent</option></select>`,
    miss: () => '', scatter: () => S('x', 'X', num, f.x) + S('y', 'Y', num, f.y) + S('g', 'Color by', cat, f.g, true) }[f.type]();
  if (state.view === 'figures') figDraw();
}
function figDraw() {
  const f = state.fig, out = { hist: figHist, box: figBox, bar: figBar, miss: figMiss, scatter: figScatter }[f.type](f);
  const rawLabels = [...state.headers, ...state.categorical.flatMap(c => [...c.freq.keys()])];
  const localizedSvg = out.svg ? (window.AppI18n?.markup(out.svg, rawLabels) || out.svg) : '';
  figCur = out.svg ? { ...out, svg: localizedSvg } : null;
  $('figplot').innerHTML = localizedSvg || `<p class="note">${esc(window.AppI18n?.text(out.msg) || out.msg)}</p>`;
  const p = $('figplot'); p.classList.remove('fade-soft'); void p.offsetWidth; p.classList.add('fade-soft');
  $('figdesc').textContent = window.AppI18n?.text(out.desc || '') || out.desc || '';
  $('figsvg').disabled = $('figpng').disabled = $('figpdf').disabled = !figCur;
}
const figName = ext => baseName().replace('_descriptive_stats', '_fig_' + state.fig.type) + '.' + ext;
function fadeShow(el) { el.hidden = false; el.classList.remove('fade-out'); el.classList.remove('fade-in'); void el.offsetWidth; el.classList.add('fade-in'); }
function fadeHide(el, done) {
  if (el.hidden) { if (done) done(); return; }
  el.classList.add('fade-out');
  setTimeout(() => { el.hidden = true; el.classList.remove('fade-out'); if (done) done(); }, 200);
}
// The workspace chooser (#home) stays visible; it shrinks into a switcher once a workspace is open.
function showView(v) {
  const prev = state.view;
  if (prev === v && v !== 'home') return;
  state.view = v;
  $('home').classList.toggle('compact', v !== 'home');
  document.querySelectorAll('.path').forEach(b => b.setAttribute('aria-pressed', String(b.dataset.view === v)));
  if (v === 'home') { ['tables', 'figures'].forEach(k => { $(k).hidden = true; }); return; }
  const open = () => { fadeShow($(v)); if (v === 'figures') figRefresh(); const h = $(v).querySelector('[data-h]'); if (h) h.focus({ preventScroll: true }); };
  if (prev !== 'home' && !$(prev).hidden) fadeHide($(prev), open); else open();
}


// ---------- Word (.docx) and PDF export: dependency-free ----------
const crcTable = (() => { const t = new Uint32Array(256); for (let n = 0; n < 256; n++) { let c = n; for (let k = 0; k < 8; k++) c = c & 1 ? 0xEDB88320 ^ (c >>> 1) : c >>> 1; t[n] = c >>> 0; } return t; })();
const crc32 = b => { let c = -1; for (let i = 0; i < b.length; i++) c = crcTable[(c ^ b[i]) & 255] ^ (c >>> 8); return (c ^ -1) >>> 0; };
function zipStore(files) {   // uncompressed ZIP container (valid for OOXML)
  const enc = new TextEncoder(), parts = [], cen = [], u16 = n => [n & 255, (n >>> 8) & 255], u32 = n => [n & 255, (n >>> 8) & 255, (n >>> 16) & 255, (n >>> 24) & 255];
  let off = 0;
  for (const f of files) {
    const nm = enc.encode(f.name), crc = crc32(f.data), sz = f.data.length;
    const common = [...u16(0x0800), ...u16(0), ...u16(0), ...u16(0x21), ...u32(crc), ...u32(sz), ...u32(sz), ...u16(nm.length), ...u16(0)];
    parts.push(new Uint8Array([0x50, 0x4b, 3, 4, ...u16(20), ...common]), nm, f.data);
    cen.push(new Uint8Array([0x50, 0x4b, 1, 2, ...u16(20), ...u16(20), ...common, ...u16(0), ...u16(0), ...u16(0), ...u32(0), ...u32(off)]), nm);
    off += 30 + nm.length + sz;
  }
  const cdSize = cen.reduce((s, b) => s + b.length, 0);
  const end = new Uint8Array([0x50, 0x4b, 5, 6, ...u16(0), ...u16(0), ...u16(files.length), ...u16(files.length), ...u32(cdSize), ...u32(off), ...u16(0)]);
  const all = [...parts, ...cen, end], out = new Uint8Array(all.reduce((s, b) => s + b.length, 0)); let p = 0;
  for (const b of all) { out.set(b, p); p += b.length; }
  return out;
}
const xe = s => String(s).replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g, '').replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
function docxBytes(title, header, rows, notes) {
  const line = k => `<w:${k} w:val="single" w:sz="4" w:space="0" w:color="AAB3BF"/>`;
  const run = (t, o = '') => `<w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/>${o}</w:rPr><w:t xml:space="preserve">${xe(t)}</w:t></w:r>`;
  const cell = (t, i, head, last) => {
    const lead = i === 0 ? Math.floor((t.length - t.trimStart().length) / 2) : 0;
    return `<w:tc><w:tcPr><w:tcBorders>${line('top')}${line('left')}${line('bottom')}${line('right')}</w:tcBorders></w:tcPr><w:p><w:pPr><w:spacing w:before="20" w:after="20"/>${lead ? `<w:ind w:left="${lead * 240}"/>` : ''}<w:jc w:val="${i ? 'center' : 'left'}"/></w:pPr>${run(t.trim(), `${head ? '<w:b/>' : ''}<w:sz w:val="20"/>`)}</w:p></w:tc>`;
  };
  const tr = (r, head, last) => `<w:tr>${r.map((t, i) => cell(String(t), i, head, last)).join('')}</w:tr>`;
  const para = (t, o) => `<w:p><w:pPr><w:spacing w:before="0" w:after="120"/></w:pPr>${run(t, o)}</w:p>`;
  const wide = header.length > 6, tbl = `<w:tbl><w:tblPr><w:tblW w:w="5000" w:type="pct"/><w:tblLayout w:type="autofit"/><w:tblBorders>${line('top')}${line('left')}${line('bottom')}${line('right')}${line('insideH')}${line('insideV')}</w:tblBorders></w:tblPr><w:tblGrid>${header.map((_, i) => `<w:gridCol w:w="${i ? 1500 : 3000}"/>`).join('')}</w:tblGrid>${tr(header, true, false)}${rows.map((r, i) => tr(r, false, i === rows.length - 1)).join('')}</w:tbl>`;
  const body = para(title, '<w:b/><w:sz w:val="22"/>') + tbl + (notes.length ? notes.map(n => para(n, '<w:i/><w:sz w:val="18"/>')).join('') : para('', '')) +
    `<w:sectPr><w:pgSz w:w="${wide ? 15840 : 12240}" w:h="${wide ? 12240 : 15840}"${wide ? ' w:orient="landscape"' : ''}/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>`;
  const X = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>', enc = new TextEncoder();
  return zipStore([
    { name: '[Content_Types].xml', data: enc.encode(`${X}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>`) },
    { name: '_rels/.rels', data: enc.encode(`${X}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>`) },
    { name: 'word/document.xml', data: enc.encode(`${X}<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>${body}</w:body></w:document>`) }
  ]);
}
const DOCX_MIME = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document';
function exportDocx(bytesFn, filename) {
  try { download(new Blob([bytesFn()], { type: DOCX_MIME }), filename, DOCX_MIME); setStatus('Word file downloaded.'); }
  catch (e) { setStatus('Word export failed. Use Copy or CSV instead.'); }
}
function summaryDocx() {
  const v = getView();
  const tx = x => window.AppI18n?.text(x) || x;
  return docxBytes(tx('Descriptive statistics'), ['Variable', 'n', 'Missing data, n (%)', 'Value', 'Statistic'].map(tx), v.map(r => [displayVariableName(r.name), r.type === 'numeric' ? r.n.toLocaleString(window.AppI18n?.locale || 'en-US') : r.nonMissing.toLocaleString(window.AppI18n?.locale || 'en-US'), missingText(r), summaryValue(r), tx(r.type === 'numeric' ? (r.choice === 'auto' ? `Auto: ${r.autoMethod}` : r.method) : (({ auto: 'Auto: Counts (percentage)', counts: 'Counts', percentage: 'Percentage', counts_percentage: 'Counts (percentage)' })[state.t1.catChoices[r.name] || 'auto']))]),
    [tx(`Includes all ${v.length} detected numeric and categorical variables.`),
     tx('Summary method as selected per variable (Auto is a heuristic, not a normality test). SD is the sample SD (n−1); quartiles use R type 7. Missing = blank, NA, N/A, NaN, null, missing, "." or "-".')]);
}
function table1Docx() { const t = localeTable1(); return docxBytes(window.AppI18n?.text('Table 1 baseline characteristics') || 'Table 1 baseline characteristics', t.header, t.rows, t.notes); }

function buildPdf(jpg, pxW, pxH, w, h) {   // one page, one embedded JPEG
  const m = 36, enc = s => new TextEncoder().encode(s), parts = [], off = []; let len = 0;
  const push = b => { parts.push(b); len += b.length; };
  const obj = (n, body) => { off[n] = len; push(enc(`${n} 0 obj\n${body}\nendobj\n`)); };
  push(enc('%PDF-1.4\n'));
  obj(1, '<< /Type /Catalog /Pages 2 0 R >>'); obj(2, '<< /Type /Pages /Kids [3 0 R] /Count 1 >>');
  obj(3, `<< /Type /Page /Parent 2 0 R /MediaBox [0 0 ${+(w + 2 * m).toFixed(2)} ${+(h + 2 * m).toFixed(2)}] /Resources << /XObject << /Im0 4 0 R >> >> /Contents 5 0 R >>`);
  off[4] = len; push(enc(`4 0 obj\n<< /Type /XObject /Subtype /Image /Width ${pxW} /Height ${pxH} /ColorSpace /DeviceRGB /BitsPerComponent 8 /Filter /DCTDecode /Length ${jpg.length} >>\nstream\n`)); push(jpg); push(enc('\nendstream\nendobj\n'));
  const cs = `q ${+w.toFixed(2)} 0 0 ${+h.toFixed(2)} ${m} ${m} cm /Im0 Do Q`;
  obj(5, `<< /Length ${cs.length} >>\nstream\n${cs}\nendstream`);
  const xr = len; let x = 'xref\n0 6\n0000000000 65535 f \n';
  for (let i = 1; i <= 5; i++) x += String(off[i]).padStart(10, '0') + ' 00000 n \n';
  push(enc(`${x}trailer\n<< /Size 6 /Root 1 0 R >>\nstartxref\n${xr}\n%%EOF\n`));
  const out = new Uint8Array(len); let p = 0; for (const b of parts) { out.set(b, p); p += b.length; } return out;
}
function figRaster(mime, cb) {   // renders the displayed SVG at 3x onto a white canvas
  const cur = figCur, url = URL.createObjectURL(new Blob([cur.svg], { type: 'image/svg+xml' })), img = new Image(), s = 3;
  img.onload = () => {
    const cv = document.createElement('canvas'); cv.width = cur.w * s; cv.height = cur.h * s;
    const g = cv.getContext('2d'); g.fillStyle = '#fff'; g.fillRect(0, 0, cv.width, cv.height); g.drawImage(img, 0, 0, cv.width, cv.height);
    URL.revokeObjectURL(url); cv.toBlob(b => cb(b, cv.width, cv.height), mime, 0.95);
  };
  img.onerror = () => { URL.revokeObjectURL(url); cb(null); };
  img.src = url;
}
function figPng() { if (figCur) figRaster('image/png', b => b ? download(b, figName('png'), 'image/png') : setStatus('PNG export failed. Use SVG.')); }
function figPdf() {
  if (!figCur) return; const cur = figCur;
  figRaster('image/jpeg', async (b, pw, ph) => {
    if (!b) { setStatus('PDF export failed. Use SVG.'); return; }
    const w = 468;   // 6.5 in wide; height keeps the figure's aspect ratio
    download(buildPdf(new Uint8Array(await b.arrayBuffer()), pw, ph, w, w * cur.h / cur.w), figName('pdf'), 'application/pdf');
  });
}

// ---------- Reset & events ----------
function reset() {
  // About page reset logic temporarily disabled: aboutReturnState = null; $('aboutPage').hidden = true;
  document.querySelector('.page-header').hidden = false;
  $('landing').hidden = false;
  Object.assign(state, { fileName: '', patientId: '', headers: [], rows: [], results: [], categorical: [], excluded: [], cols: [], types: {}, choices: {}, t1: { group: '', overall: true, testMode: 'parametric', showPValues: false, skip: new Set(), selected: new Set(), catChoices: {} }, sortKey: null, sortDir: 1 });
  state.fig = null; state.view = 'home'; $('warnings').hidden = true; $('warnings').innerHTML = '';
  state.expandedCategoricals = new Set();
  // Patient ID notice banner disabled by request: $('patientNotice').hidden = true; $('patientNotice').textContent = '';
  if ($('detail').open) $('detail').close();
  $('file').value = ''; fadeHide($('results'), () => { fadeShow($('landing')); replayPageFade(); }); showError(''); setStatus('');
}

function replayPageFade() {
  const page = document.querySelector('.wrap');
  page.classList.remove('fade-in');
  void page.offsetWidth;
  page.classList.add('fade-in');
}

/* About the Author navigation logic temporarily disabled; retained for future use.
function openAbout() {
  if (aboutReturnState) return;
  const ids = ['.page-header', 'landing', 'status', 'error', 'results'];
  aboutReturnState = Object.fromEntries(ids.map(id => { const el = id.startsWith('.') ? document.querySelector(id) : $(id); return [id, el.hidden]; }));
  ids.forEach(id => { const el = id.startsWith('.') ? document.querySelector(id) : $(id); el.hidden = true; });
  $('aboutPage').hidden = false;
  $('aboutTitle').focus({ preventScroll: true });
  $('aboutPage').classList.remove('fade-soft'); void $('aboutPage').offsetWidth; $('aboutPage').classList.add('fade-soft');
}

function closeAbout() {
  if (!aboutReturnState) return;
  for (const [id, wasHidden] of Object.entries(aboutReturnState)) {
    const el = id.startsWith('.') ? document.querySelector(id) : $(id); el.hidden = wasHidden;
  }
  aboutReturnState = null;
  $('aboutPage').hidden = true;
  $('aboutOpen').focus({ preventScroll: true });
}
*/

function init() {
  $('copyrightYear').textContent = String(new Date().getFullYear());
  /* About the Author button handlers temporarily disabled; retained for future use.
  $('aboutOpen').addEventListener('click', openAbout);
  $('aboutBack').addEventListener('click', closeAbout);
  */
  $('sheetContinue').addEventListener('click', () => $('sheetPrompt').close('choose'));
  $('sheetCancel').addEventListener('click', () => $('sheetPrompt').close('cancel'));
  $('patientColumn').addEventListener('change', e => { $('patientYes').disabled = !e.target.value; });
  $('patientYes').addEventListener('click', () => { if ($('patientColumn').value) $('patientPrompt').close('yes'); });
  $('patientNo').addEventListener('click', () => $('patientPrompt').close('no'));
  $('file').addEventListener('change', e => handleFile(e.target.files[0]));
  const drop = $('drop');
  ['dragenter', 'dragover'].forEach(t => drop.addEventListener(t, e => { e.preventDefault(); drop.classList.add('over'); }));
  ['dragleave', 'drop'].forEach(t => drop.addEventListener(t, e => { e.preventDefault(); drop.classList.remove('over'); }));
  drop.addEventListener('drop', e => handleFile(e.dataTransfer.files[0]));
  ['dragover', 'drop'].forEach(t => window.addEventListener(t, e => e.preventDefault()));
  $('sample').addEventListener('click', () => loadText(sampleCSV(), 'sample_data.csv', { skipPatientPrompt: true }));
  $('reset').addEventListener('click', reset);
  $('restart').addEventListener('click', reset);
  document.querySelectorAll('#stats [data-sort]').forEach(b => b.addEventListener('click', () => {
    const k = b.dataset.sort;
    state.sortDir = state.sortKey === k ? -state.sortDir : 1; state.sortKey = k; renderTable();
  }));
  const tbody = $('stats').querySelector('tbody');
  tbody.addEventListener('click', e => {
    const expand = e.target.closest('button[data-expand-cat]');
    if (expand) { e.preventDefault(); const name = expand.dataset.expandCat; state.expandedCategoricals.has(name) ? state.expandedCategoricals.delete(name) : state.expandedCategoricals.add(name); renderTable(); return; }
    if (e.target.closest('select')) return;
    const tr = e.target.closest('tr[data-name]'); if (tr) showDetail(tr.dataset.name);
  });
  // Per-variable method change: update only that row's summary cell (keeps focus on the select)
  tbody.addEventListener('change', e => {
    const sel = e.target.closest('select.method'); if (!sel) return;
    if (sel.dataset.catMethod) { state.t1.catChoices[sel.dataset.catMethod] = sel.value; renderCategoricalSettings(); renderTable(); renderT1(); return; }
    const r = state.results.find(x => x.name === sel.dataset.name); if (!r) return;
    setMethod(r, sel.value);
    sel.closest('tr').children[3].innerHTML = summaryCell(r);
    renderT1();
  });
  // Global buttons apply to ALL numeric variables (not just those matching the filter)
  document.querySelectorAll('[data-setall]').forEach(b => b.addEventListener('click', () => {
    state.results.forEach(r => setMethod(r, b.dataset.setall));
    renderTable();
    setStatus(b.dataset.setall === 'auto' ? 'All variables set to automatic selection.' : `All variables set to ${b.dataset.setall}.`);
  }));
  document.querySelectorAll('[data-cat-setall]').forEach(b => b.addEventListener('click', () => {
    state.categorical.forEach(r => { state.t1.catChoices[r.name] = b.dataset.catSetall; });
    renderCategoricalSettings(); renderTable(); renderT1();
  }));
  tbody.addEventListener('keydown', e => {
    if (e.target.closest('select')) return;
    if (e.key !== 'Enter' && e.key !== ' ') return;
    const tr = e.target.closest('tr[data-name]'); if (tr) { e.preventDefault(); showDetail(tr.dataset.name); }
  });
  $('copyTable').addEventListener('click', () => copyText(tableRows().map(r => r.map(tsvCell).join('\t')).join('\n'), 'Table copied (tab-separated).'));
  /* Copy for manuscript action temporarily disabled; retained for future use.
  $('copyMs').addEventListener('click', () => {
    const tx = x => window.AppI18n?.text(x) || x;
    const heading = tx('Overall summary (not a stratified Table 1); missing values excluded');
    copyText(heading + '\n' + getView().map(r => {
      const method = r.type === 'numeric' ? (r.choice === 'auto' ? `Auto: ${r.autoMethod}` : r.method) : ({ counts: 'Counts', percentage: 'Percentage', counts_percentage: 'Counts (percentage)' })[categoricalMode(r)];
      return `${r.name}, ${tx(method)}: ${summaryValue(r)}`;
    }).join('\n'), tx('Copied for manuscript.'));
  });
  */
  $('dlCsv').addEventListener('click', () => download(tableRows().map(r => r.map(csvCell).join(',')).join('\r\n'), baseName() + '.csv', 'text/csv'));
  $('dlTsv').addEventListener('click', () => download(tableRows().map(r => r.map(tsvCell).join('\t')).join('\r\n'), baseName() + '.tsv', 'text/tab-separated-values'));
}
window.onLocaleChange = () => {
  if (!state.headers.length) return;
  const detailName = $('detail').open ? $('dtitle').dataset.variable : '';
  /* renderPreview(); renderTypes(); */
  renderTable(); renderCategorical(); fillGroupSelect(); renderT1(); figRefresh();
  if (detailName) { $('detail').close(); showDetail(detailName); }
};
init(); initExtra();
