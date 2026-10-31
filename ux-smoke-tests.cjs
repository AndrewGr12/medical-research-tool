// Structural smoke checks for this standalone static website copy.
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');

const root = __dirname;
let html = fs.readFileSync(path.join(root, 'index.html'), 'utf8');
const rawHtml = html;
const css = fs.readFileSync(path.join(root, 'style.css'), 'utf8');
const js = fs.readFileSync(path.join(root, 'main.js'), 'utf8');
const locales = fs.readFileSync(path.join(root, 'locales.js'), 'utf8');
const parser = js.match(/^function parseCSV\(text\) \{[\s\S]*?^\}/m)?.[0];
const sampleGenerator = js.match(/^function sampleCSV\(\) \{[\s\S]*?^\}/m)?.[0];
const uniqueNameHelper = js.match(/^function uniqueNames\(headers\) \{[\s\S]*?^\}/m)?.[0];
const headerWarningHelper = js.match(/^function headerNameWarning\(headers\) \{[\s\S]*?^\}/m)?.[0];
assert.ok(parser && sampleGenerator && uniqueNameHelper && headerWarningHelper, 'CSV parser, sample generator, and header helpers are present');
const dataApi = vm.runInNewContext(`${parser}\n${sampleGenerator}\n${uniqueNameHelper}\n${headerWarningHelper}\n({ parseCSV, sampleCSV, uniqueNames, headerNameWarning })`);
const sample = dataApi.parseCSV(dataApi.sampleCSV());
assert.equal(sample.rows.length, 241, 'sample includes one header and 240 data rows');
assert.equal(sample.rows[0].length, 14, 'sample includes 14 clinical variables');
assert.ok(['Sex', 'Diagnosis', 'Disease_activity', 'Smoking_status', 'Diabetes', 'Hypertension', 'Treatment'].every(name => sample.rows[0].includes(name)), 'sample includes representative clinical categorical variables');
assert.equal(sample.unterminatedQuote, false, 'generated sample is valid quoted CSV');
assert.equal(dataApi.headerNameWarning(['Age','Sex']), '', 'clean unique headers do not trigger a warning');
assert.match(dataApi.headerNameWarning(['Age','Age','']), /blank or duplicated/, 'blank and duplicate headers prompt a clear review warning');
assert.deepEqual(JSON.parse(JSON.stringify(dataApi.uniqueNames(['Age','Age','']))), ['Age','Age (2)','Column 3'], 'renamed headers remain unique and stable');
const parsedQuoted = dataApi.parseCSV('\uFEFFA,B\r\n"a,b","line 1\nline 2"\r\n');
assert.deepEqual(JSON.parse(JSON.stringify(parsedQuoted.rows)), [['A','B'], ['a,b','line 1\nline 2']], 'CSV parser handles BOM, quoted delimiters, and quoted newlines');
html = html.replace(/<!--[\s\S]*?-->/g, '');

assert.match(html, /<main\b/i, 'main landmark is present');
assert.match(html, /<h1\b/i, 'page has a primary heading');
assert.match(html, /id="file"[^>]*accept="[^"]*\.csv[^"]*\.xlsx/i, 'CSV and Excel upload is available');
assert.match(html, /id="sample"/, 'sample-data action is present');
const uploadConsentHtml = html.match(/<dialog id="uploadConsentPrompt"[\s\S]*?<\/dialog>/)?.[0] || '';
assert.ok(uploadConsentHtml, 'file upload confirmation dialog exists');
assert.match(uploadConsentHtml, /aria-labelledby="uploadConsentTitle"[^>]*aria-describedby="uploadConsentDescription"/, 'upload confirmation has an accessible name and description');
assert.match(uploadConsentHtml, /id="uploadConsentConfirm"/, 'upload confirmation includes an explicit affirmative action');
assert.match(uploadConsentHtml, /id="uploadConsentCancel"/, 'upload confirmation includes a cancel action');
assert.match(js, /addEventListener\('change', e => requestUploadConsent\(e\.target\.files\[0\]\)\)/, 'file picker asks for confirmation before processing');
assert.match(js, /drop\.addEventListener\('drop', e => requestUploadConsent\(e\.dataTransfer\.files\[0\]\)\)/, 'drag-and-drop asks for confirmation before processing');
assert.match(js, /uploadConsent\.addEventListener\('close',[\s\S]*?uploadConsentGranted\(uploadConsent, file\)[\s\S]*?handleFile\(file\)/, 'file processing starts only after affirmative consent');
assert.match(js, /dialog\.showModal\(\)/, 'confirmation is presented in a native modal dialog');
assert.match(css, /#uploadConsentPrompt\[open\]\{animation:fadeIn/, 'upload confirmation uses the existing fade-in effect');
assert.match(html, /<button[^>]*id="patientClose"[^>]*aria-label="Close"[^>]*title="Close"/, 'patient ID dialog has a labeled close control');
assert.match(js, /patientClose'\)\.addEventListener\('click', \(\) => \$\('patientPrompt'\)\.close\('cancel'\)\)/, 'close control cancels patient ID selection');
assert.match(js, /if \(patientId === null\) \{[\s\S]*?\$\('file'\)\.value = '';[\s\S]*?return;[\s\S]*?\}/, 'canceling patient ID selection clears the file and returns before analysis');
const patientChoiceHelper = js.match(/^function patientIdChoice\([\s\S]*?^\}/m)?.[0];
assert.ok(patientChoiceHelper, 'patient ID dialog choices are isolated for behavior checks');
const patientChoiceApi = vm.runInNewContext(`${patientChoiceHelper}\n({ patientIdChoice })`);
assert.equal(patientChoiceApi.patientIdChoice('yes', 'record_id'), 'record_id', 'affirmative patient ID choice excludes the selected ID');
assert.equal(patientChoiceApi.patientIdChoice('no', 'record_id'), '', 'No patient ID continues analysis without exclusion');
assert.equal(patientChoiceApi.patientIdChoice('cancel', 'record_id'), null, 'closing the patient ID dialog cancels the upload flow');
const identifierFilterSource = js.match(/^function filterIdentifierNotesAfterNoPatientId\([\s\S]*?^\}/m)?.[0];
assert.ok(identifierFilterSource, 'identifier-note filtering is isolated for behavior checks');
const identifierFilterApi = vm.runInNewContext(`${identifierFilterSource}\n({ filterIdentifierNotesAfterNoPatientId })`);
const idNotices = [
  '“patient_id” looks like a possible identifier because most non-missing values are unique; review its variable type or exclude it if appropriate.',
  '“patient_id” has all-distinct values (possible identifier).',
  '“Age” has 20% missing values.'
];
assert.deepEqual(JSON.parse(JSON.stringify(identifierFilterApi.filterIdentifierNotesAfterNoPatientId(idNotices, true))), ['“Age” has 20% missing values.'], 'No patient ID suppresses only the two possible-identifier notes');
assert.deepEqual(JSON.parse(JSON.stringify(identifierFilterApi.filterIdentifierNotesAfterNoPatientId(idNotices, false))), idNotices, 'other patient-ID choices leave all notes unchanged');
assert.match(html, /id="t1"/, 'Table 1 is present');
assert.match(rawHtml, /<!-- Table 1 Copy action temporarily disabled; retained for future use\.[\s\S]*?id="t1copy"[\s\S]*?-->/, 'Table 1 Copy button remains commented out');
assert.doesNotMatch(html, /<button[^>]*id="t1copy"/, 'Table 1 Copy button is not active');
assert.match(html, /class="btn primary" id="t1csv"/, 'Table 1 CSV download has primary styling');
assert.match(html, /class="btn primary" id="t1docx"/, 'Table 1 Word download has primary styling');
assert.match(html, /id="figtypes"/, 'figure workflow is present');
assert.match(html, /id="warnings"[^>]*role="status"[^>]*aria-live="polite"/, 'dynamic data warnings are announced politely');
assert.match(js, /addEventListener\('click',\s*\(\)\s*=>\s*loadText\(sampleCSV\(\)/, 'sample-data action has a handler');
assert.match(js, /addEventListener\('click',\s*e\s*=>\s*\{[\s\S]*details\.menu\[open\]/, 'open menus close after outside clicks');
assert.match(locales, /Transform your clinical data into tables and visualizations!/,'localized post-selection description source is present');
assert.match(locales, /Some column names were blank or duplicated[\s\S]*?noms de colonnes étaient vides[\s\S]*?nombres de columna estaban vacíos[\s\S]*?названия столбцов были пустыми/, 'header-review warning is translated in all supported languages');
assert.match(js, /Table 1 Copy action temporarily disabled; retained for future use\.[\s\S]*?t1copy[\s\S]*?\*\//, 'Table 1 Copy logic remains commented out');
for (const locale of ['en', 'fr', 'es', 'ru']) assert.match(html, new RegExp(`data-locale="${locale}"`), `${locale} language button is present`);

const ids = [...html.matchAll(/\bid="([^"]+)"/g)].map(m => m[1]);
const idSet = new Set(ids);
assert.equal(idSet.size, ids.length, 'HTML IDs are unique');
for (const [, target] of html.matchAll(/<label\b[^>]*\bfor="([^"]+)"/gi)) assert.ok(idSet.has(target), `label target #${target} exists`);
for (const [, target] of html.matchAll(/\baria-labelledby="([^"]+)"/gi)) for (const id of target.split(/\s+/)) assert.ok(idSet.has(id), `aria-labelledby target #${id} exists`);
for (const [, target] of html.matchAll(/\baria-describedby="([^"]+)"/gi)) for (const id of target.split(/\s+/)) assert.ok(idSet.has(id), `aria-describedby target #${id} exists`);

for (const [, attr, ref] of html.matchAll(/<(?:script|link|img)\b[^>]*?\b(src|href)="([^"]+)"[^>]*>/gi)) {
  if (/^(?:https?:|data:|#)/i.test(ref)) continue;
  assert.ok(fs.existsSync(path.join(root, decodeURIComponent(ref))), `local resource ${ref} exists`);
}
assert.match(html, /class="tablewrap"[^>]*role="region"[^>]*tabindex="0"/,'summary table has a keyboard-scrollable named region');
assert.match(css, /@media\s*\(max-width:\s*600px\)/, 'mobile layout rules are present');
assert.match(css, /#home\.compact \.paths\{width:100%/, 'compact Tables/Figures switcher keeps the full card width');
assert.match(css, /#home\.compact \.path\{[^}]*min-height:122px/, 'compact Tables/Figures cards reduce height while preserving content');
assert.match(css, /#home\.compact \.path span\{display:block/, 'compact workspace cards keep their descriptive text on wider layouts');
assert.match(css, /\.menupanel\{position:absolute;left:0;right:auto/, 'dropdown panels align their left edge with their trigger');
assert.match(css, /#exportMenu>\.menupanel\{left:auto;right:0\}/, 'Export menu aligns its right edge with its trigger');
assert.match(css, /prefers-reduced-motion:\s*reduce/, 'reduced-motion preference is respected');
assert.match(css, /\.note\{font-size:\.875rem/, 'helper text uses a readable 14px minimum');
assert.match(css, /\.figtype span\{color:var\(--mute\);font-size:1\.28rem/, 'figure descriptions use the current readable card size');

const blockBetween = (source, startMarker, endMarker) => {
  const start = source.indexOf(startMarker);
  const end = source.indexOf(endMarker, start + startMarker.length);
  assert.ok(start >= 0 && end > start, `source block ${startMarker} exists`);
  return source.slice(start, end);
};
const outputBlocks = [
  blockBetween(js, 'function buildSummaryRow(', '\nfunction summaryCell('),
  blockBetween(js, 'function renderTable(', '\n// Make the scope'),
  blockBetween(js, 'const tableRows =', '\nconst csvCell'),
  blockBetween(js, 'function showDetail(', '\nfunction drawBoxPlot('),
  blockBetween(js, 'function buildTable1()', '\nfunction renderT1('),
  blockBetween(js, 'const t1Matrix =', '\nfunction initExtra('),
  blockBetween(js, 'function summaryDocx()', '\nfunction table1Docx('),
  blockBetween(js, 'function table1Docx()', '\nfunction buildPdf('),
  blockBetween(js, 'const tl =', '\nfunction figRefresh('),
  blockBetween(js, 'function figDraw()', '\nconst figName')
];
for (const [i, block] of outputBlocks.entries()) {
  assert.doesNotMatch(block, /window\.AppI18n|\.locale\b|\.language\b/, `output builder ${i + 1} is independent of UI language`);
}
assert.match(js, /const OUT_LOCALE = 'en-US'/, 'output locale is fixed in one shared constant');
assert.doesNotMatch(outputBlocks.join('\n'), /Intl\.NumberFormat\((?!OUT_LOCALE)|toLocaleString\((?!OUT_LOCALE)/, 'output number formatting uses OUT_LOCALE');
const chromeCountBlock = blockBetween(js, 'function renderAll(', '\n/* Data preview renderer');
assert.match(chromeCountBlock, /toLocaleString\(window\.AppI18n\?\.locale/, 'localized number formatting remains limited to UI chrome counters');
assert.doesNotMatch(js.replace(chromeCountBlock, ''), /toLocaleString\(window\.AppI18n\?\.locale/, 'output number formatting does not depend on AppI18n locale');
assert.doesNotMatch(js, /new Intl\.NumberFormat\(window\.AppI18n/, 'Intl output formatting does not read AppI18n');
assert.match(js, /localeCompare\([^,]+, 'en', \{ numeric: true \}\)/, 'output sorting uses fixed English collation');

const outputIds = ['stats', 't1', 't1notes', 'types', 'preview', 'figplot', 'figdesc', 'dbody'];
for (const id of outputIds) assert.match(rawHtml, new RegExp(`id="${id}"[^>]*data-output|data-output[^>]*id="${id}"`), `#${id} is marked as output`);
assert.match(rawHtml, /<caption[^>]*data-output/, 'output captions are marked');
assert.match(rawHtml, /class="tablewrap" data-output[^>]*aria-label="Scrollable summary statistics table"/, 'output table aria label is protected');
assert.match(locales, /acceptNode\(node\) \{ return preserve\(node\.parentElement\) \? NodeFilter\.FILTER_REJECT : NodeFilter\.FILTER_ACCEPT; \}/, 'translatePage skips output subtrees');
assert.match(locales, /new MutationObserver\(\(\) => translatePage\(\)\)\.observe\(document\.body/, 'mutation observer uses the protected translatePage path');
assert.match(locales, /if \(current === 'ar'\) el\.setAttribute\('dir', 'ltr'\)/, 'output containers stay left-to-right for Arabic');
assert.match(locales, /const languages = \['en', 'fr', 'es', 'ru', 'ar'\]/, 'all five UI languages remain configured');
assert.match(locales, /for \(const row of rows\) if \(!outputOnlyKeys\.has\(row\[0\]\)\) for \(let i = 0; i < languages\.length; i\+\+\)/, 'each retained UI dictionary row receives a value for every language');
const dynamicBlock = blockBetween(locales, 'function dynamic(en, lang)', '\n  function text(');
assert.doesNotMatch(dynamicBlock, /Missingness for|Scatter plot:|Fisher–Freeman–Halton exact test \(|Histogram of|Box plot of|Bar chart of/, 'output-only regex translations are removed');
assert.match(locales, /for \(const key of outputOnlyKeys\) arabic\.delete\(key\)/, 'output-only dictionary entries cannot re-enter through Arabic translations');

const runtimeSource = js.replace(/\ninit\(\); initExtra\(\);\s*$/, '\n');
assert.notEqual(runtimeSource, js, 'runtime builders can be loaded without attaching browser event handlers');
const outputContext = vm.createContext({ TextEncoder, window: { AppI18n: null }, document: {} });
vm.runInContext(runtimeSource, outputContext, { filename: 'main.js' });
const localeSnapshots = vm.runInContext(`(() => {
  const parsed = parseCSV(sampleCSV());
  state.headers = parsed.rows[0]; state.rows = parsed.rows.slice(1); state.fileName = 'sample_data.csv';
  const warnings = [];
  state.cols = state.headers.map((name, i) => analyzeColumn(name, i, state.rows, warnings, null));
  state.results = state.cols.filter(c => c.type === 'numeric');
  state.categorical = state.cols.filter(c => c.type === 'categorical');
  state.excluded = state.cols.filter(c => c.type === 'excluded');
  state.t1.group = state.categorical.find(c => c.name === 'Sex')?.name || '';
  state.t1.selected = new Set([state.results[0].name, 'Diabetes', 'Treatment']);
  state.t1.skip = new Set(); state.t1.overall = true; state.t1.showPValues = false;
  state.sortKey = null; state.sortDir = 1; state.patientId = '';
  const localeMap = { en: 'en-US', fr: 'fr-FR', es: 'es-ES', ru: 'ru-RU', ar: 'ar-EG' };
  return ['en','fr','es','ru','ar'].map(language => {
    window.AppI18n = { language, locale: localeMap[language], text: value => 'localized:' + value, markup: value => 'localized:' + value };
    const table1 = buildTable1(), summary = tableRows(), matrix = t1Matrix();
    const figures = [
      figHist({ x: state.results[0].name, bins: 0 }),
      figBox({ x: state.results[0].name, g: state.t1.group }),
      figBar({ c: state.t1.group, mode: 'count' }), figMiss(),
      figScatter({ x: state.results[0].name, y: state.results[1].name, g: state.t1.group })
    ];
    const csv = rows => rows.map(row => row.map(csvCell).join(',')).join('\\r\\n');
    return JSON.stringify({ summary, table1, matrix, figures, summaryCsv: csv(summary), table1Csv: csv(matrix),
      summaryTsv: summary.map(row => row.map(tsvCell).join('\\t')).join('\\r\\n'),
      summaryDocx: Array.from(summaryDocx()), table1Docx: Array.from(table1Docx()) });
  });
})()`, outputContext);
assert.equal(new Set(localeSnapshots).size, 1, 'table rows, Table 1 notes, all figure SVG/descriptions, and exports are byte-identical in all five languages');
const localeContext = vm.createContext({ window: {}, document: { addEventListener() {} } });
vm.runInContext(locales, localeContext, { filename: 'locales.js' });
for (const [language, expected] of Object.entries({ en: 'Upload', fr: 'Importation', es: 'Carga de datos', ru: 'Загрузка данных', ar: 'رفع ملف' })) {
  assert.equal(localeContext.window.AppI18n.text('Upload', language), expected, `${language} retains an Upload translation`);
}
const uploadConsentStrings = [
  'Confirm data de-identification',
  'Before this file is processed, confirm that it contains only de-identified information and no patient names, dates of birth, medical record numbers, contact information, or other direct identifiers. You are responsible for ensuring that you are authorized to use and upload this data.',
  'No, cancel upload', 'Yes, I confirm'
];
for (const phrase of uploadConsentStrings) for (const language of ['fr', 'es', 'ru', 'ar']) {
  const translation = localeContext.window.AppI18n.text(phrase, language);
  assert.ok(translation && translation !== phrase, `${language} translates upload-consent string: ${phrase}`);
  if (language === 'ar') assert.match(translation, /[\u0600-\u06FF]/, `Arabic upload-consent string uses Arabic script: ${phrase}`);
}
for (const [language, pattern] of Object.entries({ fr: /dé-identifiées/, es: /desidentificados/, ru: /обезличенные/, ar: /منزوعة الهوية/ })) {
  const text = uploadConsentStrings[1];
  assert.match(localeContext.window.AppI18n.text(text, language), pattern, `${language} retains the de-identified data requirement`);
}
const uploadConsentSource = js.match(/^function uploadConsentGranted\([\s\S]*?^\}/m)?.[0];
assert.ok(uploadConsentSource, 'upload consent decision can be behavior-tested');
const uploadConsentGate = vm.runInNewContext(`${uploadConsentSource}\n({ uploadConsentGranted })`);
const pickedFile = { name: 'research.csv' };
assert.equal(uploadConsentGate.uploadConsentGranted({ returnValue: 'confirm' }, pickedFile), true, 'affirmative consent allows file processing');
assert.equal(uploadConsentGate.uploadConsentGranted({ returnValue: 'cancel' }, pickedFile), false, 'declining or closing the modal does not process the file');
assert.equal(uploadConsentGate.uploadConsentGranted({ returnValue: 'confirm' }, null), false, 'empty file selections do not start processing');
assert.equal(localeContext.window.AppI18n.text('Missing data, n (%)', 'fr'), 'Missing data, n (%)', 'output dictionary rows are no longer translated');

const docxStart = js.indexOf('const crcTable = (() =>');
const docxEnd = js.indexOf('const DOCX_MIME', docxStart);
assert.ok(docxStart >= 0 && docxEnd > docxStart, 'Word export builder is present');
const docxContext = { TextEncoder, Uint8Array, Uint32Array };
vm.runInNewContext(`${js.slice(docxStart, docxEnd)}\nglobalThis.testDocx = docxBytes('Table 1', ['Variable','N'], [['Age','42']], []);`, docxContext);
const docxBytes = Buffer.from(docxContext.testDocx).toString('utf8');
for (const edge of ['top','left','bottom','right','insideH','insideV']) assert.ok(docxBytes.includes(`<w:${edge} w:val="single"`), `Word export includes ${edge} grid borders`);
assert.ok((docxBytes.match(/<w:tc>/g) || []).length >= 4, 'Word export contains individual table cells');

console.log('All UX package smoke checks passed.');
