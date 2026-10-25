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

const docxStart = js.indexOf('const crcTable = (() =>');
const docxEnd = js.indexOf('const DOCX_MIME', docxStart);
assert.ok(docxStart >= 0 && docxEnd > docxStart, 'Word export builder is present');
const docxContext = { TextEncoder, Uint8Array, Uint32Array };
vm.runInNewContext(`${js.slice(docxStart, docxEnd)}\nglobalThis.testDocx = docxBytes('Table 1', ['Variable','N'], [['Age','42']], []);`, docxContext);
const docxBytes = Buffer.from(docxContext.testDocx).toString('utf8');
for (const edge of ['top','left','bottom','right','insideH','insideV']) assert.ok(docxBytes.includes(`<w:${edge} w:val="single"`), `Word export includes ${edge} grid borders`);
assert.ok((docxBytes.match(/<w:tc>/g) || []).length >= 4, 'Word export contains individual table cells');

console.log('All UX package smoke checks passed.');
