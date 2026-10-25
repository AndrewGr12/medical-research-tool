// Run with: node statistics-tests.cjs
// Extracts the same pure test functions used by main.js; no browser or dependencies required.
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');

const source = fs.readFileSync(path.join(__dirname, 'main.js'), 'utf8');
const start = source.indexOf('// ---------- Group comparison tests');
const end = source.indexOf('function buildTable1()', start);
assert(start >= 0 && end > start, 'Could not find the isolated statistical functions in main.js');
const context = {};
vm.runInNewContext(`${source.slice(start, end)}\nglobalThis.api = { selectComparisonTest, pText, fisher2x2 };`, context);
const { selectComparisonTest, pText, fisher2x2 } = context.api;
const near = (actual, expected, tolerance = 1e-6) => assert.ok(Math.abs(actual - expected) <= tolerance, `${actual} differs from ${expected}`);

// Welch's two-sample t-test: reference value for [1,2] versus [3,4].
const welch = selectComparisonTest('numeric', [[1, 2], [3, 4]], 'parametric');
assert.equal(welch.method, "Welch's t-test");
near(welch.p, 0.10557280900008414);

// Welch's ANOVA: equal n and variance in three separated groups gives F=12, df=(2,2).
const welchA = selectComparisonTest('numeric', [[1, 2], [3, 4], [5, 6]], 'parametric');
assert.equal(welchA.method, "Welch's ANOVA");
near(welchA.p, 1 / 13);

// Nonparametric paths and multi-group selection.
const mw = selectComparisonTest('numeric', [[1, 2, 3], [4, 5, 6]], 'nonparametric');
assert.equal(mw.method, 'Mann–Whitney U (exact permutation)'); near(mw.p, 0.1);
const kw = selectComparisonTest('numeric', [[1, 2], [3, 4], [5, 6]], 'nonparametric');
assert.equal(kw.method, 'Kruskal–Wallis (exact permutation)'); near(kw.p, 1 / 15);

// Sparse 2x2 table: standard two-sided Fisher exact reference example.
const fisherReference = fisher2x2([[1, 9], [11, 3]]);
assert.match(fisherReference.method, /Fisher's exact/); near(fisherReference.p, 0.0027594561852200836);
const fisherSelected = selectComparisonTest('categorical', [['A'], ['B']], 'parametric');
assert.match(fisherSelected.method, /Fisher's exact/); near(fisherSelected.p, 1);

// Adequate expected counts use Pearson chi-square; sparse larger tables use conditional exact.
const pearson = selectComparisonTest('categorical', [Array(10).fill('A').concat(Array(20).fill('B')), Array(20).fill('A').concat(Array(10).fill('B'))], 'parametric');
assert.equal(pearson.method, "Pearson's chi-square test"); near(pearson.p, 0.0098232745, 2e-7);
const ffh = selectComparisonTest('categorical', [['A'], ['B'], ['A', 'B']], 'parametric');
assert.match(ffh.method, /Fisher–Freeman–Halton/); near(ffh.p, 1);

// Missing values are excluded before calling the test; degenerate groups return NA with a reason.
const missingFiltered = selectComparisonTest('numeric', [[1, 2], [3, 4]], 'parametric');
assert.ok(missingFiltered.p !== null);
const insufficient = selectComparisonTest('numeric', [[1], [2, 3]], 'parametric');
assert.equal(insufficient.p, null); assert.ok(insufficient.reason.length > 0); assert.equal(pText(insufficient), 'NA');
const degenerate = selectComparisonTest('categorical', [[], ['A']], 'parametric');
assert.equal(degenerate.p, null); assert.ok(degenerate.reason.length > 0);

console.log('All grouped-comparison statistics checks passed.');
