const fs = require('fs');
const path = require('path');

// Minimal KMP & BMH copies for unit tests
function kmpSearch(text, pattern) {
  let comparisons = 0;
  if (typeof text !== 'string' || typeof pattern !== 'string') return { found: false, index: -1, comparisons };
  const t = text.toLowerCase();
  const p = pattern.toLowerCase();
  const n = t.length; const m = p.length;
  if (m === 0) return { found: true, index: 0, comparisons };

  const lps = new Array(m).fill(0);
  let len = 0; let i = 1;
  while (i < m) {
    comparisons++;
    if (p[i] === p[len]) { len++; lps[i] = len; i++; }
    else { if (len !== 0) { len = lps[len - 1]; } else { lps[i] = 0; i++; } }
  }

  i = 0; let j = 0;
  while (i < n) {
    comparisons++;
    if (t[i] === p[j]) { i++; j++; }
    if (j === m) return { found: true, index: i - j, comparisons };
    else if (i < n && t[i] !== p[j]) { if (j !== 0) j = lps[j - 1]; else i++; }
  }
  return { found: false, index: -1, comparisons };
}

// Full Boyer-Moore for tests
function buildBadCharTable(p) {
  const m = p.length;
  const table = {};
  for (let i = 0; i < m; i++) table[p[i]] = i;
  return table;
}

function preprocessGoodSuffix(p) {
  const m = p.length;
  const suff = new Array(m + 1).fill(0);
  suff[m] = m + 1;
  let g = m, f = m + 1;
  suff[m - 1] = m;
  g = m - 1;
  f = m;
  for (let i = m - 2; i >= 0; --i) {
    if (i > g && suff[i + m - 1 - f] < i - g) {
      suff[i] = suff[i + m - 1 - f];
    } else {
      if (i < g) g = i;
      f = i;
      while (g >= 0 && p[g] === p[g + m - 1 - f]) g--;
      suff[i] = f - g;
    }
  }
  const shift = new Array(m + 1).fill(m);
  let j = 0;
  for (let i = m - 1; i >= 0; --i) {
    if (suff[i] === i + 1) {
      for (; j < m - 1 - i; ++j) {
        if (shift[j] === m) shift[j] = m - 1 - i;
      }
    }
  }
  for (let i = 0; i <= m - 2; ++i) {
    shift[m - 1 - suff[i]] = m - 1 - i;
  }
  return shift;
}

function fullBmSearch(text, pattern) {
  let comparisons = 0;
  if (typeof text !== 'string' || typeof pattern !== 'string') return { found: false, index: -1, comparisons };
  const t = text.toLowerCase();
  const p = pattern.toLowerCase();
  const n = t.length; const m = p.length;
  if (m === 0) return { found: true, index: 0, comparisons };
  const bad = buildBadCharTable(p);
  const good = preprocessGoodSuffix(p);
  let i = 0;
  while (i <= n - m) {
    let j = m - 1;
    while (j >= 0) {
      comparisons++;
      if (p[j] === t[i + j]) j--; else break;
    }
    if (j < 0) return { found: true, index: i, comparisons };
    const bcIdx = bad[t[i + j]] !== undefined ? bad[t[i + j]] : -1;
    const badShift = j - bcIdx;
    const goodShift = good[j];
    i += Math.max(1, Math.max(badShift, goodShift));
  }
  return { found: false, index: -1, comparisons };
}

function bmhSearch(text, pattern) {
  let comparisons = 0;
  if (typeof text !== 'string' || typeof pattern !== 'string') return { found: false, index: -1, comparisons };
  const t = text.toLowerCase();
  const p = pattern.toLowerCase();
  const n = t.length; const m = p.length;
  if (m === 0) return { found: true, index: 0, comparisons };
  const badChar = {};
  for (let k = 0; k < m - 1; k++) badChar[p[k]] = m - 1 - k;
  let i = 0;
  while (i <= n - m) {
    let j = m - 1;
    while (j >= 0) { comparisons++; if (p[j] === t[i + j]) j--; else break; }
    if (j < 0) return { found: true, index: i, comparisons };
    const shiftChar = t[i + m - 1];
    const shift = badChar[shiftChar] !== undefined ? badChar[shiftChar] : m;
    i += shift;
  }
  return { found: false, index: -1, comparisons };
}

function assert(condition, msg) {
  if (!condition) {
    console.error('FAIL:', msg);
    process.exit(1);
  }
}

// Basic tests
console.log('Running tests...');

// 1. Empty pattern
let r = kmpSearch('ABC', '');
assert(r.found === true, 'KMP should return found for empty pattern');
r = bmhSearch('ABC', '');
assert(r.found === true, 'BMH should return found for empty pattern');

// 2. Pattern longer than text
r = kmpSearch('AB', 'ABC');
assert(r.found === false, 'KMP pattern longer than text');
r = bmhSearch('AB', 'ABC');
assert(r.found === false, 'BMH pattern longer than text');

// 3. Exact match
r = kmpSearch('GILANG RAMADHAN', 'GILANG');
assert(r.found === true && r.index === 0, 'KMP exact match at start');
r = bmhSearch('GILANG RAMADHAN', 'GILANG');
assert(r.found === true && r.index === 0, 'BMH exact match at start');

// 4. Case insensitivity
r = kmpSearch('Andri Rustandi', 'andri');
assert(r.found === true, 'KMP case-insensitive');
r = bmhSearch('Andri Rustandi', 'ANDRI');
assert(r.found === true, 'BMH case-insensitive');

// 5. Not found
r = kmpSearch('SITI AMINAH', 'UNKNOWN');
assert(r.found === false, 'KMP not found test');
r = bmhSearch('SITI AMINAH', 'UNKNOWN');
assert(r.found === false, 'BMH not found test');

// 6. Full Boyer-Moore parity with BMH and KMP on simple cases
r = fullBmSearch('GILANG RAMADHAN', 'GILANG');
assert(r.found === true && r.index === 0, 'FullBM exact match at start');
r = fullBmSearch('Andri Rustandi', 'andri');
assert(r.found === true, 'FullBM case-insensitive');
r = fullBmSearch('AB', 'ABC');
assert(r.found === false, 'FullBM pattern longer than text');

console.log('All tests passed.');
