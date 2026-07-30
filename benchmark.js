const fs = require('fs');
const path = require('path');

// Load dataset (can be overridden with env var DATASET_PATH)
const datasetPath = process.env.DATASET_PATH ? path.resolve(process.env.DATASET_PATH) : path.join(__dirname, 'data', 'dataset_warga.json');
if (!fs.existsSync(datasetPath)) {
  console.error('Dataset not found:', datasetPath);
  process.exit(1);
}
const DATASET_WARGA = JSON.parse(fs.readFileSync(datasetPath, 'utf8'));

// KMP implementation (case-insensitive)
function kmpSearch(text, pattern) {
  let comparisons = 0;
  if (typeof text !== 'string' || typeof pattern !== 'string') return { found: false, index: -1, comparisons };
  const t = text.toLowerCase();
  const p = pattern.toLowerCase();
  const n = t.length;
  const m = p.length;
  if (m === 0) return { found: true, index: 0, comparisons };

  const lps = new Array(m).fill(0);
  let len = 0;
  let i = 1;
  while (i < m) {
    comparisons++;
    if (p[i] === p[len]) {
      len++;
      lps[i] = len;
      i++;
    } else {
      if (len !== 0) {
        len = lps[len - 1];
      } else {
        lps[i] = 0;
        i++;
      }
    }
  }

  i = 0; // index for text
  let j = 0; // index for pattern
  while (i < n) {
    comparisons++;
    if (t[i] === p[j]) {
      i++; j++;
    }
    if (j === m) {
      return { found: true, index: i - j, comparisons };
    } else if (i < n && t[i] !== p[j]) {
      if (j !== 0) {
        j = lps[j - 1];
      } else {
        i++;
      }
    }
  }

  return { found: false, index: -1, comparisons };
}

// Boyer-Moore-Horspool (BMH)
function bmhSearch(text, pattern) {
  let comparisons = 0;
  if (typeof text !== 'string' || typeof pattern !== 'string') return { found: false, index: -1, comparisons };
  const t = text.toLowerCase();
  const p = pattern.toLowerCase();
  const n = t.length;
  const m = p.length;
  if (m === 0) return { found: true, index: 0, comparisons };

  const badChar = {};
  for (let k = 0; k < m - 1; k++) {
    badChar[p[k]] = m - 1 - k;
  }

  let i = 0;
  while (i <= n - m) {
    let j = m - 1;
    while (j >= 0) {
      comparisons++;
      if (p[j] === t[i + j]) {
        j--;
      } else {
        break;
      }
    }
    if (j < 0) {
      return { found: true, index: i, comparisons };
    }
    const shiftChar = t[i + m - 1];
    const shift = badChar[shiftChar] !== undefined ? badChar[shiftChar] : m;
    i += shift;
  }
  return { found: false, index: -1, comparisons };
}

// Full Boyer-Moore (bad-character + good-suffix)
function buildBadCharTable(p) {
  const m = p.length;
  const table = {};
  for (let i = 0; i < m; i++) {
    table[p[i]] = i;
  }
  return table;
}

function preprocessGoodSuffix(p) {
  const m = p.length;
  const suff = new Array(m + 1).fill(0);
  suff[m] = m + 1;
  let g = m, f = m + 1;
  suff[m - 1] = m;
  // compute suffixes (using standard BM preprocessing)
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
  const n = t.length;
  const m = p.length;
  if (m === 0) return { found: true, index: 0, comparisons };

  // Bad character: store last occurrence
  const bad = buildBadCharTable(p);
  const good = preprocessGoodSuffix(p);

  let i = 0;
  while (i <= n - m) {
    let j = m - 1;
    while (j >= 0) {
      comparisons++;
      if (p[j] === t[i + j]) j--; else break;
    }
    if (j < 0) {
      return { found: true, index: i, comparisons };
    }
    const bcIdx = bad[t[i + j]] !== undefined ? bad[t[i + j]] : -1;
    const badShift = j - bcIdx;
    const goodShift = good[j];
    i += Math.max(1, Math.max(badShift, goodShift));
  }
  return { found: false, index: -1, comparisons };
}

// Helper: high-resolution time in milliseconds
function nowMs() {
  return Number(process.hrtime.bigint() / 1000000n);
}

// Benchmark runner
function runBenchmark({ patterns, iterations, warmup }) {
  // Allow override via env vars
  iterations = iterations || (process.env.ITERATIONS ? Number(process.env.ITERATIONS) : 50);
  warmup = warmup || (process.env.WARMUP ? Number(process.env.WARMUP) : 5);
  const results = [];

  for (const pattern of patterns) {
    // Warmup
    for (let w = 0; w < warmup; w++) {
      for (const warga of DATASET_WARGA) {
        kmpSearch(warga.nama, pattern);
        bmhSearch(warga.nama, pattern);
      }
    }

    // KMP measurements
    let kmpTotalTime = 0;
    let kmpTotalComparisons = 0;
    for (let it = 0; it < iterations; it++) {
      const start = nowMs();
      for (const warga of DATASET_WARGA) {
        const r = kmpSearch(warga.nama, pattern);
        kmpTotalComparisons += r.comparisons;
      }
      const end = nowMs();
      kmpTotalTime += (end - start);
    }

    // BM measurements (Horspool)
    let bmTotalTime = 0;
    let bmTotalComparisons = 0;
    for (let it = 0; it < iterations; it++) {
      const start = nowMs();
      for (const warga of DATASET_WARGA) {
        const r = bmhSearch(warga.nama, pattern);
        bmTotalComparisons += r.comparisons;
      }
      const end = nowMs();
      bmTotalTime += (end - start);
    }

    // Full Boyer-Moore measurements
    let fullBmTotalTime = 0;
    let fullBmTotalComparisons = 0;
    for (let it = 0; it < iterations; it++) {
      const start = nowMs();
      for (const warga of DATASET_WARGA) {
        const r = fullBmSearch(warga.nama, pattern);
        fullBmTotalComparisons += r.comparisons;
      }
      const end = nowMs();
      fullBmTotalTime += (end - start);
    }

    const row = {
      pattern,
      kmp_avg_time_ms: (kmpTotalTime / iterations).toFixed(4),
      kmp_total_comparisons: kmpTotalComparisons,
      bm_avg_time_ms: (bmTotalTime / iterations).toFixed(4),
      bm_total_comparisons: bmTotalComparisons,
      fullbm_avg_time_ms: (fullBmTotalTime / iterations).toFixed(4),
      fullbm_total_comparisons: fullBmTotalComparisons
    };

    console.log('Bench:', row);
    results.push(row);
  }

  // Write CSV
  const outPath = path.join(__dirname, 'benchmark_results.csv');
  const header = 'pattern,kmp_avg_time_ms,kmp_total_comparisons,bm_avg_time_ms,bm_total_comparisons,fullbm_avg_time_ms,fullbm_total_comparisons\n';
  const lines = results.map(r => `${r.pattern},${r.kmp_avg_time_ms},${r.kmp_total_comparisons},${r.bm_avg_time_ms},${r.bm_total_comparisons},${r.fullbm_avg_time_ms},${r.fullbm_total_comparisons}`).join('\n');
  fs.writeFileSync(outPath, header + lines, 'utf8');
  console.log('Saved benchmark_results.csv');
}

// Example patterns to test: exact matches, partials, non-existent
const patterns = [
  'ANDRI',
  'SITI',
  'GILANG',
  '3273011210010002',
  'NONEXISTENTNAME'
];

runBenchmark({ patterns, iterations: 100, warmup: 10 });
