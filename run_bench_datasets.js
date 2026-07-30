const { execSync } = require('child_process');
const path = require('path');
const fs = require('fs');
const gen = require('./generate_datasets');

const sizes = [1000, 10000, 100000];
const outDir = path.join(__dirname, 'data');
if (!fs.existsSync(outDir)) fs.mkdirSync(outDir, { recursive: true });

// Generate datasets if missing
for (const s of sizes) {
  const p = path.join(outDir, `dataset_warga_${s}.json`);
  if (!fs.existsSync(p)) {
    console.log('Generating', s);
    gen.generateDataset(s, p);
  } else {
    console.log('Dataset exists:', p);
  }
}

// Run benchmark for each dataset with tuned iterations to keep runtime reasonable
const config = {
  1000: { iterations: 100, warmup: 10 },
  10000: { iterations: 30, warmup: 5 },
  100000: { iterations: 5, warmup: 2 }
};

for (const s of sizes) {
  const datasetPath = path.join(outDir, `dataset_warga_${s}.json`);
  const it = config[s].iterations;
  const w = config[s].warmup;
  console.log(`\n=== Running benchmark for ${s} records (iterations=${it}, warmup=${w}) ===`);
  // Run benchmark.js with env overrides
  const env = Object.assign({}, process.env, {
    DATASET_PATH: datasetPath,
    ITERATIONS: String(it),
    WARMUP: String(w)
  });
  try {
    execSync('node benchmark.js', { stdio: 'inherit', env });
    // move output CSV
    const src = path.join(__dirname, 'benchmark_results.csv');
    const dest = path.join(__dirname, `benchmark_results_${s}.csv`);
    if (fs.existsSync(src)) fs.renameSync(src, dest);
    console.log('Saved', dest);
  } catch (err) {
    console.error('Benchmark failed for', s, err);
  }
}

console.log('\nAll benchmarks finished. CSVs are in repository root.');
