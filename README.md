CIPABOT — Benchmark & Tests

This folder adds reproducible benchmarking and basic unit tests for the string matching algorithms used in the CIPABOT project.

Files added:
- `benchmark.js` — Node script to benchmark KMP and BMH on `data/dataset_warga.json`. Outputs `benchmark_results.csv`.
- `test.js` — Minimal unit tests for `kmpSearch` and `bmhSearch`.
- `data/dataset_warga.json` — Sample dataset extracted from `app.js` for reproducible runs.

Requirements:
- Node.js (v14+ recommended)

Run unit tests:

```bash
node test.js
```

Run benchmark (may take a few minutes depending on iterations):

```bash
node benchmark.js
```

Notes:
- `benchmark.js` uses `process.hrtime.bigint()` for high-resolution timing and performs warmup iterations before measurements.
- For an academic evaluation, increase `iterations` and provide larger datasets (anonymized) to measure scaling behavior.
- Consider exporting results and plotting (e.g., in Python or Excel) for inclusion in your thesis.
