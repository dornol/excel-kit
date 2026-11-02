# Benchmark resource measurements

Run the benchmark suite with an installed JDK (CI uses JDK 21):

```bash
./gradlew :kit:benchmark --rerun-tasks --no-daemon -PtestJavaVersion=21
```

The task produces the usual HTML test report and
`kit/build/reports/benchmark/resources.csv`. Each CSV row identifies a scenario by
its test method name and records elapsed nanoseconds, output bytes, heap baseline
and observed peak bytes, temporary-file baseline/observed peak/remaining bytes,
sample count and configured interval, JVM/OS/architecture, heap limit and processor
count, GC collector names and CI revision. CI uploads both reports. A fresh run replaces the CSV instead of appending
results from earlier runs.

## What is measured

- Sampling begins at the existing workload boundary, after any input fixture has
  been generated, and ends after the operation. Reader fixture generation is not
  included in its elapsed time or resource peaks.
- Heap usage is read from `MemoryMXBean.getHeapMemoryUsage().getUsed()`. It is the
  entire benchmark JVM's used heap, including uncollected garbage and sampler
  overhead. It is not retained memory, allocation volume, native memory or RSS.
  No explicit GC is requested. Baseline heap helps interpret prior-scenario carryover.
- Temporary-file usage is the sum of regular files' logical lengths under the
  benchmark JVM's isolated `java.io.tmpdir`. This includes library staging files,
  POI spool files and encryption staging files. The current JUnit fixture directory
  is excluded, so input/output XLSX and CSV files are not counted again as spool
  usage. Output size is reported separately. Symlinks are not followed.
- Remaining temporary bytes are measured at the workload boundary, before JUnit
  deletes its fixture directory. A nonzero value warrants investigation; the
  baseline distinguishes pre-existing files from files left by this operation.
- Sampling runs sequentially with a nominal 10 ms delay between samples. The
  scan itself, scheduling and GC pauses can lengthen the actual interval. These
  are **observed peaks (lower bounds)**: a file or heap spike created and freed
  between samples can be missed. Logical file length is not physical disk allocation.
  Sampling failures fail the benchmark instead of producing a misleading zero.

Use `-PbenchmarkSampleMillis=5` to change the nominal interval. Lower intervals
increase measurement overhead, so compare runs using the same interval.

## Comparing results

Use the same scenario, revision/workload, JDK, heap limit, GC configuration,
OS/architecture and machine resources. The benchmark suite has no per-scenario
forks or warmup/measurement repetitions; cold-start effects, execution order and
CI machine noise affect these numbers. Compare repeated runs before attributing
changes to the library. Existing README timings predate resource sampling and
are not a baseline for this instrumented suite.

## Advisory comparisons in GitHub Actions

After a successful benchmark run, CI searches the latest 30 successful `push`
runs of this repository's `Benchmark` workflow on `main`, newest first. It ignores
this run and runs started at or after it, PR/fork results, failed runs and expired
or unavailable artifacts. The chosen artifact's CSV revision must match its
workflow commit.

A baseline must match all of the following:

- CSV schema, exact JVM version, OS/architecture, processor count, heap limit,
  GC collector names and nominal sampling interval.
- Runner image metadata and CPU model (when available).
- A SHA-256 fingerprint of the benchmark scenarios, sampler and CSV writer.
- Scenario set, execution order and row counts.

Missing context, environment changes or workload changes therefore start a new
baseline series. No comparison is preferable to treating a different workload as
a performance change. The first run after this feature is introduced may have no
compatible baseline; older artifacts do not include `context.json`.

The Actions job summary shows previous/current time, observed heap peak and
observed temporary-file peak with percentage changes for every scenario. A zero
baseline becoming nonzero is shown as `new usage`, rather than an infinite
percentage. CSV/context, `comparison.md` and `comparison.json` are uploaded in
`benchmark-results` for later analysis. Keep CSV and context from the same run.

**The initial policy applies no calibrated regression threshold.** Any upward
metric change produces one aggregate warning annotation; it is an observation,
not a claim that a regression is confirmed. Repeated runs should establish the
normal variability before percentage/absolute thresholds are chosen. This avoids
silently inventing a performance budget. Warnings never fail the benchmark job.
Invalid current measurement data still fails validation; lookup errors, absent
baselines and incompatible historical reports instead leave a current-results
summary with selection details and preserve the new measurements.

The job has read-only repository/Actions permissions. Baseline lookup uses the
existing `gh` CLI and does not post PR comments or modify historical artifacts.

## Comparing downloaded artifacts locally

Use the matching CSV and context from each run:

```bash
python3 scripts/benchmark_comparison.py compare \
  --current current/resources.csv --context current/context.json \
  --baseline previous/resources.csv --baseline-context previous/context.json \
  --output-dir build/benchmark-comparison
python3 -m unittest discover -s scripts/tests -p 'test_benchmark_comparison.py'
```

For a local benchmark, generate context before the workload with
`python3 scripts/benchmark_comparison.py context --output kit/build/reports/benchmark/context.json`.
A successful run does not assert a performance budget or prove the absence of
temporary-file leaks on every path.
