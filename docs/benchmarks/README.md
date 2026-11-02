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

This establishes structured resource measurements. Automated regression thresholds
and warnings are a subsequent step; a successful run does not assert a performance
budget or prove the absence of temporary-file leaks on every path.
