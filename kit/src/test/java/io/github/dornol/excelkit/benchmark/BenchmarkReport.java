package io.github.dornol.excelkit.benchmark;

import java.io.IOException;
import java.lang.management.ManagementFactory;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.StandardOpenOption;
import java.util.List;
import java.util.stream.Collectors;

final class BenchmarkReport {
    static final String HEADER = "schema_version,scenario,rows,elapsed_nanos,output_bytes,heap_baseline_bytes,heap_peak_bytes,temp_baseline_bytes,temp_peak_bytes,temp_remaining_bytes,samples,sample_interval_ms,java_version,os,arch,max_heap_bytes,processors,gc_collectors,revision";

    private BenchmarkReport() { }

    static void append(String scenario, long rows, long elapsed, long outputBytes,
                                    BenchmarkProbe.Snapshot snapshot) throws IOException {
        Path path = Path.of(System.getProperty("excelkit.benchmark.report", "build/reports/benchmark/resources.csv"));
        append(path, scenario, rows, elapsed, outputBytes, snapshot);
    }

    static synchronized void append(Path path, String scenario, long rows, long elapsed, long outputBytes,
                                    BenchmarkProbe.Snapshot snapshot) throws IOException {
        Files.createDirectories(path.toAbsolutePath().getParent());
        if (!Files.exists(path)) Files.writeString(path, HEADER + "\n", StandardOpenOption.CREATE_NEW);
        List<Object> values = List.of(1, scenario, rows, elapsed, outputBytes,
                snapshot.heapBaselineBytes(), snapshot.heapPeakBytes(), snapshot.tempBaselineBytes(),
                snapshot.tempPeakBytes(), snapshot.tempRemainingBytes(), snapshot.samples(), snapshot.intervalMillis(),
                System.getProperty("java.runtime.version"), System.getProperty("os.name"), System.getProperty("os.arch"),
                Runtime.getRuntime().maxMemory(), Runtime.getRuntime().availableProcessors(),
                ManagementFactory.getGarbageCollectorMXBeans().stream().map(bean -> bean.getName())
                        .collect(Collectors.joining(";")),
                System.getenv().getOrDefault("GITHUB_SHA", "local"));
        String line = values.stream().map(value -> "\"" + value.toString().replace("\"", "\"\"") + "\"")
                .collect(Collectors.joining(","));
        Files.writeString(path, line + "\n", StandardOpenOption.APPEND);
    }
}
