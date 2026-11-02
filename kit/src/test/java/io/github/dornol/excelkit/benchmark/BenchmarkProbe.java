package io.github.dornol.excelkit.benchmark;

import java.io.IOException;
import java.lang.management.ManagementFactory;
import java.nio.file.FileVisitResult;
import java.nio.file.Files;
import java.nio.file.NoSuchFileException;
import java.nio.file.Path;
import java.nio.file.SimpleFileVisitor;
import java.nio.file.attribute.BasicFileAttributes;
import java.util.concurrent.Executors;
import java.util.concurrent.ScheduledExecutorService;
import java.util.concurrent.TimeUnit;
import java.util.concurrent.atomic.AtomicReference;

/** Samples process heap and logical temporary-file bytes. Peaks are observed lower bounds. */
final class BenchmarkProbe implements AutoCloseable {
    record Snapshot(long heapBaselineBytes, long heapPeakBytes, long tempBaselineBytes,
                    long tempPeakBytes, long tempRemainingBytes, long samples, long intervalMillis) { }

    private final Path root;
    private final Path excluded;
    private final long intervalMillis;
    private final ScheduledExecutorService sampler;
    private final AtomicReference<IOException> samplingFailure = new AtomicReference<>();
    private final long heapBaseline;
    private final long tempBaseline;
    private long heapPeak;
    private long tempPeak;
    private long tempCurrent;
    private long samples;
    private Snapshot snapshot;

    static BenchmarkProbe start(Path excluded) throws IOException {
        return new BenchmarkProbe(Path.of(System.getProperty("java.io.tmpdir")), excluded,
                Long.getLong("excelkit.benchmark.sampleMillis", 10L), true);
    }

    BenchmarkProbe(Path root, Path excluded, long intervalMillis, boolean background) throws IOException {
        if (intervalMillis < 1) throw new IllegalArgumentException("Sampling interval must be positive");
        this.root = root.toAbsolutePath().normalize();
        this.excluded = excluded.toAbsolutePath().normalize();
        this.intervalMillis = intervalMillis;
        sampleNow();
        heapBaseline = heapPeak;
        tempBaseline = tempCurrent;
        sampler = Executors.newSingleThreadScheduledExecutor(runnable -> {
            var thread = new Thread(runnable, "benchmark-resource-sampler");
            thread.setDaemon(true);
            return thread;
        });
        if (background) {
            sampler.scheduleWithFixedDelay(() -> {
                try { sampleNow(); }
                catch (IOException e) { samplingFailure.compareAndSet(null, e); }
            }, intervalMillis, intervalMillis, TimeUnit.MILLISECONDS);
        }
    }

    synchronized void sampleNow() throws IOException {
        heapPeak = Math.max(heapPeak, ManagementFactory.getMemoryMXBean().getHeapMemoryUsage().getUsed());
        tempCurrent = directoryBytes(root, excluded);
        tempPeak = Math.max(tempPeak, tempCurrent);
        samples++;
    }

    Snapshot stop() throws IOException {
        if (snapshot != null) return snapshot;
        sampler.shutdown();
        try {
            if (!sampler.awaitTermination(5, TimeUnit.SECONDS)) {
                sampler.shutdownNow();
                throw new IOException("Resource sampler did not stop");
            }
        } catch (InterruptedException e) {
            sampler.shutdownNow();
            Thread.currentThread().interrupt();
            throw new IOException("Interrupted while stopping resource sampler", e);
        }
        if (samplingFailure.get() != null) throw samplingFailure.get();
        sampleNow();
        snapshot = new Snapshot(heapBaseline, heapPeak, tempBaseline, tempPeak, tempCurrent, samples, intervalMillis);
        return snapshot;
    }

    @Override
    public void close() throws IOException { stop(); }

    static long directoryBytes(Path root, Path excluded) throws IOException {
        var bytes = new long[1];
        Files.walkFileTree(root, new SimpleFileVisitor<>() {
            @Override
            public FileVisitResult preVisitDirectory(Path dir, BasicFileAttributes attributes) {
                return dir.equals(excluded) ? FileVisitResult.SKIP_SUBTREE : FileVisitResult.CONTINUE;
            }
            @Override
            public FileVisitResult visitFile(Path file, BasicFileAttributes attributes) {
                if (attributes.isRegularFile()) bytes[0] += attributes.size();
                return FileVisitResult.CONTINUE;
            }
            @Override
            public FileVisitResult visitFileFailed(Path file, IOException failure) throws IOException {
                // A spool file may disappear between directory enumeration and stat.
                if (!file.equals(root) && failure instanceof NoSuchFileException) return FileVisitResult.CONTINUE;
                throw failure;
            }
            @Override
            public FileVisitResult postVisitDirectory(Path dir, IOException failure) throws IOException {
                if (failure != null && (dir.equals(root) || !(failure instanceof NoSuchFileException))) {
                    throw failure;
                }
                return FileVisitResult.CONTINUE;
            }
        });
        return bytes[0];
    }
}
