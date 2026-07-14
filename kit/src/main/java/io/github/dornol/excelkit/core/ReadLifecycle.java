package io.github.dornol.excelkit.core;

import org.jspecify.annotations.Nullable;

import java.time.Duration;
import java.util.concurrent.atomic.AtomicBoolean;
import java.util.concurrent.atomic.AtomicLong;

/** Tracks single-use state, row counts, elapsed time, and detailed progress. */
final class ReadLifecycle {
    private final AtomicLong successes = new AtomicLong();
    private final AtomicLong errors = new AtomicLong();
    private final AtomicBoolean consumed = new AtomicBoolean();
    private final long startedNanos = System.nanoTime();

    void markConsumed() {
        if (!consumed.compareAndSet(false, true))
            throw new ExcelKitException("Read handler has already been consumed");
    }

    void record(boolean success) {
        if (success) successes.incrementAndGet();
        else errors.incrementAndGet();
    }

    void progress(long processed, int sheet, long total,
                  @Nullable ReadProgressCallback callback) {
        if (callback != null) callback.onProgress(event(processed, sheet, total, false, false));
    }

    void complete(int sheet, long total, boolean cancelled,
                  @Nullable ReadProgressCallback callback) {
        if (callback != null) callback.onProgress(event(successes.get() + errors.get(), sheet, total,
                true, cancelled));
    }

    private ReadProgress event(long processed, int sheet, long total, boolean completed, boolean cancelled) {
        return new ReadProgress(processed, successes.get(), errors.get(), sheet, total,
                Duration.ofNanos(System.nanoTime() - startedNanos), completed, cancelled);
    }
}
