package io.github.dornol.excelkit.core.internal;

import io.github.dornol.excelkit.core.*;
import java.time.Duration;
import java.util.ArrayList;
import java.util.List;
import java.util.function.*;

/** Internal shared aggregation logic for format-specific readers. */
public final class ReaderExecutionSupport {
    private ReaderExecutionSupport() {}

    public static <T> ReadSummary summarize(Consumer<Consumer<ReadResult<T>>> execution,
            BooleanSupplier stoppedEarly, Consumer<ReadResult<T>> consumer) {
        long started = System.nanoTime();
        long[] counts = new long[3];
        execution.accept(result -> {
            counts[0]++;
            if (result.success()) counts[1]++; else counts[2]++;
            consumer.accept(result);
        });
        return new ReadSummary(counts[0], counts[1], counts[2], stoppedEarly.getAsBoolean(),
                Duration.ofNanos(System.nanoTime() - started));
    }

    public static <T> ReadReport report(Consumer<Consumer<ReadResult<T>>> execution,
            BooleanSupplier stoppedEarly, int maximum) {
        if (maximum < 0) throw new IllegalArgumentException("maxCollectedErrors must be non-negative");
        List<RowError> errors = new ArrayList<>();
        long[] row = {0};
        ReadSummary summary = summarize(execution, stoppedEarly, result -> {
            row[0]++;
            if (!result.success() && errors.size() < maximum) errors.add(new RowError(row[0],
                    result.fileRowNum(), result.cause() == null ? RowError.Type.VALIDATION : RowError.Type.MAPPING,
                    result.messages() == null ? List.of() : result.messages(), result.cause(),
                    result.cellErrors(), result.rawValues()));
        });
        return new ReadReport(summary, errors, summary.errorRows() > errors.size());
    }
}
