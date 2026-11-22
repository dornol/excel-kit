package io.github.dornol.excelkit.core.read;

import java.time.Duration;
import java.util.Objects;

/** Aggregate outcome of one reader execution. */
public record ReadSummary(long totalRows, long successRows, long errorRows,
                          boolean stoppedEarly, Duration duration) {
    public ReadSummary {
        if (totalRows < 0 || successRows < 0 || errorRows < 0) {
            throw new IllegalArgumentException("row counts must be non-negative");
        }
        Objects.requireNonNull(duration, "duration cannot be null");
    }
}
