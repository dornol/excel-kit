package io.github.dornol.excelkit.core.read;

import java.util.List;
import java.util.Objects;

/** Read summary with a bounded list of row errors. */
public record ReadReport(ReadSummary summary, List<RowError> errors, boolean errorsTruncated) {
    public ReadReport {
        Objects.requireNonNull(summary, "summary cannot be null");
        errors = List.copyOf(errors);
    }
}
