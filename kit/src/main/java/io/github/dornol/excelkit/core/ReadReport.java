package io.github.dornol.excelkit.core;

import java.util.List;

/** Read summary with a bounded list of row errors. */
public record ReadReport(ReadSummary summary, List<RowError> errors, boolean errorsTruncated) {
    public ReadReport {
        java.util.Objects.requireNonNull(summary, "summary cannot be null");
        errors = List.copyOf(errors);
    }
}
