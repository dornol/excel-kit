package io.github.dornol.excelkit.core;

import org.jspecify.annotations.Nullable;

import java.util.Objects;
import java.util.function.UnaryOperator;

/** Immutable snapshot of format-independent reader settings. */
public record ReadOptions(
        boolean strictHeaders,
        DuplicateHeaderPolicy duplicateHeaderPolicy,
        @Nullable CellConversionConfig cellConversionConfig,
        long maxRows,
        boolean skipBlankRows,
        int stopAtBlankRows,
        long maxErrors,
        UnaryOperator<String> headerNormalizer
) {
    public ReadOptions {
        Objects.requireNonNull(duplicateHeaderPolicy, "duplicateHeaderPolicy cannot be null");
        Objects.requireNonNull(headerNormalizer, "headerNormalizer cannot be null");
        if (maxRows < -1) throw new IllegalArgumentException("maxRows must be >= -1");
        if (stopAtBlankRows < 0) throw new IllegalArgumentException("stopAtBlankRows must be non-negative");
        if (maxErrors < -1) throw new IllegalArgumentException("maxErrors must be >= -1");
    }
}
