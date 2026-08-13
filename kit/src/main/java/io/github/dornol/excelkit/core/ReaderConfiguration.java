package io.github.dornol.excelkit.core;

import jakarta.validation.Validator;
import org.jspecify.annotations.Nullable;

import java.util.function.UnaryOperator;

/** Mutable format-independent settings owned by a reader instance. */
final class ReaderConfiguration {

    @Nullable Validator validator;
    int headerRowIndex;
    @Nullable ProgressCallback progressCallback;
    int progressInterval;
    boolean strictHeaders;
    DuplicateHeaderPolicy duplicateHeaderPolicy = DuplicateHeaderPolicy.FIRST;
    @Nullable CellConversionConfig cellConversionConfig;
    long maxRows = -1;
    boolean skipBlankRows;
    int stopAtBlankRows;
    long maxErrors = -1;
    UnaryOperator<String> headerNormalizer = UnaryOperator.identity();
    ReadLimits limits = ReadLimits.UNLIMITED;
    CancellationToken cancellationToken = CancellationToken.NONE;
    int readProgressInterval;
    @Nullable ReadProgressCallback readProgressCallback;
    ReadSecurityPolicy securityPolicy = ReadSecurityPolicy.DEFAULT;

    ReadOptions snapshot() {
        return new ReadOptions(strictHeaders, duplicateHeaderPolicy, cellConversionConfig, maxRows,
                skipBlankRows, stopAtBlankRows, maxErrors, headerNormalizer, limits, cancellationToken,
                progressInterval, progressCallback, readProgressInterval, readProgressCallback, securityPolicy);
    }
}
