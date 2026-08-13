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

    void apply(ReaderOptions options) {
        validator = options.validator();
        headerRowIndex = options.headerRowIndex();
        progressInterval = options.progressInterval();
        progressCallback = options.progressCallback();
        ReadOptions snapshot = options.toReadOptions();
        strictHeaders = snapshot.strictHeaders();
        duplicateHeaderPolicy = snapshot.duplicateHeaderPolicy();
        cellConversionConfig = snapshot.cellConversionConfig();
        maxRows = snapshot.maxRows();
        skipBlankRows = snapshot.skipBlankRows();
        stopAtBlankRows = snapshot.stopAtBlankRows();
        maxErrors = snapshot.maxErrors();
        headerNormalizer = snapshot.headerNormalizer();
        limits = snapshot.limits();
        cancellationToken = snapshot.cancellationToken();
        readProgressInterval = snapshot.readProgressInterval();
        readProgressCallback = snapshot.readProgressCallback();
        securityPolicy = snapshot.securityPolicy();
    }

    ReadOptions snapshot() {
        return new ReadOptions(strictHeaders, duplicateHeaderPolicy, cellConversionConfig, maxRows,
                skipBlankRows, stopAtBlankRows, maxErrors, headerNormalizer, limits, cancellationToken,
                progressInterval, progressCallback, readProgressInterval, readProgressCallback, securityPolicy);
    }
}
