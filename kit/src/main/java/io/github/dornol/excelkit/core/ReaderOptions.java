package io.github.dornol.excelkit.core;

import jakarta.validation.Validator;
import org.jspecify.annotations.Nullable;

import java.util.Objects;
import java.util.function.UnaryOperator;

/** Reusable common reader configuration. Format-specific settings remain on each reader. */
public final class ReaderOptions {
    private final @Nullable Validator validator;
    private final int headerRowIndex;
    private final int progressInterval;
    private final @Nullable ProgressCallback progressCallback;
    private final boolean strictHeaders;
    private final DuplicateHeaderPolicy duplicateHeaderPolicy;
    private final @Nullable CellConversionConfig cellConversionConfig;
    private final long maxRows;
    private final boolean skipBlankRows;
    private final int stopAtBlankRows;
    private final long maxErrors;
    private final UnaryOperator<String> headerNormalizer;
    private final ReadLimits limits;
    private final CancellationToken cancellationToken;
    private final int readProgressInterval;
    private final @Nullable ReadProgressCallback readProgressCallback;
    private final ReadSecurityPolicy securityPolicy;

    private ReaderOptions(Builder builder) {
        this.validator = builder.validator;
        this.headerRowIndex = builder.headerRowIndex;
        this.progressInterval = builder.progressInterval;
        this.progressCallback = builder.progressCallback;
        this.strictHeaders = builder.strictHeaders;
        this.duplicateHeaderPolicy = builder.duplicateHeaderPolicy;
        this.cellConversionConfig = builder.cellConversionConfig;
        this.maxRows = builder.maxRows;
        this.skipBlankRows = builder.skipBlankRows;
        this.stopAtBlankRows = builder.stopAtBlankRows;
        this.maxErrors = builder.maxErrors;
        this.headerNormalizer = builder.headerNormalizer;
        this.limits = builder.limits;
        this.cancellationToken = builder.cancellationToken;
        this.readProgressInterval = builder.readProgressInterval;
        this.readProgressCallback = builder.readProgressCallback;
        this.securityPolicy = builder.securityPolicy;
        toReadOptions();
    }

    public static Builder builder() { return new Builder(); }

    ReadOptions toReadOptions() {
        return new ReadOptions(strictHeaders, duplicateHeaderPolicy, cellConversionConfig, maxRows,
                skipBlankRows, stopAtBlankRows, maxErrors, headerNormalizer, limits, cancellationToken,
                progressInterval, progressCallback, readProgressInterval, readProgressCallback, securityPolicy);
    }

    @Nullable Validator validator() { return validator; }
    int headerRowIndex() { return headerRowIndex; }
    int progressInterval() { return progressInterval; }
    @Nullable ProgressCallback progressCallback() { return progressCallback; }

    /** Mutable builder used to create reusable immutable options. */
    public static final class Builder {
        private @Nullable Validator validator;
        private int headerRowIndex;
        private int progressInterval;
        private @Nullable ProgressCallback progressCallback;
        private boolean strictHeaders;
        private DuplicateHeaderPolicy duplicateHeaderPolicy = DuplicateHeaderPolicy.FIRST;
        private @Nullable CellConversionConfig cellConversionConfig;
        private long maxRows = -1;
        private boolean skipBlankRows;
        private int stopAtBlankRows;
        private long maxErrors = -1;
        private UnaryOperator<String> headerNormalizer = UnaryOperator.identity();
        private ReadLimits limits = ReadLimits.UNLIMITED;
        private CancellationToken cancellationToken = CancellationToken.NONE;
        private int readProgressInterval;
        private @Nullable ReadProgressCallback readProgressCallback;
        private ReadSecurityPolicy securityPolicy = ReadSecurityPolicy.DEFAULT;

        public Builder validator(@Nullable Validator validator) { this.validator = validator; return this; }
        public Builder headerRowIndex(int value) { this.headerRowIndex = value; return this; }
        public Builder onProgress(int interval, ProgressCallback callback) {
            this.progressInterval = positive(interval, "progress interval");
            this.progressCallback = Objects.requireNonNull(callback, "callback cannot be null");
            return this;
        }
        public Builder strictHeaders(boolean value) { this.strictHeaders = value; return this; }
        public Builder strictHeaders() { return strictHeaders(true); }
        public Builder duplicateHeaderPolicy(DuplicateHeaderPolicy value) {
            this.duplicateHeaderPolicy = Objects.requireNonNull(value, "policy cannot be null"); return this;
        }
        public Builder cellConversion(CellConversionConfig value) {
            this.cellConversionConfig = Objects.requireNonNull(value, "config cannot be null"); return this;
        }
        public Builder maxRows(long value) { this.maxRows = nonNegative(value, "maxRows"); return this; }
        public Builder skipBlankRows(boolean value) { this.skipBlankRows = value; return this; }
        public Builder skipBlankRows() { return skipBlankRows(true); }
        public Builder stopAtBlankRows(int value) { this.stopAtBlankRows = nonNegative(value, "count"); return this; }
        public Builder maxErrors(long value) { this.maxErrors = nonNegative(value, "maxErrors"); return this; }
        public Builder headerNormalizer(UnaryOperator<String> value) {
            this.headerNormalizer = Objects.requireNonNull(value, "normalizer cannot be null"); return this;
        }
        public Builder limits(ReadLimits value) { this.limits = Objects.requireNonNull(value, "limits cannot be null"); return this; }
        public Builder cancellationToken(CancellationToken value) {
            this.cancellationToken = Objects.requireNonNull(value, "token cannot be null"); return this;
        }
        public Builder onReadProgress(int interval, ReadProgressCallback callback) {
            this.readProgressInterval = positive(interval, "progress interval");
            this.readProgressCallback = Objects.requireNonNull(callback, "callback cannot be null");
            return this;
        }
        public Builder securityPolicy(ReadSecurityPolicy value) {
            this.securityPolicy = Objects.requireNonNull(value, "policy cannot be null"); return this;
        }
        public ReaderOptions build() { return new ReaderOptions(this); }

        private static int positive(int value, String name) {
            if (value <= 0) throw new IllegalArgumentException(name + " must be positive");
            return value;
        }
        private static int nonNegative(int value, String name) {
            if (value < 0) throw new IllegalArgumentException(name + " must be non-negative");
            return value;
        }
        private static long nonNegative(long value, String name) {
            if (value < 0) throw new IllegalArgumentException(name + " must be non-negative");
            return value;
        }
    }
}
