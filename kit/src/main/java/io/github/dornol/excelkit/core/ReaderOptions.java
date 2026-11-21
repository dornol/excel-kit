package io.github.dornol.excelkit.core;

import jakarta.validation.Validator;
import org.jspecify.annotations.Nullable;

import java.util.Objects;
import java.util.function.UnaryOperator;

/**
 * Reusable common reader configuration. Format-specific settings remain on each reader.
 * <p>
 * The {@link Builder} is also what every reader keeps its own common settings in, so a setting given
 * by chaining on a reader and the same setting given through {@code ReaderOptions} are the same code:
 * one set of defaults and one set of checks.
 */
public final class ReaderOptions {
    private final ReadOptions readOptions;
    private final @Nullable Validator validator;
    private final int headerRowIndex;

    private ReaderOptions(Builder builder) {
        // Building the snapshot also runs its validation, so invalid options fail here.
        this.readOptions = builder.toReadOptions();
        this.validator = builder.validator;
        this.headerRowIndex = builder.headerRowIndex;
    }

    public static Builder builder() { return new Builder(); }

    ReadOptions toReadOptions() { return readOptions; }

    /** A mutable copy, so a reader that is given these options never modifies them. */
    Builder toBuilder() { return new Builder(this); }

    /** Mutable builder used to create reusable immutable options. */
    public static final class Builder {
        private @Nullable Validator validator;
        private int headerRowIndex;
        private int progressInterval;
        private @Nullable ProgressCallback progressCallback;
        private boolean strictHeaders;
        private DuplicateHeaderPolicy duplicateHeaderPolicy = DuplicateHeaderPolicy.FAIL;
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

        public Builder() { }

        private Builder(ReaderOptions source) {
            ReadOptions read = source.readOptions;
            this.validator = source.validator;
            this.headerRowIndex = source.headerRowIndex;
            this.progressInterval = read.progressInterval();
            this.progressCallback = read.progressCallback();
            this.strictHeaders = read.strictHeaders();
            this.duplicateHeaderPolicy = read.duplicateHeaderPolicy();
            this.cellConversionConfig = read.cellConversionConfig();
            this.maxRows = read.maxRows();
            this.skipBlankRows = read.skipBlankRows();
            this.stopAtBlankRows = read.stopAtBlankRows();
            this.maxErrors = read.maxErrors();
            this.headerNormalizer = read.headerNormalizer();
            this.limits = read.limits();
            this.cancellationToken = read.cancellationToken();
            this.readProgressInterval = read.readProgressInterval();
            this.readProgressCallback = read.readProgressCallback();
            this.securityPolicy = read.securityPolicy();
        }

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

        /** An immutable snapshot of the settings that read handlers consume. */
        ReadOptions toReadOptions() {
            return new ReadOptions(strictHeaders, duplicateHeaderPolicy, cellConversionConfig, maxRows,
                    skipBlankRows, stopAtBlankRows, maxErrors, headerNormalizer, limits, cancellationToken,
                    progressInterval, progressCallback, readProgressInterval, readProgressCallback, securityPolicy);
        }

        @Nullable Validator configuredValidator() { return validator; }
        int configuredHeaderRowIndex() { return headerRowIndex; }

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
