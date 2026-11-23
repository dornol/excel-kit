package io.github.dornol.excelkit.core.read;

import io.github.dornol.excelkit.core.ProgressCallback;
import jakarta.validation.Validator;
import org.jspecify.annotations.Nullable;

import java.util.List;
import java.util.Objects;
import java.util.function.UnaryOperator;

/**
 * Reusable common reader configuration. Format-specific settings remain on each reader.
 * <p>
 * The {@link Builder} is also what every reader keeps its own common settings in, so a setting given
 * by chaining on a reader and the same setting given through {@code ReaderOptions} are the same code:
 * one set of defaults and one set of checks.
 *
 * @since 0.23.0
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

    /** Starts a new, empty set of options. */
    public static Builder builder() { return new Builder(); }

    ReadOptions toReadOptions() { return readOptions; }

    /** A builder that starts from these options, so changing it never modifies them. */
    public Builder toBuilder() { return new Builder(this); }

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
        private List<String> uniqueKey = List.of();

        /** Creates a builder with every setting at its default. */
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
            this.uniqueKey = read.uniqueKey();
        }

        /** The Bean Validation provider that validates every mapped row; {@code null} disables validation. */
        public Builder validator(@Nullable Validator validator) { this.validator = validator; return this; }
        /** The zero-based index of the header row. */
        public Builder headerRowIndex(int value) { this.headerRowIndex = value; return this; }
        /** Calls back every {@code interval} rows (positive) with the legacy row counter; see {@link ReadProgress} for richer reporting. */
        public Builder onProgress(int interval, ProgressCallback callback) {
            this.progressInterval = positive(interval, "progress interval");
            this.progressCallback = Objects.requireNonNull(callback, "callback cannot be null");
            return this;
        }
        /** Requires positional and index-based column bindings to resolve to an existing header before any data row is read. */
        public Builder strictHeaders(boolean value) { this.strictHeaders = value; return this; }
        /** Requires positional and index-based column bindings to resolve to an existing header before any data row is read. */
        public Builder strictHeaders() { return strictHeaders(true); }
        /** How duplicate header names are handled; defaults to {@link DuplicateHeaderPolicy#FAIL}. */
        public Builder duplicateHeaderPolicy(DuplicateHeaderPolicy value) {
            this.duplicateHeaderPolicy = Objects.requireNonNull(value, "policy cannot be null"); return this;
        }
        /** Locale and date patterns used by {@link CellData} conversions. */
        public Builder cellConversion(CellConversionConfig value) {
            this.cellConversionConfig = Objects.requireNonNull(value, "config cannot be null"); return this;
        }
        /** Stops after this many emitted (non-skipped) data rows; must not be negative. */
        public Builder maxRows(long value) { this.maxRows = nonNegative(value, "maxRows"); return this; }
        /** Skips rows in which every cell is blank. */
        public Builder skipBlankRows(boolean value) { this.skipBlankRows = value; return this; }
        /** Skips rows in which every cell is blank. */
        public Builder skipBlankRows() { return skipBlankRows(true); }
        /** Stops after this many consecutive blank data rows; {@code 0} disables. */
        public Builder stopAtBlankRows(int value) { this.stopAtBlankRows = nonNegative(value, "count"); return this; }
        /** Aborts after this many failed rows; {@code 0} aborts on the first one, a negative default means no limit. */
        public Builder maxErrors(long value) { this.maxErrors = nonNegative(value, "maxErrors"); return this; }
        /** Transforms headers before names and aliases are compared. */
        public Builder headerNormalizer(UnaryOperator<String> value) {
            this.headerNormalizer = Objects.requireNonNull(value, "normalizer cannot be null"); return this;
        }
        /** Bounds on input size, sheets, columns, cell size and cell count; see {@link ReadLimits#UPLOAD}. */
        public Builder limits(ReadLimits value) { this.limits = Objects.requireNonNull(value, "limits cannot be null"); return this; }
        /** A signal that stops the read before the next row once it reports a cancellation. */
        public Builder cancellationToken(CancellationToken value) {
            this.cancellationToken = Objects.requireNonNull(value, "token cannot be null"); return this;
        }
        /** Reports {@link ReadProgress} snapshots every {@code interval} rows (positive) and once when the read ends. */
        public Builder onReadProgress(int interval, ReadProgressCallback callback) {
            this.readProgressInterval = positive(interval, "progress interval");
            this.readProgressCallback = Objects.requireNonNull(callback, "callback cannot be null");
            return this;
        }
        /** Excel content restrictions; see {@link ReadSecurityPolicy#STRICT}. */
        public Builder securityPolicy(ReadSecurityPolicy value) {
            this.securityPolicy = Objects.requireNonNull(value, "policy cannot be null"); return this;
        }
        /** The headers whose values together must be unique across the rows; see {@code AbstractReader#uniqueKey}. */
        public Builder uniqueKey(String... headers) {
            Objects.requireNonNull(headers, "headers cannot be null");
            if (headers.length == 0) { this.uniqueKey = List.of(); return this; }
            var copy = new java.util.ArrayList<String>(headers.length);
            for (String header : headers) {
                Objects.requireNonNull(header, "a header cannot be null");
                if (header.isBlank()) throw new IllegalArgumentException("a unique key header must not be blank");
                copy.add(header);
            }
            this.uniqueKey = List.copyOf(copy);
            return this;
        }
        /** Validates the settings and returns them as immutable options. */
        public ReaderOptions build() { return new ReaderOptions(this); }

        /** An immutable snapshot of the settings that read handlers consume. */
        ReadOptions toReadOptions() {
            return new ReadOptions(strictHeaders, duplicateHeaderPolicy, cellConversionConfig, maxRows,
                    skipBlankRows, stopAtBlankRows, maxErrors, headerNormalizer, limits, cancellationToken,
                    progressInterval, progressCallback, readProgressInterval, readProgressCallback, securityPolicy, uniqueKey);
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
