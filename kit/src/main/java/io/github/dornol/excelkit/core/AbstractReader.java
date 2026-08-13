package io.github.dornol.excelkit.core;

import jakarta.validation.Validator;
import org.jspecify.annotations.Nullable;

import java.util.ArrayList;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Objects;
import java.util.Set;
import java.util.function.BiConsumer;
import java.util.function.Consumer;
import java.util.function.Function;
import java.util.function.Supplier;
import java.io.FilterInputStream;
import java.io.InputStream;
import java.io.IOException;
import java.nio.file.Path;
import java.util.function.BiFunction;
import java.util.function.BooleanSupplier;
import java.util.function.Predicate;
import java.util.function.UnaryOperator;

/**
 * Shared reader configuration for {@link io.github.dornol.excelkit.excel.ExcelReader}
 * and {@link io.github.dornol.excelkit.csv.CsvReader}.
 * <p>
 * Contains column registration, progress, validation, and skip logic.
 * Format-specific configuration (sheet index, delimiter, charset, etc.) stays in subclasses.
 *
 * @param <T>    the row data type
 * @param <SELF> the concrete reader type, for fluent chaining
 * @author dhkim
 * @since 0.17.0
 */
@SuppressWarnings("unchecked")
public abstract class AbstractReader<T, SELF extends AbstractReader<T, SELF>> {

    protected final List<ReadColumn<T>> columns = new ArrayList<>();
    protected final @Nullable Supplier<T> instanceSupplier;
    protected final @Nullable Function<RowData, T> rowMapper;
    protected @Nullable Validator validator;
    private final ReaderConfiguration configuration = new ReaderConfiguration();
    protected boolean mapMode = false;
    protected @Nullable Set<String> selectedMapColumns;

    protected AbstractReader(Supplier<T> instanceSupplier, @Nullable Validator validator) {
        this.instanceSupplier = Objects.requireNonNull(instanceSupplier, "instanceSupplier cannot be null");
        this.rowMapper = null;
        this.validator = validator;
    }

    protected AbstractReader(Function<RowData, T> rowMapper, @Nullable Validator validator) {
        this.instanceSupplier = null;
        this.rowMapper = Objects.requireNonNull(rowMapper, "rowMapper cannot be null");
        this.validator = validator;
    }

    private SELF self() {
        return (SELF) this;
    }

    /**
     * Configures the Bean Validation provider used to validate each mapped row.
     *
     * @param validator Bean Validation provider
     * @return this reader for chaining
     */
    public SELF validator(Validator validator) {
        this.validator = Objects.requireNonNull(validator, "validator cannot be null");
        return self();
    }

    protected void requireNotMapMode(String method) {
        if (mapMode) {
            throw new IllegalStateException(
                    method + " cannot be called on a forMap() reader; "
                            + "map mode auto-discovers columns from the header row");
        }
    }

    /**
     * Sets the zero-based row index of the header row.
     */
    public SELF headerRowIndex(int headerRowIndex) {
        configuration.headerRowIndex = headerRowIndex;
        return self();
    }

    protected final int configuredHeaderRowIndex() {
        return configuration.headerRowIndex;
    }

    protected final int configuredProgressInterval() {
        return configuration.progressInterval;
    }

    protected final @Nullable ProgressCallback configuredProgressCallback() {
        return configuration.progressCallback;
    }

    void addColumn(ReadColumn<T> column) {
        columns.add(column);
    }

    protected void selectedMapColumns(@Nullable Set<String> selectedColumns) {
        this.selectedMapColumns = selectedColumns == null ? null : new LinkedHashSet<>(selectedColumns);
    }

    /**
     * Registers a positional column mapping.
     */
    public SELF column(BiConsumer<T, CellData> setter) {
        requireNotMapMode("column(BiConsumer)");
        columns.add(new ReadColumn<>(Objects.requireNonNull(setter, "setter cannot be null")));
        return self();
    }

    /**
     * Registers a name-based column mapping.
     */
    public SELF column(String headerName, BiConsumer<T, CellData> setter) {
        requireNotMapMode("column(String, BiConsumer)");
        columns.add(new ReadColumn<>(headerName,
                Objects.requireNonNull(setter, "setter cannot be null")));
        return self();
    }

    /**
     * Registers a name-based column mapping with header aliases.
     * The first matching alias, in list order, is used.
     */
    public SELF column(List<String> headerAliases, BiConsumer<T, CellData> setter) {
        requireNotMapMode("column(List, BiConsumer)");
        columns.add(new ReadColumn<>(headerAliases,
                Objects.requireNonNull(setter, "setter cannot be null")));
        return self();
    }

    /**
     * Registers an index-based column mapping.
     */
    public SELF columnAt(int columnIndex, BiConsumer<T, CellData> setter) {
        requireNotMapMode("columnAt(int, BiConsumer)");
        columns.add(new ReadColumn<>(null, columnIndex,
                Objects.requireNonNull(setter, "setter cannot be null")));
        return self();
    }

    /**
     * Marks the last registered column as required.
     */
    public SELF required() {
        if (columns.isEmpty()) {
            throw new IllegalStateException("required() must be called after column()");
        }
        int lastIndex = columns.size() - 1;
        columns.set(lastIndex, columns.get(lastIndex).required());
        return self();
    }

    /**
     * Skips one positional column.
     */
    public SELF skipColumn() {
        requireNotMapMode("skipColumn()");
        columns.add(new ReadColumn<>((instance, cellData) -> {}));
        return self();
    }

    /**
     * Skips the specified number of positional columns.
     */
    public SELF skipColumns(int count) {
        requireNotMapMode("skipColumns(int)");
        if (count < 0) {
            throw new IllegalArgumentException("skipColumns count must be non-negative");
        }
        for (int i = 0; i < count; i++) {
            columns.add(new ReadColumn<>((instance, cellData) -> {}));
        }
        return self();
    }

    /**
     * Registers a progress callback that fires every {@code interval} rows.
     */
    public SELF onProgress(int interval, ProgressCallback callback) {
        if (interval <= 0) {
            throw new IllegalArgumentException("progress interval must be positive");
        }
        configuration.progressInterval = interval;
        configuration.progressCallback = Objects.requireNonNull(callback, "callback cannot be null");
        return self();
    }

    /**
     * Enables strict header validation.
     * In strict mode, positional and index-based column bindings must resolve to
     * an existing header column before any data row is processed.
     */
    public SELF strictHeaders() {
        return strictHeaders(true);
    }

    /**
     * Enables or disables strict header validation.
     */
    public SELF strictHeaders(boolean enabled) {
        configuration.strictHeaders = enabled;
        return self();
    }

    /**
     * Alias for {@link #strictHeaders()}.
     */
    public SELF requireHeaders() {
        return strictHeaders();
    }

    /**
     * Sets duplicate header handling. Defaults to {@link DuplicateHeaderPolicy#FIRST}.
     */
    public SELF duplicateHeaderPolicy(DuplicateHeaderPolicy policy) {
        configuration.duplicateHeaderPolicy = Objects.requireNonNull(policy, "policy cannot be null");
        return self();
    }

    /**
     * Sets reader-scoped conversion settings for {@link CellData}.
     *
     * @since 0.19.0
     */
    public SELF cellConversion(CellConversionConfig config) {
        configuration.cellConversionConfig = Objects.requireNonNull(config, "config cannot be null");
        return self();
    }

    /**
     * Configures reader-scoped {@link CellData} conversion settings.
     *
     * @since 0.19.0
     */
    public SELF cellConversion(Consumer<CellConversionConfig.Builder> configurer) {
        Objects.requireNonNull(configurer, "configurer cannot be null");
        CellConversionConfig.Builder builder = CellConversionConfig.builder();
        configurer.accept(builder);
        return cellConversion(builder.build());
    }

    /**
     * Limits the number of non-skipped data rows emitted by this reader.
     *
     * @since 0.19.0
     */
    public SELF maxRows(long maxRows) {
        if (maxRows < 0) {
            throw new IllegalArgumentException("maxRows must be non-negative");
        }
        configuration.maxRows = maxRows;
        return self();
    }

    /**
     * Skips rows where every cell is blank.
     *
     * @since 0.19.0
     */
    public SELF skipBlankRows() {
        return skipBlankRows(true);
    }

    /**
     * Enables or disables skipping rows where every cell is blank.
     *
     * @since 0.19.0
     */
    public SELF skipBlankRows(boolean enabled) {
        configuration.skipBlankRows = enabled;
        return self();
    }

    /**
     * Stops reading after the given number of consecutive blank data rows.
     * Pass {@code 0} to disable.
     *
     * @since 0.19.0
     */
    public SELF stopAtBlankRows(int count) {
        if (count < 0) {
            throw new IllegalArgumentException("count must be non-negative");
        }
        configuration.stopAtBlankRows = count;
        return self();
    }

    /** Limits failed rows; zero aborts on the first failure. Disabled by default. */
    public SELF maxErrors(long maxErrors) {
        if (maxErrors < 0) {
            throw new IllegalArgumentException("maxErrors must be non-negative");
        }
        configuration.maxErrors = maxErrors;
        return self();
    }

    /** Normalizes headers before matching names and aliases. Original names remain available for diagnostics. */
    public SELF headerNormalizer(UnaryOperator<String> normalizer) {
        configuration.headerNormalizer = Objects.requireNonNull(normalizer, "normalizer cannot be null");
        return self();
    }

    public SELF headerPolicy(HeaderPolicy policy) {
        return headerNormalizer(Objects.requireNonNull(policy, "policy cannot be null").normalizer());
    }

    public SELF limits(ReadLimits limits) {
        configuration.limits = Objects.requireNonNull(limits, "limits cannot be null");
        return self();
    }

    public SELF cancellationToken(CancellationToken token) {
        configuration.cancellationToken = Objects.requireNonNull(token, "token cannot be null");
        return self();
    }

    public SELF onReadProgress(int interval, ReadProgressCallback callback) {
        if (interval <= 0) throw new IllegalArgumentException("progress interval must be positive");
        configuration.progressInterval = interval;
        configuration.readProgressCallback = Objects.requireNonNull(callback, "callback cannot be null");
        return self();
    }

    public SELF securityPolicy(ReadSecurityPolicy policy) {
        configuration.securityPolicy = Objects.requireNonNull(policy, "policy cannot be null");
        return self();
    }

    protected ReadOptions snapshotReadOptions() {
        return configuration.snapshot();
    }

    protected InputStream limitInput(InputStream input) {
        Objects.requireNonNull(input, "input cannot be null");
        long maximum = configuration.limits.maxInputBytes();
        if (maximum < 0) return input;
        return new FilterInputStream(input) {
            private long count;
            private void counted(long amount) {
                if (amount > 0 && (count += amount) > maximum) {
                    throw new ReadLimitExceededException(ReadLimitExceededException.Limit.INPUT_BYTES,
                            maximum, count);
                }
            }
            @Override public int read() throws IOException {
                int value = super.read();
                if (value >= 0) counted(1);
                return value;
            }
            @Override public int read(byte[] bytes, int offset, int length) throws IOException {
                int read = super.read(bytes, offset, length);
                counted(read);
                return read;
            }
            @Override public void close() { }
        };
    }

    protected <R> R executeFromSource(InputStreamSource source, Function<InputStream,R> execution,
            BiFunction<String,IOException,? extends RuntimeException> exceptionFactory) {
        return ReaderExecutionSupport.executeFromSource(source, execution, exceptionFactory);
    }

    protected InputStream requireInputStream(InputStream inputStream) {
        return Objects.requireNonNull(inputStream, "inputStream cannot be null");
    }

    protected Path requirePath(Path path) {
        return Objects.requireNonNull(path, "path cannot be null");
    }

    protected InputStreamSource requireInputSource(InputStreamSource source) {
        return Objects.requireNonNull(source, "source cannot be null");
    }

    /** Creates a format-specific handler and executes one operation against it. */
    protected <H, R> R executeWithInputStream(InputStream inputStream,
                                               Function<InputStream, H> handlerFactory,
                                               Function<H, R> execution) {
        return ReaderExecutionSupport.executeWithInputStream(inputStream, handlerFactory, execution);
    }

    protected <H> void executeInputStream(InputStream inputStream,
                                          Function<InputStream, H> handlerFactory,
                                          Consumer<H> execution) {
        ReaderExecutionSupport.executeInputStream(inputStream, handlerFactory, execution);
    }

    /** Creates a format-specific path handler and executes one operation against it. */
    protected <H, R> R executeWithPath(Path path,
                                       Function<Path, H> handlerFactory,
                                       Function<H, R> execution) {
        return ReaderExecutionSupport.executeWithPath(path, handlerFactory, execution);
    }

    protected <H> void executePath(Path path,
                                   Function<Path, H> handlerFactory,
                                   Consumer<H> execution) {
        ReaderExecutionSupport.executePath(path, handlerFactory, execution);
    }

    protected <C> C requireCallback(C callback, String name) {
        return Objects.requireNonNull(callback, name + " cannot be null");
    }

    protected int requireMaxCollectedErrors(int maximum) {
        if (maximum < 0) throw new IllegalArgumentException("maxCollectedErrors must be non-negative");
        return maximum;
    }

    protected ReadSummary summarizeRead(Consumer<Consumer<ReadResult<T>>> execution,
            BooleanSupplier stoppedEarly, Consumer<ReadResult<T>> consumer) {
        return ReaderExecutionSupport.summarizeRead(execution, stoppedEarly, consumer);
    }

    protected ReadReport collectReadReport(Consumer<Consumer<ReadResult<T>>> execution,
            BooleanSupplier stoppedEarly, int maximum) {
        return ReaderExecutionSupport.collectReadReport(execution, stoppedEarly, maximum);
    }

    protected ReadSummary summarizeReadWhile(Consumer<Predicate<ReadResult<T>>> execution,
            BooleanSupplier stoppedEarly, Predicate<ReadResult<T>> predicate) {
        return ReaderExecutionSupport.summarizeReadWhile(execution, stoppedEarly, predicate);
    }
}
