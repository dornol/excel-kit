package io.github.dornol.excelkit.core.read;

import io.github.dornol.excelkit.core.ProgressCallback;
import io.github.dornol.excelkit.core.io.InputStreamSource;
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
 * Shared reader configuration for {@link io.github.dornol.excelkit.excel.read.ExcelReader}
 * and {@link io.github.dornol.excelkit.csv.CsvReader}.
 * <p>
 * Contains column registration, progress, validation, and skip logic.
 * Format-specific configuration (sheet index, delimiter, charset, etc.) stays in subclasses.
 *
 * @param <T>    the row data type
 * @param <SELF> the concrete reader type, for fluent chaining
 * @author dhkim
 * @since 0.16.14
 */
@SuppressWarnings("unchecked")
public abstract class AbstractReader<T, SELF extends AbstractReader<T, SELF>> {

    protected final List<ReadColumn<T>> columns = new ArrayList<>();
    protected final @Nullable Supplier<T> instanceSupplier;
    protected final @Nullable Function<RowData, T> rowMapper;
    private ReaderOptions.Builder settings = ReaderOptions.builder();
    private boolean mapMode;
    private @Nullable Set<String> selectedMapColumns;

    protected AbstractReader(Supplier<T> instanceSupplier, @Nullable Validator validator) {
        this.instanceSupplier = Objects.requireNonNull(instanceSupplier, "instanceSupplier cannot be null");
        this.rowMapper = null;
        settings.validator(validator);
    }

    protected AbstractReader(Function<RowData, T> rowMapper, @Nullable Validator validator) {
        this.instanceSupplier = null;
        this.rowMapper = Objects.requireNonNull(rowMapper, "rowMapper cannot be null");
        settings.validator(validator);
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
        settings.validator(Objects.requireNonNull(validator, "validator cannot be null"));
        return self();
    }

    /**
     * Applies reusable common reader settings. Format-specific settings are unchanged.
     * <p>
     * The common settings given to this reader earlier (for example {@code headerRowIndex}) are replaced
     * by the options, so call this before chaining individual settings. The Bean Validation provider is
     * the exception: options that do not carry a validator keep the one already configured, such as the
     * one passed to {@code setter(...)} or {@code mapping(...)}.
     *
     * @param options the settings to apply
     * @return this reader for chaining
     */
    public SELF options(ReaderOptions options) {
        Validator previous = settings.configuredValidator();
        ReaderOptions.Builder applied = Objects.requireNonNull(options, "options cannot be null").toBuilder();
        if (applied.configuredValidator() == null) applied.validator(previous);
        settings = applied;
        return self();
    }

    protected final @Nullable Validator configuredValidator() {
        return settings.configuredValidator();
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
        settings.headerRowIndex(headerRowIndex);
        return self();
    }

    protected final int configuredHeaderRowIndex() {
        return settings.configuredHeaderRowIndex();
    }

    void addColumn(ReadColumn<T> column) {
        columns.add(column);
    }

    protected final void enableMapMode(@Nullable Set<String> selectedColumns) {
        mapMode = true;
        this.selectedMapColumns = selectedColumns == null ? null : new LinkedHashSet<>(selectedColumns);
    }

    protected final List<ReadColumn<T>> configuredColumns() {
        return List.copyOf(columns);
    }

    protected final @Nullable Set<String> configuredSelectedMapColumns() {
        return selectedMapColumns == null ? null : Set.copyOf(selectedMapColumns);
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
        settings.onProgress(interval, callback);
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
        settings.strictHeaders(enabled);
        return self();
    }

    /**
     * Alias for {@link #strictHeaders()}.
     */
    public SELF requireHeaders() {
        return strictHeaders();
    }

    /**
     * Sets duplicate header handling. Defaults to {@link DuplicateHeaderPolicy#FAIL}; blank headers are ignored.
     */
    public SELF duplicateHeaderPolicy(DuplicateHeaderPolicy policy) {
        settings.duplicateHeaderPolicy(policy);
        return self();
    }

    /**
     * Sets reader-scoped conversion settings for {@link CellData}.
     *
     * @since 0.19.0
     */
    public SELF cellConversion(CellConversionConfig config) {
        settings.cellConversion(config);
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
        settings.maxRows(maxRows);
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
        settings.skipBlankRows(enabled);
        return self();
    }

    /**
     * Stops reading after the given number of consecutive blank data rows.
     * Pass {@code 0} to disable.
     *
     * @since 0.19.0
     */
    public SELF stopAtBlankRows(int count) {
        settings.stopAtBlankRows(count);
        return self();
    }

    /** Limits failed rows; zero aborts on the first failure. Disabled by default. */
    public SELF maxErrors(long maxErrors) {
        settings.maxErrors(maxErrors);
        return self();
    }

    /** Normalizes headers before matching names and aliases. Original names remain available for diagnostics. */
    public SELF headerNormalizer(UnaryOperator<String> normalizer) {
        settings.headerNormalizer(normalizer);
        return self();
    }

    /**
     * Matches headers through one of the {@link HeaderPolicy} presets (trim, case-insensitive,
     * Unicode-normalized); shorthand for {@link #headerNormalizer(UnaryOperator)} with the preset's
     * normalizer.
     *
     * @param policy the preset to match headers with
     * @return this reader for chaining
     */
    public SELF headerPolicy(HeaderPolicy policy) {
        return headerNormalizer(Objects.requireNonNull(policy, "policy cannot be null").normalizer());
    }

    /**
     * Bounds the input, sheets, columns, cell size and cell count. No limits apply by default, so a
     * reader of untrusted files should set some; {@link ReadLimits#UPLOAD} is a starting point.
     *
     * @param limits the limits to enforce
     * @return this reader for chaining
     */
    public SELF limits(ReadLimits limits) {
        settings.limits(limits);
        return self();
    }

    /**
     * Lets another thread stop the read. The token is checked before every row; once it reports a
     * cancellation the read ends normally without an exception, the rows already emitted stay valid,
     * the summary reports the read as stopped early and the final {@link ReadProgress} has
     * {@code cancelled} set.
     *
     * @param token the cancellation signal to poll
     * @return this reader for chaining
     */
    public SELF cancellationToken(CancellationToken token) {
        settings.cancellationToken(token);
        return self();
    }

    /**
     * Reports progress as {@link ReadProgress} snapshots (rows processed, succeeded and failed, elapsed
     * time) every {@code interval} rows, plus one final event with {@code completed} set when the read
     * ends. The callback runs on the reading thread. Unlike {@link #onProgress(int, ProgressCallback)}
     * it also counts failed rows separately and sees completion and cancellation.
     *
     * @param interval rows between two events; must be positive
     * @param callback receives each snapshot
     * @return this reader for chaining
     */
    public SELF onReadProgress(int interval, ReadProgressCallback callback) {
        settings.onReadProgress(interval, callback);
        return self();
    }

    /**
     * Sets the Excel content restrictions. The default {@link ReadSecurityPolicy#DEFAULT} allows formulas
     * and external links; use {@link ReadSecurityPolicy#STRICT} for files from untrusted sources.
     *
     * @param policy the policy to enforce
     * @return this reader for chaining
     */
    public SELF securityPolicy(ReadSecurityPolicy policy) {
        settings.securityPolicy(policy);
        return self();
    }

    protected ReadOptions snapshotReadOptions() {
        return settings.toReadOptions();
    }

    /** Caps reads/skips at the remaining budget plus one byte to detect overflow. */
    protected InputStream limitInput(InputStream input) {
        Objects.requireNonNull(input, "input cannot be null");
        long maximum = settings.toReadOptions().limits().maxInputBytes();
        if (maximum < 0) return input;
        return new FilterInputStream(input) {
            private long count;
            private @Nullable ReadLimitExceededException exceeded;

            private void checkExceeded() {
                if (exceeded != null) throw exceeded;
            }

            private void counted(long amount) {
                if (amount <= 0) return;
                if (amount > maximum - count) {
                    // The observed count saturates if maxInputBytes itself is Long.MAX_VALUE.
                    long observed = count > Long.MAX_VALUE - amount ? Long.MAX_VALUE : count + amount;
                    exceeded = new ReadLimitExceededException(ReadLimitExceededException.Limit.INPUT_BYTES,
                            maximum, observed);
                    throw exceeded;
                }
                count += amount;
            }
            @Override public int read() throws IOException {
                checkExceeded();
                int value = in.read();
                if (value >= 0) counted(1);
                return value;
            }
            @Override public int read(byte[] bytes, int offset, int length) throws IOException {
                Objects.checkFromIndexSize(offset, length, bytes.length);
                if (length == 0) return 0;
                checkExceeded();
                long remaining = maximum - count;
                int requested = remaining >= length ? length : (int) remaining + 1;
                int read = in.read(bytes, offset, requested);
                counted(read);
                return read;
            }
            @Override public long skip(long length) throws IOException {
                if (length <= 0) return 0;
                checkExceeded();
                long remaining = maximum - count;
                long requested = remaining >= length ? length : remaining + 1;
                long skipped = in.skip(requested);
                counted(skipped);
                return skipped;
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
