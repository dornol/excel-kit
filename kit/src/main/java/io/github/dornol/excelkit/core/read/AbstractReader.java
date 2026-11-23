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

    /**
     * Requires the values of the given columns, taken together, to be unique across the rows of the file: a row
     * whose key was already seen in an earlier row fails with a {@link RowError.Type#VALIDATION} error that names
     * the key and the file row where it first appeared, and marks the key cells in
     * {@link RowError#cellErrors()}. The first row with a key is not affected. Without this a program has to keep
     * its own {@code Set} of the keys.
     * <pre>{@code
     * reader.uniqueKey("Employee ID")              // one column
     * reader.uniqueKey("Department", "Employee ID") // the pair
     * }</pre>
     * The columns are given by header name, matched like the other header names of this reader (see
     * {@link #headerNormalizer(java.util.function.UnaryOperator)}); a header that is not in the file fails the read
     * before the first row. The cell values are compared as text, trimmed and case-sensitive, using the stored
     * number for a numeric Excel cell. A row whose key has an empty cell is not checked: that is for
     * {@code required()}. A row that fails for another reason still takes its key, and its duplicate is reported
     * together with the other errors. All keys seen are held in memory until the read ends, so bound a very large
     * file with {@link #maxRows(long)}. Calling this again replaces the key; no header turns it off.
     *
     * @param headers the header names of the key columns
     * @return this reader for chaining
     * @throws IllegalArgumentException if a header is blank
     * @since 0.29.0
     */
    public SELF uniqueKey(String... headers) {
        settings.uniqueKey(headers);
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

    /**
     * Builds the format-specific handler that reads the given stream. The stream stays caller-owned.
     *
     * @param inputStream the input
     * @return a handler that has not read anything yet
     */
    protected abstract AbstractReadHandler<T> createHandler(InputStream inputStream);

    /**
     * Builds the format-specific handler that reads the given file in place.
     *
     * @param path the file
     * @return a handler that has not read anything yet
     */
    protected abstract AbstractReadHandler<T> createHandler(Path path);

    /**
     * The exception for an {@link InputStreamSource} whose stream could not be opened or closed.
     *
     * @param cause what went wrong
     * @return the format's own read exception
     */
    protected abstract RuntimeException openFailure(IOException cause);

    // The terminal operations below are the same for every format; only the handler and the failure type differ.

    /** Opens a reusable execution session backed by a caller-owned stream. */
    public ReadSession<T> open(InputStream inputStream) {
        return createHandler(inputStream);
    }

    /** Opens a session whose path resource is owned by the session. */
    public ReadSession<T> open(Path path) {
        return createHandler(path);
    }

    /** Opens a session from a source; the source stream is closed after session initialization. */
    public ReadSession<T> open(InputStreamSource source) {
        requireInputSource(source);
        try (InputStream input = source.openStream()) {
            return createHandler(input);
        } catch (IOException e) {
            throw openFailure(e);
        }
    }

    /** Reads an input stream without closing it. */
    public void read(InputStream inputStream, Consumer<ReadResult<T>> consumer) {
        requireCallback(consumer, "consumer");
        executeInputStream(inputStream, this::createHandler, handler -> handler.read(consumer));
    }

    /** Reads all successful rows into a list and fails on the first mapping or validation error. */
    public List<T> readAll(InputStream inputStream) {
        List<T> rows = new ArrayList<>();
        readStrict(inputStream, rows::add);
        return List.copyOf(rows);
    }

    /**
     * Reads all successful rows from a path into an immutable list, failing on the first mapping or
     * validation error. The file is not modified or deleted.
     */
    public List<T> readAll(Path path) {
        List<T> rows = new ArrayList<>();
        readStrict(path, rows::add);
        return List.copyOf(rows);
    }

    /**
     * Reads all successful rows from an {@link InputStreamSource} into an immutable list, failing
     * on the first mapping or validation error. The source's stream is opened and closed here.
     */
    public List<T> readAll(InputStreamSource source) {
        List<T> rows = new ArrayList<>();
        readStrict(source, rows::add);
        return List.copyOf(rows);
    }

    /**
     * Reads an input stream, passing every row's {@link ReadResult} to the consumer, and returns a
     * {@link ReadSummary} with the row counts. The stream is not closed.
     */
    public ReadSummary readWithSummary(InputStream inputStream, Consumer<ReadResult<T>> consumer) {
        requireCallback(consumer, "consumer");
        return executeWithInputStream(inputStream, this::createHandler,
                handler -> summarize(handler, consumer));
    }

    /**
     * Reads a path, passing every row's {@link ReadResult} to the consumer, and returns a {@link
     * ReadSummary} with the row counts. The file is not modified or deleted.
     */
    public ReadSummary readWithSummary(Path path, Consumer<ReadResult<T>> consumer) {
        requireCallback(consumer, "consumer");
        return executeWithPath(path, this::createHandler,
                handler -> summarize(handler, consumer));
    }

    /**
     * Reads an {@link InputStreamSource}, passing every row's {@link ReadResult} to the consumer,
     * and returns a {@link ReadSummary} with the row counts. The source's stream is opened and
     * closed here.
     */
    public ReadSummary readWithSummary(InputStreamSource source, Consumer<ReadResult<T>> consumer) {
        requireCallback(consumer, "consumer");
        return executeFromSource(source, input -> readWithSummary(input, consumer),
                (message, error) -> openFailure(error));
    }

    /**
     * Reads an input stream without delivering rows, and returns a {@link ReadReport} with the row
     * counts and up to {@code maxCollectedErrors} failed rows. The stream is not closed.
     */
    public ReadReport readReport(InputStream inputStream, int maxCollectedErrors) {
        requireMaxCollectedErrors(maxCollectedErrors);
        return executeWithInputStream(inputStream, this::createHandler,
                handler -> report(handler, maxCollectedErrors));
    }

    /**
     * Reads a path without delivering rows, and returns a {@link ReadReport} with the row counts
     * and up to {@code maxCollectedErrors} failed rows. The file is not modified or deleted.
     */
    public ReadReport readReport(Path path, int maxCollectedErrors) {
        requireMaxCollectedErrors(maxCollectedErrors);
        return executeWithPath(path, this::createHandler,
                handler -> report(handler, maxCollectedErrors));
    }

    /**
     * Reads an {@link InputStreamSource} without delivering rows, and returns a {@link ReadReport}
     * with the row counts and up to {@code maxCollectedErrors} failed rows. The source's stream is
     * opened and closed here.
     */
    public ReadReport readReport(InputStreamSource source, int maxCollectedErrors) {
        requireMaxCollectedErrors(maxCollectedErrors);
        return executeFromSource(source, input -> readReport(input, maxCollectedErrors),
                (message, error) -> openFailure(error));
    }

    /**
     * Reads an input stream, passing each valid row to {@code onSuccess} and each failed row to
     * {@code onError}. The stream is not closed.
     */
    public void read(InputStream inputStream, Consumer<T> onSuccess, Consumer<RowError> onError) {
        requireCallback(onSuccess, "onSuccess");
        requireCallback(onError, "onError");
        executeInputStream(inputStream, this::createHandler,
                handler -> handler.read(onSuccess, onError));
    }

    /**
     * Reads an input stream, passing each valid row to the consumer, and throws
     * {@link io.github.dornol.excelkit.core.read.ReadAbortException ReadAbortException} at the
     * first row that fails mapping or validation. The stream is not closed.
     */
    public void readStrict(InputStream inputStream, Consumer<T> consumer) {
        requireCallback(consumer, "consumer");
        executeInputStream(inputStream, this::createHandler, handler -> handler.readStrict(consumer));
    }

    /**
     * Reads an input stream while the predicate returns true. Stopping early is normal, not an
     * error; the returned {@link ReadSummary} says whether it happened. The stream is not closed.
     */
    public ReadSummary readWhile(InputStream inputStream, Predicate<ReadResult<T>> predicate) {
        requireCallback(predicate, "predicate");
        return executeWithInputStream(inputStream, this::createHandler,
                handler -> readWhile(handler, predicate));
    }

    /** Reads directly from a caller-owned path without modifying or deleting it. */
    public void read(Path path, Consumer<ReadResult<T>> consumer) {
        requireCallback(consumer, "consumer");
        executePath(path, this::createHandler, handler -> handler.read(consumer));
    }

    /**
     * Reads a path, passing each valid row to {@code onSuccess} and each failed row to {@code
     * onError}. The file is not modified or deleted.
     */
    public void read(Path path, Consumer<T> onSuccess, Consumer<RowError> onError) {
        requireCallback(onSuccess, "onSuccess");
        requireCallback(onError, "onError");
        executePath(path, this::createHandler, handler -> handler.read(onSuccess, onError));
    }

    /**
     * Reads a path, passing each valid row to the consumer, and throws
     * {@link io.github.dornol.excelkit.core.read.ReadAbortException ReadAbortException} at the
     * first row that fails mapping or validation. The file is not modified or deleted.
     */
    public void readStrict(Path path, Consumer<T> consumer) {
        requireCallback(consumer, "consumer");
        executePath(path, this::createHandler, handler -> handler.readStrict(consumer));
    }

    /**
     * Reads a path while the predicate returns true. Stopping early is normal, not an error; the
     * returned {@link ReadSummary} says whether it happened. The file is not modified or deleted.
     */
    public ReadSummary readWhile(Path path, Predicate<ReadResult<T>> predicate) {
        requireCallback(predicate, "predicate");
        return executeWithPath(path, this::createHandler,
                handler -> readWhile(handler, predicate));
    }

    /**
     * Reads an {@link InputStreamSource}, passing every row's {@link ReadResult}, valid or failed,
     * to the consumer. The source's stream is opened and closed here.
     */
    public void read(InputStreamSource source, Consumer<ReadResult<T>> consumer) {
        requireCallback(consumer, "consumer");
        executeFromSource(source, input -> { read(input, consumer); return null; },
                (message, error) -> openFailure(error));
    }

    /**
     * Reads an {@link InputStreamSource}, passing each valid row to {@code onSuccess} and each
     * failed row to {@code onError}. The source's stream is opened and closed here.
     */
    public void read(InputStreamSource source, Consumer<T> onSuccess, Consumer<RowError> onError) {
        requireCallback(onSuccess, "onSuccess");
        requireCallback(onError, "onError");
        executeFromSource(source, input -> { read(input, onSuccess, onError); return null; },
                (message, error) -> openFailure(error));
    }

    /**
     * Reads an {@link InputStreamSource}, passing each valid row to the consumer, and throws
     * {@link io.github.dornol.excelkit.core.read.ReadAbortException ReadAbortException} at the
     * first row that fails mapping or validation. The source's stream is opened and closed here.
     */
    public void readStrict(InputStreamSource source, Consumer<T> consumer) {
        requireCallback(consumer, "consumer");
        executeFromSource(source, input -> { readStrict(input, consumer); return null; },
                (message, error) -> openFailure(error));
    }

    /**
     * Reads an {@link InputStreamSource} while the predicate returns true. Stopping early is
     * normal, not an error; the returned {@link ReadSummary} says whether it happened. The source's
     * stream is opened and closed here.
     */
    public ReadSummary readWhile(InputStreamSource source, Predicate<ReadResult<T>> predicate) {
        requireCallback(predicate, "predicate");
        return executeFromSource(source, input -> readWhile(input, predicate),
                (message, error) -> openFailure(error));
    }

    /** Runs the handler, passing every row's result to the consumer, and summarizes the read. */
    protected final ReadSummary summarize(AbstractReadHandler<T> handler, Consumer<ReadResult<T>> consumer) {
        return summarizeRead(handler::read, handler::wasStoppedEarly, consumer);
    }

    /** Runs the handler without delivering rows and reports the counts and up to the given number of failed rows. */
    protected final ReadReport report(AbstractReadHandler<T> handler, int maxCollectedErrors) {
        return collectReadReport(handler::read, handler::wasStoppedEarly, maxCollectedErrors);
    }

    private ReadSummary readWhile(AbstractReadHandler<T> handler, Predicate<ReadResult<T>> predicate) {
        return summarizeReadWhile(handler::readWhile, handler::wasStoppedEarly, predicate);
    }
}
