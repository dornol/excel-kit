package io.github.dornol.excelkit.core.read;

import io.github.dornol.excelkit.core.Cursor;
import io.github.dornol.excelkit.core.ExcelKitException;
import io.github.dornol.excelkit.core.io.TempResourceContainer;
import io.github.dornol.excelkit.core.io.TempResourceCreator;
import jakarta.validation.Validator;
import org.jspecify.annotations.Nullable;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.IOException;
import java.io.InputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.StandardCopyOption;
import java.util.List;
import java.util.Objects;
import java.util.Map;
import java.util.Set;
import java.util.UUID;
import java.util.concurrent.atomic.AtomicLong;
import java.util.function.BiConsumer;
import java.util.function.Consumer;
import java.util.function.Function;
import java.util.function.IntFunction;
import java.util.function.IntUnaryOperator;
import java.util.function.Predicate;
import java.util.function.Supplier;

/**
 * Abstract base class for file read handlers (Excel, CSV).
 * <p>
 * Provides common functionality including:
 * <ul>
 *     <li>Constructor parameter validation</li>
 *     <li>Temporary file initialization from an InputStream</li>
 *     <li>Optional Bean Validation support</li>
 * </ul>
 *
 * @param <T> The target row data type to map each row into
 * @author dhkim
 * @since 0.2.0
 */
public abstract class AbstractReadHandler<T> extends TempResourceContainer implements ReadSession<T> {
    private static final Logger log = LoggerFactory.getLogger(AbstractReadHandler.class);

    /** Supplier for creating new row instances (setter mode). */
    protected final @Nullable Supplier<T> instanceSupplier;
    /** Function for mapping row data to instances (mapping mode). */
    protected final @Nullable Function<RowData, T> rowMapper;
    /** Optional bean validator for row validation. */
    protected final @Nullable Validator validator;
    protected final @Nullable Set<String> selectedMapColumns;
    protected final ReadOptions readOptions;
    private final ReadLifecycle lifecycle = new ReadLifecycle();
    private final ReadRowPolicy rowPolicy;
    protected boolean stoppedEarly;

    // initTempFile only uses this class's own temp file accessors, and the only subclasses are the
    // final CSV and Excel handlers of this library, so nothing observes a half-built object.
    @SuppressWarnings("this-escape")
    protected AbstractReadHandler(InputStream input, @Nullable Supplier<T> supplier,
            @Nullable Function<RowData, T> mapper, @Nullable Validator validator, String extension,
            @Nullable Set<String> selectedColumns, ReadOptions readOptions) {
        if ((supplier == null) == (mapper == null))
            throw new IllegalArgumentException("Exactly one of supplier or mapper is required");
        this.instanceSupplier = supplier;
        this.rowMapper = mapper;
        this.validator = validator;
        this.selectedMapColumns = selectedColumns;
        this.readOptions = Objects.requireNonNull(readOptions, "readOptions cannot be null");
        this.rowPolicy = new ReadRowPolicy(this.readOptions);
        initTempFile(Objects.requireNonNull(input, "input cannot be null"), extension);
    }

    protected CellData cellData(int columnIndex, @Nullable String formattedValue) {
        return cellData(columnIndex, formattedValue, null);
    }

    protected CellData cellData(int columnIndex, @Nullable String formattedValue,
                                @Nullable String rawNumericValue) {
        if (readOptions.limits().maxColumns() >= 0
                && columnIndex >= readOptions.limits().maxColumns()) {
            throw new ReadLimitExceededException(ReadLimitExceededException.Limit.COLUMNS,
                    readOptions.limits().maxColumns(), columnIndex + 1L);
        }
        if (formattedValue != null && readOptions.limits().maxCellCharacters() >= 0
                && formattedValue.length() > readOptions.limits().maxCellCharacters()) {
            throw new ReadLimitExceededException(ReadLimitExceededException.Limit.CELL_CHARACTERS,
                    readOptions.limits().maxCellCharacters(), formattedValue.length());
        }
        return new CellData(columnIndex, formattedValue, readOptions.cellConversionConfig(), rawNumericValue);
    }

    protected final void validateColumnCount(int columnCount) {
        if (readOptions.limits().maxColumns() >= 0
                && columnCount > readOptions.limits().maxColumns()) {
            throw new ReadLimitExceededException(ReadLimitExceededException.Limit.COLUMNS,
                    readOptions.limits().maxColumns(), columnCount);
        }
    }

    protected List<String> rawValues(List<CellData> cells) {
        return cells.stream().map(CellData::sourceValue).toList();
    }

    protected static final int PROCESS_ROW = 0;
    protected static final int SKIP_ROW = 1;
    protected static final int STOP_READING = 2;

    /** Evaluates common row handling without exposing the internal policy type. */
    protected final int evaluateRow(List<String> rawValues) {
        ReadRowPolicy.Decision decision = rowPolicy.evaluate(rawValues);
        return decision == ReadRowPolicy.Decision.PROCESS ? PROCESS_ROW
                : decision == ReadRowPolicy.Decision.SKIP ? SKIP_ROW : STOP_READING;
    }

    protected final void recordEmittedRow() {
        rowPolicy.recordEmitted();
    }

    protected final void notifyLegacyProgress(long processedRows, @Nullable Cursor cursor) {
        if (readOptions.progressCallback() != null && readOptions.progressInterval() > 0
                && processedRows % readOptions.progressInterval() == 0) {
            readOptions.progressCallback().onProgress(processedRows, cursor);
        }
    }

    /**
     * Validates that headerRowIndex is non-negative.
     *
     * @param headerRowIndex the header row index to validate
     */
    protected static void validateHeaderRowIndex(int headerRowIndex) {
        if (headerRowIndex < 0) {
            throw new IllegalArgumentException("headerRowIndex must be non-negative");
        }
    }

    /**
     * Validates that the columns list is non-null and non-empty.
     *
     * @param columns the column list to validate
     */
    protected static void validateColumns(List<?> columns) {
        if (columns == null || columns.isEmpty()) {
            throw new IllegalArgumentException("Columns cannot be null or empty");
        }
    }

    /**
     * The message of a failed read.
     * <p>
     * When the cause is one of the library's own checks on the file (a missing, duplicate or unusable
     * header, ...), its message says what is wrong with the file, so it goes into the message instead
     * of only into the cause chain. Callers that report just {@code getMessage()} (an API error body,
     * a log line) would otherwise see nothing but the generic text. Any other cause keeps the generic
     * message: it describes the environment or the parser, not the file's contents.
     *
     * @param generic the generic message, such as {@code "Failed to read CSV"}
     * @param cause   what went wrong
     * @return {@code generic}, followed by the cause's message for a library check
     */
    protected static String readFailureMessage(String generic, Exception cause) {
        String detail = cause instanceof ExcelKitException ? cause.getMessage() : null;
        return detail == null || detail.isBlank() ? generic : generic + ": " + detail;
    }

    /**
     * Runs a subclass constructor's validation. The superclass constructor has already copied the
     * input to a temp file, and a handler whose constructor throws is never returned, so nobody
     * could close it: release the temp file here before the exception propagates.
     */
    protected final void validateOrClose(Runnable validation) {
        try {
            validation.run();
        } catch (RuntimeException e) {
            close();
            throw e;
        }
    }

    private void initTempFile(InputStream inputStream, String extension) {
        try {
            setTempDir(TempResourceCreator.createTempDirectory());
            setTempFile(TempResourceCreator.createTempFile(getTempDir(), UUID.randomUUID().toString(), extension));
            Files.copy(inputStream, getTempFile(), StandardCopyOption.REPLACE_EXISTING);
        } catch (IOException e) {
            // Clean up partially created temp resources before rethrowing
            close();
            throw new ExcelKitException("Failed to initialize temporary file", e);
        } catch (RuntimeException e) {
            close();
            throw e;
        }
    }

    protected final void useExternalInput(Path path) {
        if (!Files.isRegularFile(Objects.requireNonNull(path, "path cannot be null"))) {
            close();
            throw new ExcelKitException("Input path is not a regular file: " + path);
        }
        useExternalFile(path);
    }

    /**
     * Reads the file and invokes the given consumer for each row result.
     *
     * @param consumer Callback to receive parsed and validated row results
     */
    public abstract void read(Consumer<ReadResult<T>> consumer);

    /**
     * Tells whether the read ended because a {@link #readWhile} predicate said so rather than because the
     * input was exhausted. It is only meaningful after a read has finished.
     *
     * @return {@code true} if the read stopped early
     * @since 0.28.0
     */
    public final boolean wasStoppedEarly() {
        return stoppedEarly;
    }

    /**
     * Reads until the callback returns false. Returning false is normal completion; a predicate that
     * throws stops the read and its exception is rethrown once the handler is closed.
     */
    public void readWhile(Predicate<ReadResult<T>> predicate) {
        Objects.requireNonNull(predicate, "predicate cannot be null");
        java.util.concurrent.atomic.AtomicReference<RuntimeException> failure =
                new java.util.concurrent.atomic.AtomicReference<>();
        read(result -> {
            boolean proceed;
            try {
                proceed = predicate.test(result);
            } catch (RuntimeException e) {
                failure.set(e);
                throw new ReadStoppedException();
            }
            if (!proceed) throw new ReadStoppedException();
        });
        if (failure.get() != null) throw failure.get();
    }

    /** Validates input against the immutable configuration snapshot for this session. */
    protected final void validateReadOptions() {
        if (readOptions.limits().maxInputBytes() >= 0) {
            try {
                long size = Files.size(Objects.requireNonNull(getTempFile()));
                if (size > readOptions.limits().maxInputBytes()) {
                    close();
                    throw new ReadLimitExceededException(ReadLimitExceededException.Limit.INPUT_BYTES,
                            readOptions.limits().maxInputBytes(), size);
                }
            } catch (IOException e) {
                close();
                throw new ExcelKitException("Failed to inspect input size", e);
            }
        }
    }

    protected Consumer<ReadResult<T>> guardedConsumer(Consumer<ReadResult<T>> consumer) {
        Objects.requireNonNull(consumer, "consumer cannot be null");
        AtomicLong errors = new AtomicLong();
        return result -> {
            if (readOptions.cancellationToken().isCancellationRequested()) {
                throw new ReadStoppedException();
            }
            if (!result.success()) {
                lifecycle.record(false);
                long count = errors.incrementAndGet();
                if (readOptions.maxErrors() >= 0 && count > readOptions.maxErrors()) {
                    throw new ReadAbortException("Maximum read errors exceeded: " + readOptions.maxErrors(),
                            ReadAbortReason.MAX_ERRORS_EXCEEDED, readOptions.maxErrors(), count);
                }
            } else lifecycle.record(true);
            consumer.accept(result);
        };
    }

    protected void notifyReadProgress(long processedRows, int sheetIndex, long totalRows) {
        lifecycle.progress(processedRows, sheetIndex, totalRows, readOptions.readProgressCallback());
    }

    protected void notifyReadCompletion(int sheetIndex, long totalRows) {
        lifecycle.complete(sheetIndex, totalRows, readOptions.cancellationToken().isCancellationRequested(),
                readOptions.readProgressCallback());
    }

    /**
     * Reads the file and invokes the given consumer only for successfully parsed rows.
     * If any row fails validation or mapping, a {@link ReadAbortException} is thrown immediately.
     *
     * @param consumer Callback to receive successfully parsed row data
     * @throws ReadAbortException if any row fails validation or mapping
     */
    public void readStrict(Consumer<T> consumer) {
        ReadResultCallbacks.strict(this::read, consumer);
    }

    /**
     * Reads the file and routes each row to one of two callbacks.
     * <p>
     * Successful rows are passed to {@code onSuccess}; failed rows (validation,
     * mapping, or cell-conversion errors) are described in a {@link RowError}
     * and passed to {@code onError}. The library buffers nothing — memory
     * management is entirely up to the caller. Reading continues past errors.
     * To abort, throw from {@code onError}.
     *
     * @param onSuccess callback for each successfully parsed and validated row
     * @param onError   callback for each failed row
     * @since 0.16.12
     */
    public void read(Consumer<T> onSuccess, Consumer<RowError> onError) {
        ReadResultCallbacks.route(this::read, onSuccess, onError);
    }

    /**
     * Marks this handler as consumed. Read handlers own temporary resources and
     * can only be consumed once.
     */
    protected void markConsumed() {
        lifecycle.markConsumed();
    }

    /**
     * Validates the given instance using Bean Validation (if a validator is configured).
     *
     * @param instance The object to validate
     * @param messages A mutable list to collect violation messages
     * @return {@code true} if valid or no validator is configured, {@code false} if violations exist
     */
    protected boolean validateIfNeeded(T instance, List<String> messages) {
        return RowMappingSupport.validate(instance, validator, messages);
    }

    /**
     * Resolves column indices based on header aliases, columnIndex, or positional order.
     *
     * @param columnCount     number of columns to resolve
     * @param headerAliasesFn function to get accepted header aliases for column i
     * @param columnIndexFn   function to get explicit columnIndex for column i (-1 if not set)
     * @param headerNames     the header names from the file
     * @param errorPrefix     prefix for error messages (e.g., "sheet" or "CSV")
     * @return resolved index array
     */
    protected int[] resolveColumnIndices(int columnCount,
                                          IntFunction<List<String>> headerAliasesFn,
                                          IntUnaryOperator columnIndexFn,
                                          List<String> headerNames, String errorPrefix) {
        return headerResolver().resolve(columnCount, headerAliasesFn, columnIndexFn, headerNames, errorPrefix);
    }

    /**
     * Builds a header-to-index map using this handler's duplicate header policy.
     */
    protected Map<String, Integer> buildHeaderIndexMap(List<String> headerNames, String errorPrefix) {
        return headerResolver().index(headerNames, errorPrefix);
    }

    /**
     * Validates selected map-mode columns in strict mode.
     */
    protected void validateSelectedMapColumns(Map<String, Integer> headerIndexMap, List<String> headerNames, String errorPrefix) {
        headerResolver().validateSelected(selectedMapColumns, headerIndexMap, headerNames, errorPrefix);
    }

    protected String normalizeHeader(String header) {
        return headerResolver().normalize(header);
    }

    private HeaderResolver headerResolver() {
        return new HeaderResolver(readOptions.strictHeaders(), readOptions.duplicateHeaderPolicy(),
                readOptions.headerNormalizer(), readOptions.limits().maxColumns());
    }

    /**
     * Maps a single column value to the instance, handling exceptions.
     *
     * @param setter      The setter to apply
     * @param instance    The target object
     * @param cellData    The cell data to set
     * @param columnIndex The column index (for error reporting)
     * @param headerNames The header names (for error reporting)
     * @param messages    A mutable list to collect error messages
     * @return {@code true} if mapping succeeded, {@code false} if an exception occurred
     */
    protected boolean mapColumn(BiConsumer<T, CellData> setter, T instance, CellData cellData,
                                int columnIndex, List<String> headerNames, List<String> messages) {
        return mapColumn(setter, instance, cellData, columnIndex, headerNames, messages, null);
    }

    protected boolean mapColumn(BiConsumer<T, CellData> setter, T instance, CellData cellData,
                                int columnIndex, List<String> headerNames, List<String> messages,
                                @Nullable List<CellError> cellErrors) {
        return RowMappingSupport.map(setter, instance, cellData, columnIndex, headerNames, messages, cellErrors);
    }

    /**
     * Maps a single column value to the instance, with required-field validation.
     *
     * @param column      The column definition (includes required flag)
     * @param instance    The target object
     * @param cellData    The cell data to set
     * @param columnIndex The column index (for error reporting)
     * @param headerNames The header names (for error reporting)
     * @param messages    A mutable list to collect error messages
     * @return {@code true} if mapping succeeded, {@code false} if an error occurred
     */
    protected boolean mapColumn(ReadColumn<T> column, T instance, CellData cellData,
                                int columnIndex, List<String> headerNames, List<String> messages) {
        return mapColumn(column, instance, cellData, columnIndex, headerNames, messages, null);
    }

    protected boolean mapColumn(ReadColumn<T> column, T instance, CellData cellData,
                                int columnIndex, List<String> headerNames, List<String> messages,
                                @Nullable List<CellError> cellErrors) {
        return RowMappingSupport.map(column, instance, cellData, columnIndex, headerNames, messages, cellErrors);
    }

    /**
     * Maps a row using the row mapper function (mapping mode).
     * Catches exceptions from the mapper and wraps them in a failed {@link ReadResult}.
     *
     * @param rowData The row data to map
     * @return A {@link ReadResult} containing the mapped instance or error messages
     */
    protected ReadResult<T> mapWithRowMapper(RowData rowData) {
        return mapWithRowMapper(rowData, -1, List.of());
    }

    /**
     * Maps a row using the row mapper function and records its physical file row number.
     */
    protected ReadResult<T> mapWithRowMapper(RowData rowData, long fileRowNum) {
        return mapWithRowMapper(rowData, fileRowNum, List.of());
    }

    protected ReadResult<T> mapWithRowMapper(RowData rowData, long fileRowNum, List<String> rawValues) {
        if (rowMapper == null) {
            throw new IllegalStateException("rowMapper must not be null in mapping mode");
        }
        return RowMappingSupport.mapRow(rowMapper, rowData, validator, fileRowNum, rawValues);
    }
}
