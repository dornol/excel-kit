package io.github.dornol.excelkit.core;

import jakarta.validation.ConstraintViolation;
import jakarta.validation.Validator;
import org.jspecify.annotations.Nullable;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.IOException;
import java.io.InputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.StandardCopyOption;
import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.Set;
import java.util.UUID;
import java.util.concurrent.atomic.AtomicBoolean;
import java.util.concurrent.atomic.AtomicLong;
import java.util.function.Consumer;
import java.util.function.Function;
import java.util.function.IntFunction;
import java.util.function.IntUnaryOperator;
import java.util.function.Predicate;
import java.util.function.Supplier;
import java.util.function.UnaryOperator;

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
 * @since 2025-07-19
 */
public abstract class AbstractReadHandler<T> extends TempResourceContainer {
    private static final Logger log = LoggerFactory.getLogger(AbstractReadHandler.class);

    /** Supplier for creating new row instances (setter mode). */
    protected final @Nullable Supplier<T> instanceSupplier;
    /** Function for mapping row data to instances (mapping mode). */
    protected final @Nullable Function<RowData, T> rowMapper;
    /** Optional bean validator for row validation. */
    protected final @Nullable Validator validator;
    protected boolean strictHeaders;
    protected DuplicateHeaderPolicy duplicateHeaderPolicy;
    protected final @Nullable Set<String> selectedMapColumns;
    protected @Nullable CellConversionConfig cellConversionConfig;
    protected long maxRows;
    protected boolean skipBlankRows;
    protected int stopAtBlankRows;
    protected long maxErrors = -1;
    protected UnaryOperator<String> headerNormalizer = UnaryOperator.identity();
    protected ReadLimits limits = ReadLimits.UNLIMITED;
    protected CancellationToken cancellationToken = CancellationToken.NONE;
    protected @Nullable ReadProgressCallback readProgressCallback;
    protected ReadSecurityPolicy securityPolicy = ReadSecurityPolicy.DEFAULT;
    private final ReadLifecycle lifecycle = new ReadLifecycle();
    protected boolean stoppedEarly;

    protected AbstractReadHandler(InputStream input, @Nullable Supplier<T> supplier,
            @Nullable Function<RowData,T> mapper, @Nullable Validator validator, String extension,
            ReadOptions options, @Nullable Set<String> selectedColumns) {
        if ((supplier == null) == (mapper == null))
            throw new IllegalArgumentException("Exactly one of supplier or mapper is required");
        this.instanceSupplier = supplier;
        this.rowMapper = mapper;
        this.validator = validator;
        this.strictHeaders = options.strictHeaders();
        this.duplicateHeaderPolicy = options.duplicateHeaderPolicy();
        this.selectedMapColumns = selectedColumns;
        this.cellConversionConfig = options.cellConversionConfig();
        this.maxRows = options.maxRows();
        this.skipBlankRows = options.skipBlankRows();
        this.stopAtBlankRows = options.stopAtBlankRows();
        initTempFile(java.util.Objects.requireNonNull(input, "input cannot be null"), extension);
    }

    /**
     * Constructs a read handler in setter mode by validating inputs and initializing a temporary file.
     *
     * @param inputStream      The input stream of the uploaded file
     * @param instanceSupplier A supplier to instantiate new row objects
     * @param validator        Optional bean validator for validating mapped instances
     * @param extension        File extension for the temporary file (e.g., ".xlsx", ".csv")
     */
    protected AbstractReadHandler(InputStream inputStream, Supplier<T> instanceSupplier, @Nullable Validator validator, String extension) {
        this(inputStream, instanceSupplier, validator, extension, false, DuplicateHeaderPolicy.FIRST);
    }

    protected AbstractReadHandler(InputStream inputStream, Supplier<T> instanceSupplier, @Nullable Validator validator,
                                  String extension, boolean strictHeaders, DuplicateHeaderPolicy duplicateHeaderPolicy) {
        this(inputStream, instanceSupplier, validator, extension, strictHeaders, duplicateHeaderPolicy, null);
    }

    protected AbstractReadHandler(InputStream inputStream, Supplier<T> instanceSupplier, @Nullable Validator validator,
                                  String extension, boolean strictHeaders, DuplicateHeaderPolicy duplicateHeaderPolicy,
                                  @Nullable Set<String> selectedMapColumns) {
        this(inputStream, instanceSupplier, validator, extension, strictHeaders, duplicateHeaderPolicy,
                selectedMapColumns, null);
    }

    protected AbstractReadHandler(InputStream inputStream, Supplier<T> instanceSupplier, @Nullable Validator validator,
                                  String extension, boolean strictHeaders, DuplicateHeaderPolicy duplicateHeaderPolicy,
                                  @Nullable Set<String> selectedMapColumns,
                                  @Nullable CellConversionConfig cellConversionConfig) {
        this(inputStream, instanceSupplier, validator, extension, strictHeaders, duplicateHeaderPolicy,
                selectedMapColumns, cellConversionConfig, -1, false, 0);
    }

    protected AbstractReadHandler(InputStream inputStream, Supplier<T> instanceSupplier, @Nullable Validator validator,
                                  String extension, boolean strictHeaders, DuplicateHeaderPolicy duplicateHeaderPolicy,
                                  @Nullable Set<String> selectedMapColumns,
                                  @Nullable CellConversionConfig cellConversionConfig,
                                  long maxRows, boolean skipBlankRows, int stopAtBlankRows) {
        if (inputStream == null) {
            throw new IllegalArgumentException("InputStream cannot be null");
        }
        if (instanceSupplier == null) {
            throw new IllegalArgumentException("Instance supplier cannot be null");
        }
        this.instanceSupplier = instanceSupplier;
        this.rowMapper = null;
        this.validator = validator;
        this.strictHeaders = strictHeaders;
        this.duplicateHeaderPolicy = java.util.Objects.requireNonNull(duplicateHeaderPolicy, "duplicateHeaderPolicy cannot be null");
        this.selectedMapColumns = selectedMapColumns;
        this.cellConversionConfig = cellConversionConfig;
        this.maxRows = maxRows;
        this.skipBlankRows = skipBlankRows;
        this.stopAtBlankRows = stopAtBlankRows;
        initTempFile(inputStream, extension);
    }

    /**
     * Constructs a read handler in mapping mode for immutable object construction.
     *
     * @param inputStream The input stream of the uploaded file
     * @param rowMapper   A function that creates an instance from a {@link RowData}
     * @param validator   Optional bean validator for validating mapped instances
     * @param extension   File extension for the temporary file (e.g., ".xlsx", ".csv")
     */
    protected AbstractReadHandler(InputStream inputStream, Function<RowData, T> rowMapper, @Nullable Validator validator, String extension) {
        this(inputStream, rowMapper, validator, extension, false, DuplicateHeaderPolicy.FIRST);
    }

    protected AbstractReadHandler(InputStream inputStream, Function<RowData, T> rowMapper, @Nullable Validator validator,
                                  String extension, boolean strictHeaders, DuplicateHeaderPolicy duplicateHeaderPolicy) {
        this(inputStream, rowMapper, validator, extension, strictHeaders, duplicateHeaderPolicy, null);
    }

    protected AbstractReadHandler(InputStream inputStream, Function<RowData, T> rowMapper, @Nullable Validator validator,
                                  String extension, boolean strictHeaders, DuplicateHeaderPolicy duplicateHeaderPolicy,
                                  @Nullable Set<String> selectedMapColumns) {
        this(inputStream, rowMapper, validator, extension, strictHeaders, duplicateHeaderPolicy,
                selectedMapColumns, null);
    }

    protected AbstractReadHandler(InputStream inputStream, Function<RowData, T> rowMapper, @Nullable Validator validator,
                                  String extension, boolean strictHeaders, DuplicateHeaderPolicy duplicateHeaderPolicy,
                                  @Nullable Set<String> selectedMapColumns,
                                  @Nullable CellConversionConfig cellConversionConfig) {
        this(inputStream, rowMapper, validator, extension, strictHeaders, duplicateHeaderPolicy,
                selectedMapColumns, cellConversionConfig, -1, false, 0);
    }

    protected AbstractReadHandler(InputStream inputStream, Function<RowData, T> rowMapper, @Nullable Validator validator,
                                  String extension, boolean strictHeaders, DuplicateHeaderPolicy duplicateHeaderPolicy,
                                  @Nullable Set<String> selectedMapColumns,
                                  @Nullable CellConversionConfig cellConversionConfig,
                                  long maxRows, boolean skipBlankRows, int stopAtBlankRows) {
        if (inputStream == null) {
            throw new IllegalArgumentException("InputStream cannot be null");
        }
        if (rowMapper == null) {
            throw new IllegalArgumentException("Row mapper cannot be null");
        }
        this.instanceSupplier = null;
        this.rowMapper = rowMapper;
        this.validator = validator;
        this.strictHeaders = strictHeaders;
        this.duplicateHeaderPolicy = java.util.Objects.requireNonNull(duplicateHeaderPolicy, "duplicateHeaderPolicy cannot be null");
        this.selectedMapColumns = selectedMapColumns;
        this.cellConversionConfig = cellConversionConfig;
        this.maxRows = maxRows;
        this.skipBlankRows = skipBlankRows;
        this.stopAtBlankRows = stopAtBlankRows;
        initTempFile(inputStream, extension);
    }

    protected CellData cellData(int columnIndex, @Nullable String formattedValue) {
        if (formattedValue != null && limits.maxCellCharacters() >= 0
                && formattedValue.length() > limits.maxCellCharacters()) {
            throw new ReadLimitExceededException(ReadLimitExceededException.Limit.CELL_CHARACTERS,
                    limits.maxCellCharacters(), formattedValue.length());
        }
        return new CellData(columnIndex, formattedValue, cellConversionConfig);
    }

    protected boolean isBlankValues(List<String> values) {
        return values.stream().allMatch(value -> value == null || value.isBlank());
    }

    protected List<String> rawValues(List<CellData> cells) {
        return cells.stream().map(CellData::formattedValue).toList();
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
        if (!Files.isRegularFile(java.util.Objects.requireNonNull(path, "path cannot be null"))) {
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

    /** Reads until the callback returns false. Returning false is normal completion. */
    public abstract void readWhile(Predicate<ReadResult<T>> predicate);

    /** Applies the immutable configuration snapshot for this one-shot session. */
    public AbstractReadHandler<T> options(ReadOptions options) {
        this.strictHeaders = options.strictHeaders();
        this.duplicateHeaderPolicy = options.duplicateHeaderPolicy();
        this.cellConversionConfig = options.cellConversionConfig();
        this.maxRows = options.maxRows();
        this.skipBlankRows = options.skipBlankRows();
        this.stopAtBlankRows = options.stopAtBlankRows();
        this.maxErrors = options.maxErrors();
        this.headerNormalizer = options.headerNormalizer();
        this.limits = options.limits();
        this.cancellationToken = options.cancellationToken();
        this.readProgressCallback = options.readProgressCallback();
        this.securityPolicy = options.securityPolicy();
        if (limits.maxInputBytes() >= 0) {
            try {
                if (Files.size(java.util.Objects.requireNonNull(getTempFile())) > limits.maxInputBytes()) {
                    close();
                    throw new ReadLimitExceededException(ReadLimitExceededException.Limit.INPUT_BYTES,
                            limits.maxInputBytes(), Files.size(java.util.Objects.requireNonNull(getTempFile())));
                }
            } catch (java.io.IOException e) {
                close();
                throw new ExcelKitException("Failed to inspect input size", e);
            }
        }
        return this;
    }

    protected Consumer<ReadResult<T>> guardedConsumer(Consumer<ReadResult<T>> consumer) {
        AtomicLong errors = new AtomicLong();
        return result -> {
            if (cancellationToken.isCancellationRequested()) {
                throw new ReadStoppedException();
            }
            if (!result.success()) {
                lifecycle.record(false);
                long count = errors.incrementAndGet();
                if (maxErrors >= 0 && count > maxErrors) {
                    throw new ReadAbortException("Maximum read errors exceeded: " + maxErrors,
                            ReadAbortReason.MAX_ERRORS_EXCEEDED, maxErrors, count);
                }
            } else lifecycle.record(true);
            consumer.accept(result);
        };
    }

    protected void notifyReadProgress(long processedRows, int sheetIndex, long totalRows) {
        lifecycle.progress(processedRows, sheetIndex, totalRows, readProgressCallback);
    }

    protected void notifyReadCompletion(int sheetIndex, long totalRows) {
        lifecycle.complete(sheetIndex, totalRows, cancellationToken.isCancellationRequested(), readProgressCallback);
    }

    /**
     * Reads the file and invokes the given consumer only for successfully parsed rows.
     * If any row fails validation or mapping, a {@link ReadAbortException} is thrown immediately.
     *
     * @param consumer Callback to receive successfully parsed row data
     * @throws ReadAbortException if any row fails validation or mapping
     */
    public void readStrict(Consumer<T> consumer) {
        AtomicLong rowNum = new AtomicLong(0);
        read(result -> {
            long row = rowNum.incrementAndGet();
            if (!result.success()) {
                String detail = (result.messages() != null && !result.messages().isEmpty())
                        ? String.join("; ", result.messages()) : "Unknown error";
                throw new ReadAbortException("Row " + row + " read failed: " + detail);
            }
            consumer.accept(result.data());
        });
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
        AtomicLong rowNum = new AtomicLong(0);
        read(result -> {
            long n = rowNum.incrementAndGet();
            if (result.success()) {
                onSuccess.accept(result.data());
            } else {
                List<String> msgs = result.messages() != null ? result.messages() : List.of();
                RowError.Type type = result.cause() != null ? RowError.Type.MAPPING : RowError.Type.VALIDATION;
                onError.accept(new RowError(n, result.fileRowNum(), type, msgs, result.cause(),
                        result.cellErrors(), result.rawValues()));
            }
        });
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
        if (validator == null) {
            return true;
        }

        Set<ConstraintViolation<T>> violations = validator.validate(instance);
        if (violations.isEmpty()) return true;

        violations.stream()
                .map(ConstraintViolation::getMessage)
                .forEach(messages::add);

        return false;
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
        return new HeaderResolver(strictHeaders, duplicateHeaderPolicy, headerNormalizer, limits.maxColumns());
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
    protected boolean mapColumn(java.util.function.BiConsumer<T, CellData> setter, T instance, CellData cellData,
                                int columnIndex, List<String> headerNames, List<String> messages) {
        return mapColumn(setter, instance, cellData, columnIndex, headerNames, messages, null);
    }

    protected boolean mapColumn(java.util.function.BiConsumer<T, CellData> setter, T instance, CellData cellData,
                                int columnIndex, List<String> headerNames, List<String> messages,
                                @Nullable List<CellError> cellErrors) {
        try {
            setter.accept(instance, cellData);
            return true;
        } catch (Exception e) {
            String header = (columnIndex < headerNames.size()) ? headerNames.get(columnIndex) : "column#" + columnIndex;
            String message = "Failed to set column '" + header + "': value='" + cellData.formattedValue() + "', reason=" + e.getMessage();
            messages.add(message);
            if (cellErrors != null) {
                cellErrors.add(new CellError(columnIndex, header, cellData.formattedValue(), message));
            }
            log.warn("Column mapping failed for '{}': value='{}'", header, cellData.formattedValue(), e);
            return false;
        }
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
        if (column.isRequired() && cellData.isEmpty()) {
            String header = (columnIndex < headerNames.size()) ? headerNames.get(columnIndex) : "column#" + columnIndex;
            String message = "Required column '" + header + "' is empty";
            messages.add(message);
            if (cellErrors != null) {
                cellErrors.add(new CellError(columnIndex, header, cellData.formattedValue(), message));
            }
            return false;
        }
        return mapColumn(column.setter(), instance, cellData, columnIndex, headerNames, messages, cellErrors);
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
        T instance;
        try {
            instance = rowMapper.apply(rowData);
        } catch (Exception e) {
            log.warn("Row mapping failed", e);
            List<String> messages = new ArrayList<>();
            messages.add("Row mapping failed: " + e.getMessage());
            return new ReadResult<>(null, false, messages, e, fileRowNum, List.of(), rawValues);
        }
        List<String> messages = new ArrayList<>();
        boolean valid = validateIfNeeded(instance, messages);
        return new ReadResult<>(instance, valid, messages.isEmpty() ? null : messages, null, fileRowNum, List.of(), rawValues);
    }
}
