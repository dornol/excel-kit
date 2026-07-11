package io.github.dornol.excelkit.csv;

import com.opencsv.CSVParser;
import com.opencsv.CSVParserBuilder;
import com.opencsv.CSVReader;
import com.opencsv.CSVReaderBuilder;
import com.opencsv.ICSVParser;
import io.github.dornol.excelkit.core.AbstractReadHandler;
import io.github.dornol.excelkit.core.CellData;
import io.github.dornol.excelkit.core.CellConversionConfig;
import io.github.dornol.excelkit.core.DuplicateHeaderPolicy;
import io.github.dornol.excelkit.core.ReadColumn;
import io.github.dornol.excelkit.core.ReadAbortException;
import io.github.dornol.excelkit.core.ReadResult;
import io.github.dornol.excelkit.core.RowData;
import jakarta.validation.Validator;
import org.jspecify.annotations.Nullable;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.InputStream;
import java.io.InputStreamReader;
import java.nio.charset.Charset;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.Collections;
import java.util.List;
import java.util.Map;
import java.util.Set;
import java.util.function.Consumer;
import java.util.function.Function;
import java.util.function.Predicate;
import java.util.function.Supplier;

/**
 * Reads CSV files and maps rows to Java objects.
 *
 * <h2>Resource management</h2>
 * On construction, the handler copies the input stream to a temporary file on disk.
 * Temporary resources are released when {@link #read(java.util.function.Consumer)},
 * {@link #readWhile(java.util.function.Predicate)}, or {@link #readStrict(java.util.function.Consumer)}
 * returns or throws.
 *
 * @param <T> The target row data type
 * @author dhkim
 * @since 2025-07-19
 */
final class CsvReadHandler<T> extends AbstractReadHandler<T> {
    private static final Logger log = LoggerFactory.getLogger(CsvReadHandler.class);

    static <T> CsvReadHandler<T> forPath(Path path, List<ReadColumn<T>> columns, Supplier<T> supplier,
            @Nullable Validator validator, int headerRowIndex, char delimiter, Charset charset,
            int progressInterval, io.github.dornol.excelkit.core.@Nullable ProgressCallback progressCallback,
            boolean strictHeaders, DuplicateHeaderPolicy duplicateHeaderPolicy,
            @Nullable CellConversionConfig conversion, char quoteChar, char escapeChar,
            boolean strictQuotes, boolean ignoreLeadingWhiteSpace, long maxRows,
            boolean skipBlankRows, int stopAtBlankRows) {
        var handler = new CsvReadHandler<>(InputStream.nullInputStream(), columns, supplier, validator,
                headerRowIndex, delimiter, charset, progressInterval, progressCallback, strictHeaders,
                duplicateHeaderPolicy, conversion, quoteChar, escapeChar, strictQuotes,
                ignoreLeadingWhiteSpace, maxRows, skipBlankRows, stopAtBlankRows);
        handler.useExternalInput(path);
        return handler;
    }

    static <T> CsvReadHandler<T> forPath(Path path, Function<RowData, T> mapper,
            @Nullable Validator validator, int headerRowIndex, char delimiter, Charset charset,
            int progressInterval, io.github.dornol.excelkit.core.@Nullable ProgressCallback progressCallback,
            boolean strictHeaders, DuplicateHeaderPolicy duplicateHeaderPolicy,
            @Nullable Set<String> selectedColumns, @Nullable CellConversionConfig conversion,
            char quoteChar, char escapeChar, boolean strictQuotes, boolean ignoreLeadingWhiteSpace,
            long maxRows, boolean skipBlankRows, int stopAtBlankRows) {
        var handler = new CsvReadHandler<>(InputStream.nullInputStream(), mapper, validator,
                headerRowIndex, delimiter, charset, progressInterval, progressCallback, strictHeaders,
                duplicateHeaderPolicy, selectedColumns, conversion, quoteChar, escapeChar, strictQuotes,
                ignoreLeadingWhiteSpace, maxRows, skipBlankRows, stopAtBlankRows);
        handler.useExternalInput(path);
        return handler;
    }

    private final List<String> headerNames = new ArrayList<>();
    private final @Nullable List<ReadColumn<T>> columns;
    private final int headerRowIndex;
    private final char delimiter;
    private final Charset charset;
    private final char quoteChar;
    private final char escapeChar;
    private final boolean strictQuotes;
    private final boolean ignoreLeadingWhiteSpace;
    private final int progressInterval;
    private final io.github.dornol.excelkit.core.@Nullable ProgressCallback progressCallback;

    CsvReadHandler(InputStream inputStream, List<ReadColumn<T>> columns, Supplier<T> instanceSupplier, @Nullable Validator validator) {
        this(inputStream, columns, instanceSupplier, validator, 0, ',', StandardCharsets.UTF_8, 0, null);
    }

    CsvReadHandler(InputStream inputStream, List<ReadColumn<T>> columns, Supplier<T> instanceSupplier, @Nullable Validator validator, int headerRowIndex) {
        this(inputStream, columns, instanceSupplier, validator, headerRowIndex, ',', StandardCharsets.UTF_8, 0, null);
    }

    CsvReadHandler(InputStream inputStream, List<ReadColumn<T>> columns, Supplier<T> instanceSupplier,
                   @Nullable Validator validator, int headerRowIndex, char delimiter, Charset charset) {
        this(inputStream, columns, instanceSupplier, validator, headerRowIndex, delimiter, charset, 0, null);
    }

    CsvReadHandler(InputStream inputStream, List<ReadColumn<T>> columns, Supplier<T> instanceSupplier,
                   @Nullable Validator validator, int headerRowIndex, char delimiter, Charset charset,
                   int progressInterval, io.github.dornol.excelkit.core.@Nullable ProgressCallback progressCallback) {
        this(inputStream, columns, instanceSupplier, validator, headerRowIndex, delimiter, charset,
                progressInterval, progressCallback, false, DuplicateHeaderPolicy.FIRST);
    }

    CsvReadHandler(InputStream inputStream, List<ReadColumn<T>> columns, Supplier<T> instanceSupplier,
                   @Nullable Validator validator, int headerRowIndex, char delimiter, Charset charset,
                   int progressInterval, io.github.dornol.excelkit.core.@Nullable ProgressCallback progressCallback,
                   boolean strictHeaders, DuplicateHeaderPolicy duplicateHeaderPolicy) {
        this(inputStream, columns, instanceSupplier, validator, headerRowIndex, delimiter, charset,
                progressInterval, progressCallback, strictHeaders, duplicateHeaderPolicy, null);
    }

    CsvReadHandler(InputStream inputStream, List<ReadColumn<T>> columns, Supplier<T> instanceSupplier,
                   @Nullable Validator validator, int headerRowIndex, char delimiter, Charset charset,
                   int progressInterval, io.github.dornol.excelkit.core.@Nullable ProgressCallback progressCallback,
                   boolean strictHeaders, DuplicateHeaderPolicy duplicateHeaderPolicy,
                   @Nullable CellConversionConfig cellConversionConfig) {
        this(inputStream, columns, instanceSupplier, validator, headerRowIndex, delimiter, charset,
                progressInterval, progressCallback, strictHeaders, duplicateHeaderPolicy, cellConversionConfig,
                ICSVParser.DEFAULT_QUOTE_CHARACTER, ICSVParser.DEFAULT_ESCAPE_CHARACTER,
                ICSVParser.DEFAULT_STRICT_QUOTES, ICSVParser.DEFAULT_IGNORE_LEADING_WHITESPACE);
    }

    CsvReadHandler(InputStream inputStream, List<ReadColumn<T>> columns, Supplier<T> instanceSupplier,
                   @Nullable Validator validator, int headerRowIndex, char delimiter, Charset charset,
                   int progressInterval, io.github.dornol.excelkit.core.@Nullable ProgressCallback progressCallback,
                   boolean strictHeaders, DuplicateHeaderPolicy duplicateHeaderPolicy,
                   @Nullable CellConversionConfig cellConversionConfig,
                   char quoteChar, char escapeChar, boolean strictQuotes, boolean ignoreLeadingWhiteSpace) {
        this(inputStream, columns, instanceSupplier, validator, headerRowIndex, delimiter, charset,
                progressInterval, progressCallback, strictHeaders, duplicateHeaderPolicy, cellConversionConfig,
                quoteChar, escapeChar, strictQuotes, ignoreLeadingWhiteSpace, -1, false, 0);
    }

    CsvReadHandler(InputStream inputStream, List<ReadColumn<T>> columns, Supplier<T> instanceSupplier,
                   @Nullable Validator validator, int headerRowIndex, char delimiter, Charset charset,
                   int progressInterval, io.github.dornol.excelkit.core.@Nullable ProgressCallback progressCallback,
                   boolean strictHeaders, DuplicateHeaderPolicy duplicateHeaderPolicy,
                   @Nullable CellConversionConfig cellConversionConfig,
                   char quoteChar, char escapeChar, boolean strictQuotes, boolean ignoreLeadingWhiteSpace,
                   long maxRows, boolean skipBlankRows, int stopAtBlankRows) {
        super(inputStream, instanceSupplier, validator, ".csv", strictHeaders, duplicateHeaderPolicy,
                null, cellConversionConfig, maxRows, skipBlankRows, stopAtBlankRows);
        validateColumns(columns);
        validateHeaderRowIndex(headerRowIndex);
        this.columns = columns;
        this.headerRowIndex = headerRowIndex;
        this.delimiter = delimiter;
        this.charset = charset;
        this.quoteChar = quoteChar;
        this.escapeChar = escapeChar;
        this.strictQuotes = strictQuotes;
        this.ignoreLeadingWhiteSpace = ignoreLeadingWhiteSpace;
        this.progressInterval = progressInterval;
        this.progressCallback = progressCallback;
    }

    /**
     * Constructs a handler in mapping mode for immutable object construction.
     */
    CsvReadHandler(InputStream inputStream, Function<RowData, T> rowMapper,
                   @Nullable Validator validator, int headerRowIndex, char delimiter, Charset charset,
                   int progressInterval, io.github.dornol.excelkit.core.@Nullable ProgressCallback progressCallback) {
        this(inputStream, rowMapper, validator, headerRowIndex, delimiter, charset, progressInterval,
                progressCallback, false, DuplicateHeaderPolicy.FIRST);
    }

    CsvReadHandler(InputStream inputStream, Function<RowData, T> rowMapper,
                   @Nullable Validator validator, int headerRowIndex, char delimiter, Charset charset,
                   int progressInterval, io.github.dornol.excelkit.core.@Nullable ProgressCallback progressCallback,
                   boolean strictHeaders, DuplicateHeaderPolicy duplicateHeaderPolicy) {
        this(inputStream, rowMapper, validator, headerRowIndex, delimiter, charset, progressInterval,
                progressCallback, strictHeaders, duplicateHeaderPolicy, null);
    }

    CsvReadHandler(InputStream inputStream, Function<RowData, T> rowMapper,
                   @Nullable Validator validator, int headerRowIndex, char delimiter, Charset charset,
                   int progressInterval, io.github.dornol.excelkit.core.@Nullable ProgressCallback progressCallback,
                   boolean strictHeaders, DuplicateHeaderPolicy duplicateHeaderPolicy,
                   @Nullable Set<String> selectedMapColumns) {
        this(inputStream, rowMapper, validator, headerRowIndex, delimiter, charset, progressInterval,
                progressCallback, strictHeaders, duplicateHeaderPolicy, selectedMapColumns, null);
    }

    CsvReadHandler(InputStream inputStream, Function<RowData, T> rowMapper,
                   @Nullable Validator validator, int headerRowIndex, char delimiter, Charset charset,
                   int progressInterval, io.github.dornol.excelkit.core.@Nullable ProgressCallback progressCallback,
                   boolean strictHeaders, DuplicateHeaderPolicy duplicateHeaderPolicy,
                   @Nullable Set<String> selectedMapColumns,
                   @Nullable CellConversionConfig cellConversionConfig) {
        this(inputStream, rowMapper, validator, headerRowIndex, delimiter, charset, progressInterval,
                progressCallback, strictHeaders, duplicateHeaderPolicy, selectedMapColumns, cellConversionConfig,
                ICSVParser.DEFAULT_QUOTE_CHARACTER, ICSVParser.DEFAULT_ESCAPE_CHARACTER,
                ICSVParser.DEFAULT_STRICT_QUOTES, ICSVParser.DEFAULT_IGNORE_LEADING_WHITESPACE);
    }

    CsvReadHandler(InputStream inputStream, Function<RowData, T> rowMapper,
                   @Nullable Validator validator, int headerRowIndex, char delimiter, Charset charset,
                   int progressInterval, io.github.dornol.excelkit.core.@Nullable ProgressCallback progressCallback,
                   boolean strictHeaders, DuplicateHeaderPolicy duplicateHeaderPolicy,
                   @Nullable Set<String> selectedMapColumns,
                   @Nullable CellConversionConfig cellConversionConfig,
                   char quoteChar, char escapeChar, boolean strictQuotes, boolean ignoreLeadingWhiteSpace) {
        this(inputStream, rowMapper, validator, headerRowIndex, delimiter, charset, progressInterval,
                progressCallback, strictHeaders, duplicateHeaderPolicy, selectedMapColumns, cellConversionConfig,
                quoteChar, escapeChar, strictQuotes, ignoreLeadingWhiteSpace, -1, false, 0);
    }

    CsvReadHandler(InputStream inputStream, Function<RowData, T> rowMapper,
                   @Nullable Validator validator, int headerRowIndex, char delimiter, Charset charset,
                   int progressInterval, io.github.dornol.excelkit.core.@Nullable ProgressCallback progressCallback,
                   boolean strictHeaders, DuplicateHeaderPolicy duplicateHeaderPolicy,
                   @Nullable Set<String> selectedMapColumns,
                   @Nullable CellConversionConfig cellConversionConfig,
                   char quoteChar, char escapeChar, boolean strictQuotes, boolean ignoreLeadingWhiteSpace,
                   long maxRows, boolean skipBlankRows, int stopAtBlankRows) {
        super(inputStream, rowMapper, validator, ".csv", strictHeaders, duplicateHeaderPolicy,
                selectedMapColumns, cellConversionConfig, maxRows, skipBlankRows, stopAtBlankRows);
        validateHeaderRowIndex(headerRowIndex);
        this.columns = null;
        this.headerRowIndex = headerRowIndex;
        this.delimiter = delimiter;
        this.charset = charset;
        this.quoteChar = quoteChar;
        this.escapeChar = escapeChar;
        this.strictQuotes = strictQuotes;
        this.ignoreLeadingWhiteSpace = ignoreLeadingWhiteSpace;
        this.progressInterval = progressInterval;
        this.progressCallback = progressCallback;
    }

    @Override
    public void read(Consumer<ReadResult<T>> consumer) {
        markConsumed();
        Consumer<ReadResult<T>> guardedConsumer = guardedConsumer(consumer);
        try (CSVReader reader = buildCsvReader()) {
            skipToHeader(reader);
            String[] headerLine = readHeaderLine(reader);
            prepareColumnHeaders(headerLine);

            String[] line;
            long rowCount = 0;
            long emittedRows = 0;
            int consecutiveBlankRows = 0;

            if (rowMapper != null) {
                Map<String, Integer> headerIndexMap = buildHeaderIndexMap(headerNames, "CSV");
                validateSelectedMapColumns(headerIndexMap, headerNames, "CSV");
                while ((line = reader.readNext()) != null) {
                    if (cancellationToken.isCancellationRequested()) throw new io.github.dornol.excelkit.core.ReadStoppedException();
                    List<String> rawValues = rawValues(line);
                    if (isBlankValues(rawValues)) {
                        consecutiveBlankRows++;
                        if (stopAtBlankRows > 0 && consecutiveBlankRows >= stopAtBlankRows) {
                            break;
                        }
                        if (skipBlankRows) {
                            rowCount++;
                            continue;
                        }
                    } else {
                        consecutiveBlankRows = 0;
                    }
                    if (maxRows >= 0 && emittedRows >= maxRows) {
                        break;
                    }
                    guardedConsumer.accept(processRowMapping(line, headerIndexMap, fileRowNum(rowCount), rawValues));
                    emittedRows++;
                    rowCount++;
                    if (progressCallback != null && progressInterval > 0 && rowCount % progressInterval == 0) {
                        progressCallback.onProgress(rowCount, null);
                    }
                    if (progressInterval > 0 && rowCount % progressInterval == 0) notifyReadProgress(rowCount, -1, -1);
                }
            } else {
                int[] resolvedIndices = resolveIndices();
                while ((line = reader.readNext()) != null) {
                    if (cancellationToken.isCancellationRequested()) throw new io.github.dornol.excelkit.core.ReadStoppedException();
                    List<String> rawValues = rawValues(line);
                    if (isBlankValues(rawValues)) {
                        consecutiveBlankRows++;
                        if (stopAtBlankRows > 0 && consecutiveBlankRows >= stopAtBlankRows) {
                            break;
                        }
                        if (skipBlankRows) {
                            rowCount++;
                            continue;
                        }
                    } else {
                        consecutiveBlankRows = 0;
                    }
                    if (maxRows >= 0 && emittedRows >= maxRows) {
                        break;
                    }
                    guardedConsumer.accept(processRow(line, resolvedIndices, fileRowNum(rowCount), rawValues));
                    emittedRows++;
                    rowCount++;
                    if (progressCallback != null && progressInterval > 0 && rowCount % progressInterval == 0) {
                        progressCallback.onProgress(rowCount, null);
                    }
                    if (progressInterval > 0 && rowCount % progressInterval == 0) notifyReadProgress(rowCount, -1, -1);
                }
            }
        } catch (io.github.dornol.excelkit.core.ReadStoppedException e) {
            // Normal early completion requested by readWhile.
            stoppedEarly = true;
        } catch (io.github.dornol.excelkit.core.ReadLimitExceededException e) {
            throw e;
        } catch (CsvReadException | ReadAbortException e) {
            throw e;
        } catch (Exception e) {
            throw new CsvReadException("Failed to read CSV", e);
        } finally {
            notifyReadCompletion(-1, -1);
            close();
        }
    }

    boolean wasStoppedEarly() { return stoppedEarly; }

    @Override
    public void readWhile(Predicate<ReadResult<T>> predicate) {
        java.util.concurrent.atomic.AtomicReference<RuntimeException> failure = new java.util.concurrent.atomic.AtomicReference<>();
        read(result -> {
            boolean proceed;
            try {
                proceed = predicate.test(result);
            } catch (RuntimeException e) {
                failure.set(e);
                throw new io.github.dornol.excelkit.core.ReadStoppedException();
            }
            if (!proceed) {
                throw new io.github.dornol.excelkit.core.ReadStoppedException();
            }
        });
        if (failure.get() != null) throw failure.get();
    }

    private ReadResult<T> processRow(String[] line, int[] resolvedIndices, long fileRowNum, List<String> rawValues) {
        if (columns == null || instanceSupplier == null) {
            throw new IllegalStateException("columns and instanceSupplier must not be null in setter mode");
        }
        T currentInstance = instanceSupplier.get();
        boolean success = true;
        List<String> messages = new ArrayList<>();
        List<io.github.dornol.excelkit.core.CellError> cellErrors = new ArrayList<>();

        for (int i = 0; i < columns.size(); i++) {
            int actualIndex = resolvedIndices[i];
            String columnValue = (actualIndex < line.length) ? line[actualIndex] : null;
            if (!mapColumn(columns.get(i), currentInstance, cellData(actualIndex, columnValue),
                    actualIndex, headerNames, messages, cellErrors)) {
                success = false;
            }
        }

        boolean validationSuccess = success && validateIfNeeded(currentInstance, messages);
        return new ReadResult<>(currentInstance, validationSuccess, messages.isEmpty() ? null : messages,
                null, fileRowNum, cellErrors, rawValues);
    }

    private ReadResult<T> processRowMapping(String[] line, Map<String, Integer> headerIndexMap, long fileRowNum,
                                            List<String> rawValues) {
        List<CellData> cells = new ArrayList<>();
        for (int i = 0; i < line.length; i++) {
            cells.add(cellData(i, line[i]));
        }
        RowData rowData = new RowData(cells, headerNames, headerIndexMap, headerNormalizer);
        return mapWithRowMapper(rowData, fileRowNum, rawValues);
    }

    private List<String> rawValues(String[] line) {
        return java.util.Arrays.asList(line.clone());
    }

    private void skipToHeader(CSVReader reader) throws Exception {
        for (int i = 0; i < headerRowIndex; i++) {
            if (reader.readNext() == null) {
                throw new CsvReadException("CSV file has insufficient rows for headerRowIndex=" + headerRowIndex);
            }
        }
    }

    private String[] readHeaderLine(CSVReader reader) throws Exception {
        String[] headerLine = reader.readNext();
        if (headerLine == null) {
            throw new CsvReadException("CSV file is empty or missing header row");
        }
        return headerLine;
    }

    private int[] resolveIndices() {
        return resolveColumnIndices(
                columns.size(),
                i -> columns.get(i).headerAliases(),
                i -> columns.get(i).columnIndex(),
                headerNames, "CSV"
        );
    }

    private long fileRowNum(long zeroBasedDataRowIndex) {
        return headerRowIndex + 2L + zeroBasedDataRowIndex;
    }

    private void closeQuietly(@Nullable CSVReader reader) {
        if (reader == null) return;
        try {
            reader.close();
        } catch (Exception e) {
            // Log at debug level — closeQuietly is used in cleanup paths
            // where reporting errors is less critical
            if (log.isDebugEnabled()) {
                log.debug("Failed to close CSVReader", e);
            }
        }
    }

    private CSVReader buildCsvReader() throws Exception {
        CSVParser csvParser = new CSVParserBuilder()
                .withSeparator(this.delimiter)
                .withQuoteChar(this.quoteChar)
                .withEscapeChar(this.escapeChar)
                .withStrictQuotes(this.strictQuotes)
                .withIgnoreLeadingWhiteSpace(this.ignoreLeadingWhiteSpace)
                .build();
        return new CSVReaderBuilder(new InputStreamReader(Files.newInputStream(getTempFile()), this.charset))
                .withCSVParser(csvParser).build();
    }

    private void prepareColumnHeaders(String[] line) {
        if (line.length > 0 && line[0] != null && line[0].startsWith("\uFEFF")) {
            line[0] = line[0].substring(1);
            if (line[0].isEmpty()) {
                throw new CsvReadException("First header column is empty after BOM removal");
            }
        }
        Collections.addAll(headerNames, line);
    }
}
