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
import java.util.ArrayList;
import java.util.Collections;
import java.util.concurrent.atomic.AtomicLong;
import java.util.List;
import java.util.Map;
import java.util.Set;
import java.util.Spliterator;
import java.util.Spliterators;
import java.util.function.Consumer;
import java.util.function.Function;
import java.util.function.Supplier;
import java.util.stream.Stream;
import java.util.stream.StreamSupport;

/**
 * Reads CSV files and maps rows to Java objects.
 *
 * <h2>Resource management</h2>
 * On construction, the handler copies the input stream to a temporary file on disk.
 * These temp resources are released when:
 * <ul>
 *     <li>{@link #read(java.util.function.Consumer)} / {@link #readStrict(java.util.function.Consumer)}
 *         returns or throws — cleanup is automatic.</li>
 *     <li>The stream returned by {@link #readAsStream()} is closed — <strong>always use
 *         try-with-resources</strong>, since this stream also holds a background producer thread:
 *         <pre>{@code
 * try (Stream<ReadResult<T>> stream = handler.readAsStream()) {
 *     stream.forEach(result -> ...);
 * }
 *         }</pre>
 *         Abandoning the stream without closing it leaks the temp file until the JVM exits.</li>
 * </ul>
 *
 * @param <T> The target row data type
 * @author dhkim
 * @since 2025-07-19
 */
public class CsvReadHandler<T> extends AbstractReadHandler<T> {
    private static final Logger log = LoggerFactory.getLogger(CsvReadHandler.class);

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
                    consumer.accept(processRowMapping(line, headerIndexMap, fileRowNum(rowCount), rawValues));
                    emittedRows++;
                    rowCount++;
                    if (progressCallback != null && progressInterval > 0 && rowCount % progressInterval == 0) {
                        progressCallback.onProgress(rowCount, null);
                    }
                }
            } else {
                int[] resolvedIndices = resolveIndices();
                while ((line = reader.readNext()) != null) {
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
                    consumer.accept(processRow(line, resolvedIndices, fileRowNum(rowCount), rawValues));
                    emittedRows++;
                    rowCount++;
                    if (progressCallback != null && progressInterval > 0 && rowCount % progressInterval == 0) {
                        progressCallback.onProgress(rowCount, null);
                    }
                }
            }
        } catch (CsvReadException | ReadAbortException e) {
            throw e;
        } catch (Exception e) {
            throw new CsvReadException("Failed to read CSV", e);
        } finally {
            close();
        }
    }

    /**
     * Reads the CSV file as a stream of row results.
     * <p>
     * <strong>Important:</strong> The returned stream holds file resources (CSVReader, temp file).
     * Always use try-with-resources to ensure proper cleanup:
     * <pre>{@code
     * try (Stream<ReadResult<T>> stream = handler.readAsStream()) {
     *     stream.forEach(result -> ...);
     * }
     * }</pre>
     *
     * @return A stream of parsed and validated row results
     */
    @Override
    public Stream<ReadResult<T>> readAsStream() {
        markConsumed();
        CSVReader reader = null;
        try {
            reader = buildCsvReader();
            skipToHeader(reader);
            String[] headerLine = readHeaderLine(reader);
            if (headerLine == null) {
                closeQuietly(reader);
                throw new CsvReadException("CSV file is empty or missing header row");
            }
            prepareColumnHeaders(headerLine);

            final boolean mappingMode = rowMapper != null;
            final int[] resolvedIndices = mappingMode ? null : resolveIndices();
            final Map<String, Integer> headerIndexMap = mappingMode ? buildHeaderIndexMap(headerNames, "CSV") : null;
            if (mappingMode) {
                validateSelectedMapColumns(headerIndexMap, headerNames, "CSV");
            }

            final CSVReader csvReader = reader;
            final AtomicLong streamRowCount = new AtomicLong(0);
            final AtomicLong emittedRows = new AtomicLong(0);
            final java.util.concurrent.atomic.AtomicInteger consecutiveBlankRows = new java.util.concurrent.atomic.AtomicInteger(0);
            Spliterator<ReadResult<T>> spliterator = new Spliterators.AbstractSpliterator<>(
                    Long.MAX_VALUE, Spliterator.ORDERED | Spliterator.NONNULL) {
                @Override
                public boolean tryAdvance(Consumer<? super ReadResult<T>> action) {
                    try {
                        String[] line;
                        List<String> rawValues;
                        long countBeforeIncrement;
                        while (true) {
                            line = csvReader.readNext();
                            if (line == null) {
                                closeQuietly(csvReader);
                                close();
                                return false;
                            }
                            countBeforeIncrement = streamRowCount.get();
                            rawValues = rawValues(line);
                            if (isBlankValues(rawValues)) {
                                int blankCount = consecutiveBlankRows.incrementAndGet();
                                if (stopAtBlankRows > 0 && blankCount >= stopAtBlankRows) {
                                    closeQuietly(csvReader);
                                    close();
                                    return false;
                                }
                                if (skipBlankRows) {
                                    streamRowCount.incrementAndGet();
                                    continue;
                                }
                            } else {
                                consecutiveBlankRows.set(0);
                            }
                            if (maxRows >= 0 && emittedRows.get() >= maxRows) {
                                closeQuietly(csvReader);
                                close();
                                return false;
                            }
                            break;
                        }
                        ReadResult<T> result = mappingMode
                                ? processRowMapping(line, headerIndexMap, fileRowNum(countBeforeIncrement), rawValues)
                                : processRow(line, resolvedIndices, fileRowNum(countBeforeIncrement), rawValues);
                        action.accept(result);
                        emittedRows.incrementAndGet();
                        long count = streamRowCount.incrementAndGet();
                        if (progressCallback != null && progressInterval > 0
                                && count % progressInterval == 0) {
                            progressCallback.onProgress(count, null);
                        }
                        return true;
                    } catch (Exception e) {
                        closeQuietly(csvReader);
                        close();
                        throw new CsvReadException("Failed to read CSV row", e);
                    }
                }
            };

            reader = null; // ownership transferred to spliterator/onClose
            return StreamSupport.stream(spliterator, false)
                    .onClose(() -> {
                        closeQuietly(csvReader);
                        close();
                    });
        } catch (CsvReadException e) {
            closeQuietly(reader);
            close();
            throw e;
        } catch (Exception e) {
            closeQuietly(reader);
            close();
            throw new CsvReadException("Failed to initialize CSV reading", e);
        }
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
        RowData rowData = new RowData(cells, headerNames, headerIndexMap);
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
