package io.github.dornol.excelkit.csv;

import com.opencsv.CSVParser;
import com.opencsv.CSVParserBuilder;
import com.opencsv.CSVReader;
import com.opencsv.CSVReaderBuilder;
import io.github.dornol.excelkit.shared.AbstractReadHandler;
import io.github.dornol.excelkit.shared.CellData;
import io.github.dornol.excelkit.shared.ReadAbortException;
import io.github.dornol.excelkit.shared.ReadResult;
import io.github.dornol.excelkit.shared.RowData;
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
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
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
 * @param <T> The target row data type
 * @author dhkim
 * @since 2025-07-19
 */
public class CsvReadHandler<T> extends AbstractReadHandler<T> {
    private static final Logger log = LoggerFactory.getLogger(CsvReadHandler.class);

    private final List<String> headerNames = new ArrayList<>();
    private final @Nullable List<CsvReadColumn<T>> columns;
    private final int headerRowIndex;
    private final char delimiter;
    private final Charset charset;
    private final int progressInterval;
    private final io.github.dornol.excelkit.shared.@Nullable ProgressCallback progressCallback;

    CsvReadHandler(InputStream inputStream, List<CsvReadColumn<T>> columns, Supplier<T> instanceSupplier, @Nullable Validator validator) {
        this(inputStream, columns, instanceSupplier, validator, 0, ',', StandardCharsets.UTF_8, 0, null);
    }

    CsvReadHandler(InputStream inputStream, List<CsvReadColumn<T>> columns, Supplier<T> instanceSupplier, @Nullable Validator validator, int headerRowIndex) {
        this(inputStream, columns, instanceSupplier, validator, headerRowIndex, ',', StandardCharsets.UTF_8, 0, null);
    }

    CsvReadHandler(InputStream inputStream, List<CsvReadColumn<T>> columns, Supplier<T> instanceSupplier,
                   @Nullable Validator validator, int headerRowIndex, char delimiter, Charset charset) {
        this(inputStream, columns, instanceSupplier, validator, headerRowIndex, delimiter, charset, 0, null);
    }

    CsvReadHandler(InputStream inputStream, List<CsvReadColumn<T>> columns, Supplier<T> instanceSupplier,
                   @Nullable Validator validator, int headerRowIndex, char delimiter, Charset charset,
                   int progressInterval, io.github.dornol.excelkit.shared.@Nullable ProgressCallback progressCallback) {
        super(inputStream, instanceSupplier, validator, ".csv");
        if (columns == null || columns.isEmpty()) {
            throw new IllegalArgumentException("Columns cannot be null or empty");
        }
        if (headerRowIndex < 0) {
            throw new IllegalArgumentException("headerRowIndex must be non-negative");
        }
        this.columns = columns;
        this.headerRowIndex = headerRowIndex;
        this.delimiter = delimiter;
        this.charset = charset;
        this.progressInterval = progressInterval;
        this.progressCallback = progressCallback;
    }

    /**
     * Constructs a handler in mapping mode for immutable object construction.
     */
    CsvReadHandler(InputStream inputStream, Function<RowData, T> rowMapper,
                   @Nullable Validator validator, int headerRowIndex, char delimiter, Charset charset,
                   int progressInterval, io.github.dornol.excelkit.shared.@Nullable ProgressCallback progressCallback) {
        super(inputStream, rowMapper, validator, ".csv");
        if (headerRowIndex < 0) {
            throw new IllegalArgumentException("headerRowIndex must be non-negative");
        }
        this.columns = null;
        this.headerRowIndex = headerRowIndex;
        this.delimiter = delimiter;
        this.charset = charset;
        this.progressInterval = progressInterval;
        this.progressCallback = progressCallback;
    }

    @Override
    public void read(Consumer<ReadResult<T>> consumer) {
        try (CSVReader reader = buildCsvReader()) {
            skipToHeader(reader);
            String[] headerLine = readHeaderLine(reader);
            prepareColumnHeaders(headerLine);

            String[] line;
            long rowCount = 0;

            if (rowMapper != null) {
                Map<String, Integer> headerIndexMap = buildHeaderIndexMap();
                while ((line = reader.readNext()) != null) {
                    consumer.accept(processRowMapping(line, headerIndexMap));
                    rowCount++;
                    if (progressCallback != null && progressInterval > 0 && rowCount % progressInterval == 0) {
                        progressCallback.onProgress(rowCount, null);
                    }
                }
            } else {
                int[] resolvedIndices = resolveIndices();
                while ((line = reader.readNext()) != null) {
                    consumer.accept(processRow(line, resolvedIndices));
                    rowCount++;
                    if (progressCallback != null && progressInterval > 0 && rowCount % progressInterval == 0) {
                        progressCallback.onProgress(rowCount, null);
                    }
                }
            }
        } catch (CsvReadException e) {
            throw e;
        } catch (ReadAbortException e) {
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
        try {
            CSVReader reader = buildCsvReader();
            skipToHeader(reader);
            String[] headerLine = readHeaderLine(reader);
            if (headerLine == null) {
                closeQuietly(reader);
                throw new CsvReadException("CSV file is empty or missing header row");
            }
            prepareColumnHeaders(headerLine);

            final boolean mappingMode = rowMapper != null;
            final int[] resolvedIndices = mappingMode ? null : resolveIndices();
            final Map<String, Integer> headerIndexMap = mappingMode ? buildHeaderIndexMap() : null;

            final long[] streamRowCount = {0};
            Spliterator<ReadResult<T>> spliterator = new Spliterators.AbstractSpliterator<>(
                    Long.MAX_VALUE, Spliterator.ORDERED | Spliterator.NONNULL) {
                @Override
                public boolean tryAdvance(Consumer<? super ReadResult<T>> action) {
                    try {
                        String[] line = reader.readNext();
                        if (line == null) {
                            closeQuietly(reader);
                            close();
                            return false;
                        }
                        ReadResult<T> result = mappingMode
                                ? processRowMapping(line, headerIndexMap)
                                : processRow(line, resolvedIndices);
                        action.accept(result);
                        streamRowCount[0]++;
                        if (progressCallback != null && progressInterval > 0
                                && streamRowCount[0] % progressInterval == 0) {
                            progressCallback.onProgress(streamRowCount[0], null);
                        }
                        return true;
                    } catch (Exception e) {
                        closeQuietly(reader);
                        close();
                        throw new CsvReadException("Failed to read CSV row", e);
                    }
                }
            };

            return StreamSupport.stream(spliterator, false)
                    .onClose(() -> {
                        closeQuietly(reader);
                        close();
                    });
        } catch (CsvReadException e) {
            close();
            throw e;
        } catch (Exception e) {
            close();
            throw new CsvReadException("Failed to initialize CSV reading", e);
        }
    }

    private ReadResult<T> processRow(String[] line, int[] resolvedIndices) {
        if (columns == null || instanceSupplier == null) {
            throw new IllegalStateException("columns and instanceSupplier must not be null in setter mode");
        }
        T currentInstance = instanceSupplier.get();
        boolean success = true;
        List<String> messages = new ArrayList<>();

        for (int i = 0; i < columns.size(); i++) {
            int actualIndex = resolvedIndices[i];
            String columnValue = (actualIndex < line.length) ? line[actualIndex] : null;
            if (!mapColumn(columns.get(i).setter(), currentInstance, new CellData(actualIndex, columnValue),
                    actualIndex, headerNames, messages)) {
                success = false;
            }
        }

        boolean validationSuccess = success && validateIfNeeded(currentInstance, messages);
        return new ReadResult<>(currentInstance, validationSuccess, messages);
    }

    private ReadResult<T> processRowMapping(String[] line, Map<String, Integer> headerIndexMap) {
        List<CellData> cells = new ArrayList<>();
        for (int i = 0; i < line.length; i++) {
            cells.add(new CellData(i, line[i]));
        }
        RowData rowData = new RowData(cells, headerNames, headerIndexMap);
        return mapWithRowMapper(rowData);
    }

    private Map<String, Integer> buildHeaderIndexMap() {
        Map<String, Integer> map = new LinkedHashMap<>();
        for (int i = 0; i < headerNames.size(); i++) {
            map.putIfAbsent(headerNames.get(i), i);
        }
        return map;
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
                i -> columns.get(i).headerName(),
                i -> columns.get(i).columnIndex(),
                headerNames, "CSV"
        );
    }

    private void closeQuietly(CSVReader reader) {
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
        CSVParser csvParser = new CSVParserBuilder().withSeparator(this.delimiter).build();
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
