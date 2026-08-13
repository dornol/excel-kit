package io.github.dornol.excelkit.csv;

import com.opencsv.CSVParser;
import com.opencsv.CSVParserBuilder;
import com.opencsv.CSVReader;
import com.opencsv.CSVReaderBuilder;
import com.opencsv.ICSVParser;
import com.opencsv.exceptions.CsvValidationException;
import io.github.dornol.excelkit.core.AbstractReadHandler;
import io.github.dornol.excelkit.core.CellData;
import io.github.dornol.excelkit.core.CellConversionConfig;
import io.github.dornol.excelkit.core.DuplicateHeaderPolicy;
import io.github.dornol.excelkit.core.ReadColumn;
import io.github.dornol.excelkit.core.ReadAbortException;
import io.github.dornol.excelkit.core.ReadResult;
import io.github.dornol.excelkit.core.ReadRowPolicy;
import io.github.dornol.excelkit.core.ProgressCallback;
import io.github.dornol.excelkit.core.ReadLimitExceededException;
import io.github.dornol.excelkit.core.ReadStoppedException;
import io.github.dornol.excelkit.core.CellError;
import io.github.dornol.excelkit.core.RowData;
import io.github.dornol.excelkit.core.RowError;
import jakarta.validation.Validator;
import org.jspecify.annotations.Nullable;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.InputStream;
import java.io.InputStreamReader;
import java.io.IOException;
import java.nio.charset.Charset;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.Collections;
import java.util.List;
import java.util.Objects;
import java.util.Map;
import java.util.Set;
import java.util.function.Consumer;
import java.util.function.Function;
import java.util.function.Predicate;
import java.util.function.Supplier;
import java.util.Arrays;
import java.util.concurrent.atomic.AtomicReference;

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

    CsvReadHandler(InputStream input, CsvReadSessionConfig<T> config) {
        this(input, config, false, null);
    }

    CsvReadHandler(Path path, CsvReadSessionConfig<T> config) {
        this(InputStream.nullInputStream(), config, true, path);
    }

    private CsvReadHandler(InputStream input, CsvReadSessionConfig<T> config,
                           boolean externalPath, @Nullable Path path) {
        super(input, config.supplier(), config.mapper(), config.validator(), ".csv",
                config.selectedColumns(), config.options());
        validateHeaderRowIndex(config.headerRowIndex());
        if (config.columns() != null) validateColumns(config.columns());
        this.columns = config.columns();
        this.headerRowIndex = config.headerRowIndex();
        this.delimiter = config.delimiter();
        this.charset = config.charset();
        this.quoteChar = config.quoteChar();
        this.escapeChar = config.escapeChar();
        this.strictQuotes = config.strictQuotes();
        this.ignoreLeadingWhiteSpace = config.ignoreLeadingWhiteSpace();
        this.progressInterval = config.progressInterval();
        this.progressCallback = config.progressCallback();
        if (externalPath) useExternalInput(path);
        validateReadOptions();
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
    private final @Nullable ProgressCallback progressCallback;

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
            ReadRowPolicy rowPolicy = new ReadRowPolicy(readOptions);

            if (rowMapper != null) {
                Map<String, Integer> headerIndexMap = buildHeaderIndexMap(headerNames, "CSV");
                validateSelectedMapColumns(headerIndexMap, headerNames, "CSV");
                while ((line = reader.readNext()) != null) {
                    List<String> rawValues = rawValues(line);
                    ReadRowPolicy.Decision decision = rowPolicy.evaluate(rawValues);
                    if (decision == ReadRowPolicy.Decision.STOP) break;
                    if (decision == ReadRowPolicy.Decision.SKIP) {
                        rowCount++;
                        continue;
                    }
                    guardedConsumer.accept(processRowMapping(line, headerIndexMap, fileRowNum(rowCount), rawValues));
                    rowPolicy.recordEmitted();
                    rowCount++;
                    if (progressCallback != null && progressInterval > 0 && rowCount % progressInterval == 0) {
                        progressCallback.onProgress(rowCount, null);
                    }
                    if (readOptions.readProgressInterval() > 0
                            && rowCount % readOptions.readProgressInterval() == 0) {
                        notifyReadProgress(rowCount, -1, -1);
                    }
                }
            } else {
                int[] resolvedIndices = resolveIndices();
                while ((line = reader.readNext()) != null) {
                    List<String> rawValues = rawValues(line);
                    ReadRowPolicy.Decision decision = rowPolicy.evaluate(rawValues);
                    if (decision == ReadRowPolicy.Decision.STOP) break;
                    if (decision == ReadRowPolicy.Decision.SKIP) {
                        rowCount++;
                        continue;
                    }
                    guardedConsumer.accept(processRow(line, resolvedIndices, fileRowNum(rowCount), rawValues));
                    rowPolicy.recordEmitted();
                    rowCount++;
                    if (progressCallback != null && progressInterval > 0 && rowCount % progressInterval == 0) {
                        progressCallback.onProgress(rowCount, null);
                    }
                    if (readOptions.readProgressInterval() > 0
                            && rowCount % readOptions.readProgressInterval() == 0) {
                        notifyReadProgress(rowCount, -1, -1);
                    }
                }
            }
        } catch (ReadStoppedException e) {
            // Normal early completion requested by readWhile.
            stoppedEarly = true;
        } catch (ReadLimitExceededException e) {
            throw e;
        } catch (CsvReadException | ReadAbortException e) {
            throw e;
        } catch (IOException | CsvValidationException | RuntimeException e) {
            throw new CsvReadException("Failed to read CSV", e);
        } finally {
            notifyReadCompletion(-1, -1);
            close();
        }
    }

    boolean wasStoppedEarly() { return stoppedEarly; }

    @Override
    public void readWhile(Predicate<ReadResult<T>> predicate) {
        Objects.requireNonNull(predicate, "predicate cannot be null");
        AtomicReference<RuntimeException> failure = new AtomicReference<>();
        read(result -> {
            boolean proceed;
            try {
                proceed = predicate.test(result);
            } catch (RuntimeException e) {
                failure.set(e);
                throw new ReadStoppedException();
            }
            if (!proceed) {
                throw new ReadStoppedException();
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
        boolean mappingFailure = false;
        List<String> messages = new ArrayList<>();
        List<CellError> cellErrors = new ArrayList<>();

        for (int i = 0; i < columns.size(); i++) {
            int actualIndex = resolvedIndices[i];
            String columnValue = (actualIndex < line.length) ? line[actualIndex] : null;
            ReadColumn<T> column = columns.get(i);
            CellData cell = cellData(actualIndex, columnValue);
            if (!mapColumn(column, currentInstance, cell,
                    actualIndex, headerNames, messages, cellErrors)) {
                success = false;
                if (!(column.isRequired() && cell.isEmpty())) mappingFailure = true;
            }
        }

        boolean validationSuccess = success && validateIfNeeded(currentInstance, messages);
        if (validationSuccess) return ReadResult.success(currentInstance, fileRowNum, cellErrors, rawValues);
        return ReadResult.failure(currentInstance,
                mappingFailure ? RowError.Type.MAPPING : RowError.Type.VALIDATION,
                messages, null, fileRowNum, cellErrors, rawValues);
    }

    private ReadResult<T> processRowMapping(String[] line, Map<String, Integer> headerIndexMap, long fileRowNum,
                                            List<String> rawValues) {
        List<CellData> cells = new ArrayList<>();
        for (int i = 0; i < line.length; i++) {
            cells.add(cellData(i, line[i]));
        }
        RowData rowData = new RowData(cells, headerNames, headerIndexMap, readOptions.headerNormalizer());
        return mapWithRowMapper(rowData, fileRowNum, rawValues);
    }

    private List<String> rawValues(String[] line) {
        return Arrays.asList(line.clone());
    }

    private void skipToHeader(CSVReader reader) throws IOException, CsvValidationException {
        for (int i = 0; i < headerRowIndex; i++) {
            if (reader.readNext() == null) {
                throw new CsvReadException("CSV file has insufficient rows for headerRowIndex=" + headerRowIndex);
            }
        }
    }

    private String[] readHeaderLine(CSVReader reader) throws IOException, CsvValidationException {
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
        } catch (IOException e) {
            // Log at debug level — closeQuietly is used in cleanup paths
            // where reporting errors is less critical
            if (log.isDebugEnabled()) {
                log.debug("Failed to close CSVReader", e);
            }
        }
    }

    private CSVReader buildCsvReader() throws IOException {
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
