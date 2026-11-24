package io.github.dornol.excelkit.csv;

import com.opencsv.CSVParser;
import com.opencsv.CSVParserBuilder;
import io.github.dornol.excelkit.core.read.AbstractReadHandler;
import io.github.dornol.excelkit.core.read.CellData;
import io.github.dornol.excelkit.core.Cursor;
import io.github.dornol.excelkit.core.ExcelKitException;
import io.github.dornol.excelkit.core.read.ReadColumn;
import io.github.dornol.excelkit.core.read.ReadAbortException;
import io.github.dornol.excelkit.core.read.ReadResult;
import io.github.dornol.excelkit.core.read.ReadLimitExceededException;
import io.github.dornol.excelkit.core.read.ReadStoppedException;
import io.github.dornol.excelkit.core.read.CellError;
import io.github.dornol.excelkit.core.read.RowData;
import io.github.dornol.excelkit.core.read.RowError;
import org.jspecify.annotations.Nullable;

import java.io.InputStream;
import java.io.InputStreamReader;
import java.io.IOException;
import java.nio.charset.CharacterCodingException;
import java.nio.charset.Charset;
import java.nio.charset.CodingErrorAction;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.Collections;
import java.util.List;
import java.util.Map;
import java.util.function.Consumer;
import java.util.Arrays;

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
 * @since 0.1.3
 */
final class CsvReadHandler<T> extends AbstractReadHandler<T> {
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
        validateOrClose(() -> {
            validateHeaderRowIndex(config.headerRowIndex());
            if (config.columns() != null) validateColumns(config.columns());
        });
        this.columns = config.columns();
        this.headerRowIndex = config.headerRowIndex();
        this.delimiter = config.delimiter();
        this.charset = config.charset();
        this.quoteChar = config.quoteChar();
        this.escapeChar = config.escapeChar();
        this.strictQuotes = config.strictQuotes();
        this.ignoreLeadingWhiteSpace = config.ignoreLeadingWhiteSpace();
        this.countRows = config.countRows();
        if (externalPath) useExternalInput(path);
        validateOrClose(this::validateReadOptions);
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
    private final boolean countRows;

    @Override
    public void read(Consumer<ReadResult<T>> consumer) {
        markConsumed();
        Consumer<ReadResult<T>> guardedConsumer = guardedConsumer(consumer);
        long totalRows = -1;
        try {
            if (countRows) totalRows = preScanRowCount();
            Cursor cursor = null;
            if (readOptions.progressCallback() != null) {
                cursor = new Cursor();
                if (totalRows >= 0) cursor.setTotalRows(totalRows);
            }
            readRecords(guardedConsumer, totalRows, cursor);
        } finally {
            notifyReadCompletion(-1, totalRows);
            close();
        }
    }

    private void readRecords(Consumer<ReadResult<T>> guardedConsumer, long totalRows, @Nullable Cursor cursor) {
        try (CsvRecordReader reader = buildCsvReader()) {
            skipToHeader(reader);
            String[] headerLine = readHeaderLine(reader);
            prepareColumnHeaders(headerLine);
            prepareUniqueKey(headerNames);

            String[] line;
            long rowCount = 0;

            if (rowMapper != null) {
                Map<String, Integer> headerIndexMap = buildHeaderIndexMap(headerNames, "CSV");
                validateSelectedMapColumns(headerIndexMap, headerNames, "CSV");
                while ((line = reader.readNext()) != null) {
                    List<String> rawValues = rawValues(line);
                    int rowDecision = evaluateRow(rawValues);
                    if (rowDecision == STOP_READING) break;
                    // A skipped row is not delivered, so it is not counted: progress and the final summary agree.
                    if (rowDecision == SKIP_ROW) continue;
                    guardedConsumer.accept(applyUniqueKey(
                            processRowMapping(line, headerIndexMap, reader.recordStartLine(), rawValues)));
                    recordEmittedRow();
                    rowCount++;
                    notifyLegacyProgress(rowCount, cursor);
                    if (readOptions.readProgressInterval() > 0
                            && rowCount % readOptions.readProgressInterval() == 0) {
                        notifyReadProgress(rowCount, -1, totalRows);
                    }
                }
            } else {
                int[] resolvedIndices = resolveIndices();
                while ((line = reader.readNext()) != null) {
                    List<String> rawValues = rawValues(line);
                    int rowDecision = evaluateRow(rawValues);
                    if (rowDecision == STOP_READING) break;
                    // A skipped row is not delivered, so it is not counted: progress and the final summary agree.
                    if (rowDecision == SKIP_ROW) continue;
                    guardedConsumer.accept(applyUniqueKey(
                            processRow(line, resolvedIndices, reader.recordStartLine(), rawValues)));
                    recordEmittedRow();
                    rowCount++;
                    notifyLegacyProgress(rowCount, cursor);
                    if (readOptions.readProgressInterval() > 0
                            && rowCount % readOptions.readProgressInterval() == 0) {
                        notifyReadProgress(rowCount, -1, totalRows);
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
        } catch (IOException | RuntimeException e) {
            throw new CsvReadException(failureMessage(e), e);
        }
    }

    /**
     * The records after the header row, parsed as the read will parse them. A failure is not this pass's to report:
     * the read that follows meets the same input and reports it with its own message, so here the total is unknown.
     */
    private long preScanRowCount() {
        try (CsvRecordReader scan = buildCsvReader()) {
            for (int i = 0; i <= headerRowIndex; i++) {
                if (scan.readNext() == null) return -1;
            }
            long rows = 0;
            while (scan.readNext() != null) rows++;
            return rows;
        } catch (IOException | RuntimeException e) {
            return -1;
        }
    }

    /**
     * With {@code strictQuotes} the parser drops text outside quotes, so a header row that is not
     * quoted turns into blank headers and the columns are then "not found". Say so, because the
     * cause alone ("Available headers: [, ]") does not explain it.
     */
    private String failureMessage(Exception cause) {
        if (strictQuotes && cause instanceof ExcelKitException
                && headerNames.stream().anyMatch(String::isBlank)) {
            return "Failed to read CSV: a header cell is blank. With strictQuotes(true) text outside quotes is "
                    + "dropped, so every header cell must be quoted; quote the header row or turn strictQuotes "
                    + "off. Cause: " + cause.getMessage();
        }
        if (cause instanceof CsvFormatException) {
            return "Failed to read CSV: " + cause.getMessage();
        }
        if (cause instanceof CharacterCodingException) {
            return "Failed to read CSV: the input is not valid " + charset.name() + " text. If the file uses another "
                    + "encoding (EUC-KR, windows-1252, UTF-16, ...), set it with charset(...); " + cause.getMessage();
        }
        return readFailureMessage("Failed to read CSV", cause);
    }


    private ReadResult<T> processRow(String[] line, int[] resolvedIndices, long fileRowNum, List<String> rawValues) {
        validateColumnCount(line.length);
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
        validateColumnCount(line.length);
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

    private void skipToHeader(CsvRecordReader reader) throws IOException {
        for (int i = 0; i < headerRowIndex; i++) {
            if (reader.readNext() == null) {
                throw new CsvReadException("CSV file has insufficient rows for headerRowIndex=" + headerRowIndex);
            }
        }
    }

    private String[] readHeaderLine(CsvRecordReader reader) throws IOException {
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

    private CsvRecordReader buildCsvReader() throws IOException {
        CSVParser csvParser = new CSVParserBuilder()
                .withSeparator(this.delimiter)
                .withQuoteChar(this.quoteChar)
                .withEscapeChar(this.escapeChar)
                .withStrictQuotes(this.strictQuotes)
                .withIgnoreLeadingWhiteSpace(this.ignoreLeadingWhiteSpace)
                .build();
        var decoder = charset.newDecoder()
                .onMalformedInput(CodingErrorAction.REPORT)
                .onUnmappableCharacter(CodingErrorAction.REPORT);
        return new CsvRecordReader(new InputStreamReader(Files.newInputStream(getTempFile()), decoder), csvParser);
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
