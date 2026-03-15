package io.github.dornol.excelkit.csv;

import com.opencsv.CSVParser;
import com.opencsv.CSVParserBuilder;
import com.opencsv.CSVReader;
import com.opencsv.CSVReaderBuilder;
import io.github.dornol.excelkit.shared.AbstractReadHandler;
import io.github.dornol.excelkit.shared.CellData;
import io.github.dornol.excelkit.shared.ReadAbortException;
import io.github.dornol.excelkit.shared.ReadResult;
import jakarta.validation.Validator;
import org.jspecify.annotations.NonNull;

import java.io.InputStream;
import java.io.InputStreamReader;
import java.nio.charset.Charset;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.util.ArrayList;
import java.util.Collections;
import java.util.List;
import java.util.Spliterator;
import java.util.Spliterators;
import java.util.function.Consumer;
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
    private final List<String> headerNames = new ArrayList<>();
    private final List<CsvReadColumn<T>> columns;
    private final int headerRowIndex;
    private final char delimiter;
    private final Charset charset;

    CsvReadHandler(InputStream inputStream, List<CsvReadColumn<T>> columns, Supplier<T> instanceSupplier, Validator validator) {
        this(inputStream, columns, instanceSupplier, validator, 0, ',', StandardCharsets.UTF_8);
    }

    CsvReadHandler(InputStream inputStream, List<CsvReadColumn<T>> columns, Supplier<T> instanceSupplier, Validator validator, int headerRowIndex) {
        this(inputStream, columns, instanceSupplier, validator, headerRowIndex, ',', StandardCharsets.UTF_8);
    }

    CsvReadHandler(InputStream inputStream, List<CsvReadColumn<T>> columns, Supplier<T> instanceSupplier,
                   Validator validator, int headerRowIndex, char delimiter, Charset charset) {
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
    }

    @Override
    public void read(@NonNull Consumer<ReadResult<T>> consumer) {
        try (CSVReader reader = buildCsvReader()) {
            skipToHeader(reader);
            String[] headerLine = readHeaderLine(reader);
            prepareColumnHeaders(headerLine);
            int[] resolvedIndices = resolveIndices();

            String[] line;
            while ((line = reader.readNext()) != null) {
                consumer.accept(processRow(line, resolvedIndices));
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
            int[] resolvedIndices = resolveIndices();

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
                        action.accept(processRow(line, resolvedIndices));
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
        } catch (Exception ignored) {
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
                throw new CsvReadException("First header column is empty (contained only BOM character)");
            }
        }
        Collections.addAll(headerNames, line);
    }
}
