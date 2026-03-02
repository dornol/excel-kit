package io.github.dornol.excelkit.csv;

import com.opencsv.CSVReader;
import io.github.dornol.excelkit.shared.AbstractReadHandler;
import io.github.dornol.excelkit.shared.CellData;
import io.github.dornol.excelkit.shared.ReadResult;
import jakarta.validation.Validator;
import org.jspecify.annotations.NonNull;

import java.io.InputStream;
import java.io.InputStreamReader;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.util.ArrayList;
import java.util.Collections;
import java.util.List;
import java.util.function.Consumer;
import java.util.function.Supplier;

/**
 * Reads CSV files and maps rows to Java objects.
 * <p>
 * Stores the input stream as a temporary file and uses OpenCSV for parsing.
 *
 * @param <T> The target row data type
 * @author dhkim
 * @since 2025-07-19
 */
public class CsvReadHandler<T> extends AbstractReadHandler<T> {
    private final List<String> headerNames = new ArrayList<>();
    private final List<CsvReadColumn<T>> columns;
    private final int headerRowIndex;

    CsvReadHandler(InputStream inputStream, List<CsvReadColumn<T>> columns, Supplier<T> instanceSupplier, Validator validator) {
        this(inputStream, columns, instanceSupplier, validator, 0);
    }

    CsvReadHandler(InputStream inputStream, List<CsvReadColumn<T>> columns, Supplier<T> instanceSupplier, Validator validator, int headerRowIndex) {
        super(inputStream, instanceSupplier, validator, ".csv");
        if (columns == null || columns.isEmpty()) {
            throw new IllegalArgumentException("Columns cannot be null or empty");
        }
        if (headerRowIndex < 0) {
            throw new IllegalArgumentException("headerRowIndex must be non-negative");
        }
        this.columns = columns;
        this.headerRowIndex = headerRowIndex;
    }

    /**
     * Reads the CSV file and invokes the given consumer for each row result.
     *
     * @param consumer Callback to receive parsed and validated row results
     */
    @Override
    public void read(@NonNull Consumer<ReadResult<T>> consumer) {
        try (CSVReader reader = new CSVReader(new InputStreamReader(Files.newInputStream(getTempFile()), StandardCharsets.UTF_8))) {
            for (int i = 0; i < headerRowIndex; i++) {
                if (reader.readNext() == null) {
                    throw new CsvReadException("CSV file has insufficient rows for headerRowIndex=" + headerRowIndex);
                }
            }
            String[] headerLine = reader.readNext();
            if (headerLine == null) {
                throw new CsvReadException("CSV file is empty or missing header row");
            }
            prepareColumnHeaders(headerLine);

            String[] line;

            while ((line = reader.readNext()) != null) {
                T currentInstance = instanceSupplier.get();
                boolean success = true;
                List<String> messages = new ArrayList<>();

                for (int i = 0; i < columns.size(); i++) {
                    String columnValue = (i < line.length) ? line[i] : null;
                    if (!mapColumn(columns.get(i).setter(), currentInstance, new CellData(i, columnValue),
                            i, headerNames, messages)) {
                        success = false;
                    }
                }

                boolean validationSuccess = success && validateIfNeeded(currentInstance, messages);
                consumer.accept(new ReadResult<>(currentInstance, validationSuccess, messages));
            }
        } catch (CsvReadException e) {
            throw e;
        } catch (Exception e) {
            throw new CsvReadException("Failed to read CSV", e);
        } finally {
            close();
        }
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
