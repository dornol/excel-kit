package io.github.dornol.excelkit.csv;

import io.github.dornol.excelkit.shared.TempResourceCreator;

import java.io.IOException;
import java.io.OutputStream;
import java.io.OutputStreamWriter;
import java.io.PrintWriter;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.List;
import java.util.UUID;
import java.util.function.Function;
import java.util.stream.Stream;

/**
 * CSV writer for streaming large datasets into a temporary file.
 * <p>
 * Supports building columns dynamically, writing to a file line-by-line, and handling
 * basic CSV escaping (quotes, commas, line breaks).
 *
 * @param <T> The type of the data row
 * @author dhkim
 * @since 2025-07-19
 */
public class CsvWriter<T> {
    private final List<CsvColumn<T>> columns = new ArrayList<>();

    /**
     * Adds a new column to the CSV output using a row+cursor-based function.
     *
     * @param name     The column header
     * @param function A function to compute the value for each row
     * @return This writer instance (for chaining)
     */
    public CsvWriter<T> column(String name, CsvRowFunction<T, Object> function) {
        var column = new CsvColumn<>(name, function);
        this.columns.add(column);
        return this;
    }

    /**
     * Adds a new column using a basic row-only function.
     *
     * @param name     The column header
     * @param function A function to compute the value from the row
     * @return This writer instance
     */
    public CsvWriter<T> column(String name, Function<T, Object> function) {
        return column(name, (r, c) -> function.apply(r));
    }

    /**
     * Adds a column with a constant value for all rows.
     *
     * @param name  The column header
     * @param value The constant value
     * @return This writer instance
     */
    public CsvWriter<T> constColumn(String name, Object value) {
        return column(name, (r, c) -> value);
    }

    /**
     * Writes the given stream of rows to a temporary CSV file.
     * <p>
     * The returned {@link CsvHandler} can be used to write the file to an {@link OutputStream}.
     *
     * @param stream The row data stream
     * @return A handler for streaming the resulting CSV
     */
    public CsvHandler write(Stream<T> stream) {
        Path tempDir;
        Path tempFile;
        tempDir = TempResourceCreator.createTempDirectory();
        tempFile = TempResourceCreator.createTempFile(tempDir, UUID.randomUUID().toString(), ".csv");

        try (OutputStream os = Files.newOutputStream(tempFile)) {
            writeTempFile(stream, os);
        } catch (IOException e) {
            throw new IllegalStateException(e);
        }

        return new CsvHandler(tempDir, tempFile);
    }

    /**
     * Internal method to write CSV lines into the output stream.
     *
     * @param stream       The data stream
     * @param outputStream The output stream to write to
     */
    private void writeTempFile(Stream<T> stream, OutputStream outputStream) {
        Stream<T> sequential = stream.sequential();
        try (
                sequential;
                var writer = new PrintWriter(new OutputStreamWriter(outputStream, StandardCharsets.UTF_8))
        ) {
            CsvCursor cursor = new CsvCursor();
            cursor.initRow();

            // 헤더 출력
            writer.println(columns.stream()
                    .map(CsvColumn::getName)
                    .reduce((a, b) -> a + "," + b).orElse(""));
            cursor.plusRow();

            // 데이터 출력
            sequential.forEach(row -> {
                cursor.plusTotal();
                cursor.plusRow();
                String line = columns.stream()
                        .map(col -> col.applyFunction(row, cursor))
                        .map(CsvWriter::escapeCsv)
                        .reduce((a, b) -> a + "," + b)
                        .orElse("");
                writer.println(line);
            });
        }
    }

    /**
     * Escapes CSV value by wrapping in quotes and escaping inner quotes when necessary.
     *
     * @param input The input value (nullable)
     * @return A properly escaped CSV field
     */
    private static String escapeCsv(Object input) {
        if (input == null) {
            return "";
        }
        String value = input.toString();
        if (value.contains(",") || value.contains("\"") || value.contains("\n") || value.contains("\r")) {
            return "\"" + value.replace("\"", "\"\"") + "\"";
        }
        return value;
    }

}
