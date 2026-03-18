package io.github.dornol.excelkit.csv;

import io.github.dornol.excelkit.shared.Cursor;
import io.github.dornol.excelkit.shared.TempResourceCreator;
import org.jspecify.annotations.Nullable;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.IOException;
import java.io.OutputStream;
import java.io.OutputStreamWriter;
import java.io.PrintWriter;
import java.nio.charset.Charset;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import io.github.dornol.excelkit.shared.ProgressCallback;

import java.util.ArrayList;
import java.util.List;
import java.util.UUID;
import java.util.function.Function;
import java.util.stream.Collectors;
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
    private static final Logger log = LoggerFactory.getLogger(CsvWriter.class);
    private final List<CsvColumn<T>> columns = new ArrayList<>();
    private char delimiter = ',';
    private Charset charset = StandardCharsets.UTF_8;
    private boolean bom = true;
    private @Nullable CsvAfterDataWriter afterDataWriter;
    private @Nullable ProgressCallback progressCallback;
    private int progressInterval;
    private boolean csvInjectionDefense = true;

    /**
     * Sets the delimiter character used to separate fields.
     * Defaults to comma ({@code ','}).
     *
     * @param delimiter The delimiter character
     * @return This writer instance (for chaining)
     */
    public CsvWriter<T> delimiter(char delimiter) {
        this.delimiter = delimiter;
        return this;
    }

    /**
     * Sets the character encoding for the output file.
     * Defaults to {@link StandardCharsets#UTF_8}.
     *
     * @param charset The charset to use
     * @return This writer instance (for chaining)
     */
    public CsvWriter<T> charset(Charset charset) {
        this.charset = charset;
        return this;
    }

    /**
     * Sets whether to write a UTF-8 BOM at the start of the file.
     * Defaults to {@code true}.
     *
     * @param bom Whether to write the BOM
     * @return This writer instance (for chaining)
     */
    public CsvWriter<T> bom(boolean bom) {
        this.bom = bom;
        return this;
    }

    /**
     * Registers a callback that writes custom content after all data rows.
     * <p>
     * The callback receives the {@link java.io.PrintWriter} used to write the CSV,
     * allowing additional lines to be appended after the data rows.
     *
     * @param afterDataWriter the callback to invoke after data rows
     * @return This writer instance (for chaining)
     */
    public CsvWriter<T> afterData(CsvAfterDataWriter afterDataWriter) {
        this.afterDataWriter = afterDataWriter;
        return this;
    }

    /**
     * Adds a new column to the CSV output using a row+cursor-based function.
     *
     * @param name     The column header
     * @param function A function to compute the value for each row
     * @return This writer instance (for chaining)
     */
    public CsvWriter<T> column(String name, CsvRowFunction<T, @Nullable Object> function) {
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
    public CsvWriter<T> column(String name, Function<T, @Nullable Object> function) {
        return column(name, (r, c) -> function.apply(r));
    }

    /**
     * Conditionally adds a column using a row+cursor-based function.
     * If condition is false, the column is not added.
     *
     * @param name      The column header
     * @param condition Whether to include this column
     * @param function  A function to compute the value for each row
     * @return This writer instance
     */
    public CsvWriter<T> columnIf(String name, boolean condition, CsvRowFunction<T, @Nullable Object> function) {
        if (condition) {
            column(name, function);
        }
        return this;
    }

    /**
     * Conditionally adds a column using a basic row-only function.
     * If condition is false, the column is not added.
     *
     * @param name      The column header
     * @param condition Whether to include this column
     * @param function  A function to compute the value from the row
     * @return This writer instance
     */
    public CsvWriter<T> columnIf(String name, boolean condition, Function<T, @Nullable Object> function) {
        if (condition) {
            column(name, function);
        }
        return this;
    }

    /**
     * Adds a column with a constant value for all rows.
     *
     * @param name  The column header
     * @param value The constant value
     * @return This writer instance
     */
    public CsvWriter<T> constColumn(String name, @Nullable Object value) {
        return column(name, (r, c) -> value);
    }

    /**
     * Registers a progress callback that fires every {@code interval} rows.
     *
     * @param interval the number of rows between each callback invocation (must be positive)
     * @param callback the callback to invoke
     * @return This writer instance (for chaining)
     */
    public CsvWriter<T> onProgress(int interval, ProgressCallback callback) {
        if (interval <= 0) {
            throw new IllegalArgumentException("progress interval must be positive");
        }
        this.progressInterval = interval;
        this.progressCallback = callback;
        return this;
    }

    /**
     * Enables or disables CSV injection defense.
     * <p>
     * When enabled (default), cell values starting with formula characters
     * ({@code =}, {@code +}, {@code -}, {@code @}, {@code \t}, {@code \r})
     * are prefixed with a single quote to prevent formula injection.
     * <p>
     * Disable only when writing trusted data where the prefix would corrupt values.
     *
     * @param enabled whether to enable injection defense (default: true)
     * @return This writer instance (for chaining)
     */
    public CsvWriter<T> csvInjectionDefense(boolean enabled) {
        this.csvInjectionDefense = enabled;
        return this;
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
        if (this.columns.isEmpty()) {
            throw new CsvWriteException("columns setting required");
        }
        validateUniqueColumnNames();
        Path tempDir = TempResourceCreator.createTempDirectory();
        Path tempFile = TempResourceCreator.createTempFile(tempDir, UUID.randomUUID().toString(), ".csv");

        try (OutputStream os = Files.newOutputStream(tempFile)) {
            writeTempFile(stream, os);
        } catch (Exception e) {
            cleanup(tempDir);
            throw new CsvWriteException("Failed to write CSV", e);
        }

        return new CsvHandler(tempDir, tempFile);
    }

    private void cleanup(Path tempDir) {
        try {
            try (var files = Files.walk(tempDir)) {
                files.sorted(java.util.Comparator.reverseOrder())
                        .forEach(path -> {
                            try {
                                Files.deleteIfExists(path);
                            } catch (IOException e) {
                                log.warn("Failed to delete temp path: {}", path, e);
                                path.toFile().deleteOnExit();
                            }
                        });
            }
        } catch (IOException e) {
            log.warn("Failed to walk temp dir for cleanup: {}", tempDir, e);
            tempDir.toFile().deleteOnExit();
        }
    }

    /**
     * Internal method to write CSV lines into the output stream.
     *
     * @param stream       The data stream
     * @param outputStream The output stream to write to
     */
    private void writeTempFile(Stream<T> stream, OutputStream outputStream) {
        Stream<T> sequential = stream.sequential();
        String joining = String.valueOf(this.delimiter);
        try (
                sequential;
                var writer = new PrintWriter(new OutputStreamWriter(outputStream, this.charset))
        ) {
            Cursor cursor = new Cursor();
            cursor.initRow();

            // UTF-8 BOM for Excel compatibility
            if (this.bom) {
                writer.write('\uFEFF');
            }

            // Write header row
            writer.println(columns.stream()
                    .map(CsvColumn::getName)
                    .map(this::escapeCsv)
                    .collect(Collectors.joining(joining)));
            cursor.plusRow();

            // Write data rows
            sequential.forEach(row -> {
                cursor.plusTotal();
                cursor.plusRow();
                String line = columns.stream()
                        .map(col -> col.applyFunction(row, cursor))
                        .map(this::escapeCsv)
                        .collect(Collectors.joining(joining));
                writer.println(line);
                if (progressCallback != null && progressInterval > 0
                        && cursor.getCurrentTotal() % progressInterval == 0) {
                    progressCallback.onProgress(cursor.getCurrentTotal(), cursor);
                }
            });

            // Write after-data content
            if (this.afterDataWriter != null) {
                this.afterDataWriter.write(writer);
            }
        }
    }

    /**
     * Escapes CSV value by wrapping in quotes and escaping inner quotes when necessary.
     * Also defends against CSV injection by prefixing dangerous leading characters with a single quote.
     *
     * @param input The input value (nullable)
     * @return A properly escaped CSV field
     */
    private String escapeCsv(@Nullable Object input) {
        if (input == null) {
            return "";
        }
        String value = input.toString();
        // CSV Injection defense: prefix formula-triggering characters with a single quote
        if (csvInjectionDefense && !value.isEmpty() && isFormulaCharacter(value.charAt(0))) {
            value = "'" + value;
        }
        if (value.contains(String.valueOf(this.delimiter)) || value.contains("\"") || value.contains("\n") || value.contains("\r")) {
            return "\"" + value.replace("\"", "\"\"") + "\"";
        }
        return value;
    }

    private void validateUniqueColumnNames() {
        java.util.Set<String> seen = new java.util.HashSet<>();
        for (CsvColumn<T> col : columns) {
            if (!seen.add(col.getName())) {
                throw new CsvWriteException("Duplicate column name: '" + col.getName() + "'");
            }
        }
    }

    private static boolean isFormulaCharacter(char c) {
        return c == '=' || c == '+' || c == '-' || c == '@' || c == '\t' || c == '\r';
    }

}
