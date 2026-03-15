package io.github.dornol.excelkit.csv;


import io.github.dornol.excelkit.shared.CellData;
import io.github.dornol.excelkit.shared.ProgressCallback;
import jakarta.validation.Validator;
import org.jspecify.annotations.NonNull;

import java.io.InputStream;
import java.nio.charset.Charset;
import java.nio.charset.StandardCharsets;
import java.util.ArrayList;
import java.util.List;
import java.util.Objects;
import java.util.function.BiConsumer;
import java.util.function.Supplier;

/**
 * Builder-style class for configuring CSV row readers.
 * <p>
 * {@code CsvReader} allows you to define how each CSV cell maps to your target object {@code T},
 * and optionally integrates Bean Validation support.
 * Once configuration is complete, use {@link #build(InputStream)} to create a {@link CsvReadHandler}.
 *
 * @param <T> The type of the object that represents one CSV row
 * @author dhkim
 * @since 2025-07-19
 */
public class CsvReader<T> {
    private final List<CsvReadColumn<T>> columns = new ArrayList<>();
    private final Supplier<T> instanceSupplier;
    private final Validator validator;
    private int headerRowIndex = 0;
    private char delimiter = ',';
    private Charset charset = StandardCharsets.UTF_8;
    private ProgressCallback progressCallback;
    private int progressInterval;

    /**
     * Constructs a CsvReader with instance supplier and optional validator.
     *
     * @param instanceSupplier A supplier to create new instances of {@code T} for each row
     * @param validator        Optional Bean Validation validator (nullable)
     */
    public CsvReader(@NonNull Supplier<T> instanceSupplier, Validator validator) {
        this.instanceSupplier = Objects.requireNonNull(instanceSupplier, "instanceSupplier cannot be null");
        this.validator = validator;
    }

    /**
     * Sets the zero-based row index of the header row.
     * Rows before this index will be skipped during reading.
     * Defaults to 0 (the first row).
     *
     * @param headerRowIndex The zero-based index of the header row
     * @return This CsvReader instance for chaining
     */
    public CsvReader<T> headerRowIndex(int headerRowIndex) {
        this.headerRowIndex = headerRowIndex;
        return this;
    }

    /**
     * Sets the delimiter character used to separate fields.
     * Defaults to comma ({@code ','}).
     *
     * @param delimiter The delimiter character
     * @return This CsvReader instance for chaining
     */
    public CsvReader<T> delimiter(char delimiter) {
        this.delimiter = delimiter;
        return this;
    }

    /**
     * Sets the character encoding for reading the CSV file.
     * Defaults to {@link StandardCharsets#UTF_8}.
     *
     * @param charset The charset to use
     * @return This CsvReader instance for chaining
     */
    public CsvReader<T> charset(Charset charset) {
        this.charset = charset;
        return this;
    }

    /**
     * Adds a column mapping to the internal list.
     *
     * @param column A CSV column with setter logic
     */
    void addColumn(CsvReadColumn<T> column) {
        columns.add(column);
    }

    /**
     * Adds a column mapping using a setter function.
     * Useful for schema-based column registration.
     *
     * @param setter A {@code BiConsumer} that sets a value from {@link CellData} to the row object
     * @return This CsvReader instance for chaining
     */
    public CsvReader<T> addColumn(BiConsumer<T, CellData> setter) {
        columns.add(new CsvReadColumn<>(setter));
        return this;
    }

    /**
     * Adds a name-based column mapping using a setter function.
     * The column is matched by header name instead of positional index.
     *
     * @param headerName The header name to match in the CSV file
     * @param setter     A {@code BiConsumer} that sets a value from {@link CellData} to the row object
     * @return This CsvReader instance for chaining
     */
    public CsvReader<T> addColumn(String headerName, BiConsumer<T, CellData> setter) {
        columns.add(new CsvReadColumn<>(headerName, setter));
        return this;
    }

    /**
     * Adds an index-based column mapping.
     * The column is matched by explicit 0-based column index.
     *
     * @param columnIndex 0-based column index in the CSV file
     * @param setter      A {@code BiConsumer} that sets a value from {@link CellData} to the row object
     * @return This CsvReader instance for chaining
     */
    public CsvReader<T> columnAt(int columnIndex, BiConsumer<T, CellData> setter) {
        columns.add(new CsvReadColumn<>(null, columnIndex, setter));
        return this;
    }

    /**
     * Skips one column during reading by adding a no-op column mapping.
     *
     * @return This CsvReader instance for chaining
     */
    public CsvReader<T> skipColumn() {
        columns.add(new CsvReadColumn<>((instance, cellData) -> {}));
        return this;
    }

    /**
     * Skips the specified number of columns during reading by adding no-op column mappings.
     *
     * @param count The number of columns to skip (must be non-negative)
     * @return This CsvReader instance for chaining
     * @throws IllegalArgumentException if count is negative
     */
    public CsvReader<T> skipColumns(int count) {
        if (count < 0) {
            throw new IllegalArgumentException("skipColumns count must be non-negative");
        }
        for (int i = 0; i < count; i++) {
            columns.add(new CsvReadColumn<>((instance, cellData) -> {}));
        }
        return this;
    }

    /**
     * Begins a new positional column mapping using a setter function.
     *
     * @param setter A {@code BiConsumer} that sets a value from {@link CellData} to the row object
     * @return A builder for further column configuration
     */
    public CsvReadColumn.CsvReadColumnBuilder<T> column(BiConsumer<T, CellData> setter) {
        return new CsvReadColumn.CsvReadColumnBuilder<>(this, setter);
    }

    /**
     * Begins a new name-based column mapping using a setter function.
     * The column is matched by header name instead of positional index.
     *
     * @param headerName The header name to match in the CSV file
     * @param setter     A {@code BiConsumer} that sets a value from {@link CellData} to the row object
     * @return A builder for further column configuration
     */
    public CsvReadColumn.CsvReadColumnBuilder<T> column(String headerName, BiConsumer<T, CellData> setter) {
        return new CsvReadColumn.CsvReadColumnBuilder<>(this, headerName, setter);
    }

    /**
     * Registers a progress callback that fires every {@code interval} rows during reading.
     *
     * @param interval the number of rows between each callback invocation (must be positive)
     * @param callback the callback to invoke
     * @return This CsvReader instance for chaining
     */
    public CsvReader<T> onProgress(int interval, ProgressCallback callback) {
        if (interval <= 0) {
            throw new IllegalArgumentException("progress interval must be positive");
        }
        this.progressInterval = interval;
        this.progressCallback = callback;
        return this;
    }

    /**
     * Finalizes the configuration and builds a {@link CsvReadHandler} for parsing the given CSV stream.
     *
     * @param inputStream The input stream of the CSV file
     * @return A handler to execute CSV parsing
     */
    public CsvReadHandler<T> build(@NonNull InputStream inputStream) {
        return new CsvReadHandler<>(inputStream, columns, instanceSupplier, validator,
                headerRowIndex, delimiter, charset, progressInterval, progressCallback);
    }
}
