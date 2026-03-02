package io.github.dornol.excelkit.csv;


import io.github.dornol.excelkit.shared.CellData;
import jakarta.validation.Validator;
import org.jspecify.annotations.NonNull;

import java.io.InputStream;
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
     * Adds a column mapping to the internal list.
     *
     * @param column A CSV column with setter logic
     */
    void addColumn(CsvReadColumn<T> column) {
        columns.add(column);
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
     * Begins a new column mapping using a setter function.
     *
     * @param setter A {@code BiConsumer} that sets a value from {@link CellData} to the row object
     * @return A builder for further column configuration
     */
    public CsvReadColumn.CsvReadColumnBuilder<T> column(BiConsumer<T, CellData> setter) {
        return new CsvReadColumn.CsvReadColumnBuilder<>(this, setter);
    }

    /**
     * Finalizes the configuration and builds a {@link CsvReadHandler} for parsing the given CSV stream.
     *
     * @param inputStream The input stream of the CSV file
     * @return A handler to execute CSV parsing
     */
    public CsvReadHandler<T> build(@NonNull InputStream inputStream) {
        return new CsvReadHandler<>(inputStream, columns, instanceSupplier, validator, headerRowIndex);
    }
}
