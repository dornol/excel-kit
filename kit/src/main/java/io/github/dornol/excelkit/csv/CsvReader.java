package io.github.dornol.excelkit.csv;


import io.github.dornol.excelkit.shared.CellData;
import jakarta.validation.Validator;

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

    /**
     * Constructs a CsvReader with instance supplier and optional validator.
     *
     * @param instanceSupplier A supplier to create new instances of {@code T} for each row
     * @param validator        Optional Bean Validation validator (nullable)
     */
    public CsvReader(Supplier<T> instanceSupplier, Validator validator) {
        this.instanceSupplier = Objects.requireNonNull(instanceSupplier, "instanceSupplier cannot be null");
        this.validator = validator;
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
    public CsvReadHandler<T> build(InputStream inputStream) {
        return new CsvReadHandler<>(inputStream, columns, instanceSupplier, validator);
    }
}
