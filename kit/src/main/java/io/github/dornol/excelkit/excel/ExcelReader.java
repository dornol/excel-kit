package io.github.dornol.excelkit.excel;

import io.github.dornol.excelkit.shared.CellData;
import jakarta.validation.Validator;

import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;
import java.util.Objects;
import java.util.function.BiConsumer;
import java.util.function.Supplier;

/**
 * Builder-style class for configuring Excel row readers.
 * <p>
 * {@code ExcelReader} allows you to define how each Excel cell maps to your target object {@code T},
 * and optionally integrates Bean Validation support.
 * Once configuration is complete, use {@link #build(InputStream)} to create a {@link ExcelReadHandler}.
 *
 * @param <T> The type of the object that represents one Excel row
 *
 * @author dhkim
 * @since 2025-07-19
 */
public class ExcelReader<T> {
    private final List<ExcelReadColumn<T>> columns = new ArrayList<>();
    private final Supplier<T> instanceSupplier;
    private final Validator validator;

    /**
     * Constructs an ExcelReader with instance supplier and optional validator.
     *
     * @param instanceSupplier A supplier to create new instances of {@code T} for each row
     * @param validator        Optional Bean Validation validator (nullable)
     */
    public ExcelReader(Supplier<T> instanceSupplier, Validator validator) {
        this.instanceSupplier = Objects.requireNonNull(instanceSupplier, "instanceSupplier cannot be null");
        this.validator = validator;
    }

    /**
     * Adds a column mapping to the internal list.
     *
     * @param column An Excel column with setter logic
     */
    void addColumn(ExcelReadColumn<T> column) {
        columns.add(column);
    }

    /**
     * Begins a new column mapping using a setter function.
     *
     * @param setter A {@code BiConsumer} that sets a value from {@link io.github.dornol.excelkit.shared.CellData} to the row object
     * @return A builder for further column configuration
     */
    public ExcelReadColumn.ExcelReadColumnBuilder<T> column(BiConsumer<T, CellData> setter) {
        return new ExcelReadColumn.ExcelReadColumnBuilder<>(this, setter);
    }

    /**
     * Finalizes the configuration and builds an {@link ExcelReadHandler} for parsing the given Excel stream.
     *
     * @param inputStream The input stream of the Excel file
     * @return A handler to execute Excel parsing
     */
    public ExcelReadHandler<T> build(InputStream inputStream) {
        return new ExcelReadHandler<>(inputStream, columns, instanceSupplier, validator);
    }
}
