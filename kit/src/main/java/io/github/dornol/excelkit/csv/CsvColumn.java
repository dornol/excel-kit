package io.github.dornol.excelkit.csv;

import io.github.dornol.excelkit.core.read.RowFunction;
import io.github.dornol.excelkit.core.Cursor;
import org.jspecify.annotations.Nullable;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

/**
 * Represents a single column in a CSV export operation.
 * <p>
 * Holds the column name (used as header) and a function that extracts
 * the column's value from a row object.
 *
 * @param <T> The type of the row data object
 * @author dhkim
 * @since 0.1.2
 */
public class CsvColumn<T> {
    private static final Logger log = LoggerFactory.getLogger(CsvColumn.class);

    private final String name;
    private final RowFunction<T, @Nullable Object> function;
    private final @Nullable Object nullValue;
    private boolean failureLogged;

    /**
     * Constructs a CSV column definition.
     *
     * @param name     The name of the column (used as header)
     * @param function A function that maps a row object and cursor to a cell value
     * @throws IllegalArgumentException if name or function is null
     */
    CsvColumn(String name, RowFunction<T, @Nullable Object> function) {
        this(name, function, null);
    }

    /**
     * Constructs a CSV column definition with a value for the rows where the function gives nothing.
     *
     * @param name      The name of the column (used as header)
     * @param function  A function that maps a row object and cursor to a cell value
     * @param nullValue What to write when the value is {@code null}, or {@code null} for an empty field
     * @throws IllegalArgumentException if name or function is null
     */
    CsvColumn(String name, RowFunction<T, @Nullable Object> function, @Nullable Object nullValue) {
        if (name == null) {
            throw new IllegalArgumentException("name must not be null");
        }
        if (function == null) {
            throw new IllegalArgumentException("function must not be null");
        }
        this.name = name;
        this.function = function;
        this.nullValue = nullValue;
    }

    /**
     * Applies the value-extracting function to the given row and cursor.
     *
     * @param rowData The current row object
     * @param cursor  The cursor tracking current row index
     * @param failOnColumnError whether to throw when value extraction fails
     * @return The column's value for this row; the null value of the column when there is none or when tolerated extraction fails
     */
    @Nullable Object applyFunction(T rowData, Cursor cursor, boolean failOnColumnError) {
        try {
            Object value = function.apply(rowData, cursor);
            return value != null ? value : nullValue;
        } catch (RuntimeException e) {
            if (failOnColumnError) {
                throw new CsvWriteException("Failed to extract CSV column '" + name
                        + "' at row " + cursor.getRowOfSheet(), e);
            }
            if (!failureLogged) {
                failureLogged = true;
                log.error("Failed to apply function for column '{}' at row {}; further failures of this column are logged at DEBUG",
                        name, cursor.getRowOfSheet(), e);
            } else {
                log.debug("Failed to apply function for column '{}' at row {}", name, cursor.getRowOfSheet(), e);
            }
            return nullValue;
        }
    }

    /**
     * Returns the column name, used in the header row.
     */
    String getName() {
        return name;
    }
}
