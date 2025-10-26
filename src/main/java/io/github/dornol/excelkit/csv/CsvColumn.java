package io.github.dornol.excelkit.csv;

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
 * @since 2025-07-19
 */
class CsvColumn<T> {
    private static final Logger log = LoggerFactory.getLogger(CsvColumn.class);

    private final String name;
    private final CsvRowFunction<T, Object> function;

    /**
     * Constructs a CSV column definition.
     *
     * @param name     The name of the column (used as header)
     * @param function A function that maps a row object and cursor to a cell value
     * @throws IllegalArgumentException if name or function is null
     */
    CsvColumn(String name, CsvRowFunction<T, Object> function) {
        if (name == null) {
            throw new IllegalArgumentException("name must not be null");
        }
        if (function == null) {
            throw new IllegalArgumentException("function must not be null");
        }
        this.name = name;
        this.function = function;
    }

    /**
     * Applies the value-extracting function to the given row and cursor.
     *
     * @param rowData The current row object
     * @param cursor  The cursor tracking current row index
     * @return The column's value for this row, or null if evaluation fails
     */
    Object applyFunction(T rowData, CsvCursor cursor) {
        try {
            return function.apply(rowData, cursor);
        } catch (Exception e) {
            log.error("Failed to apply function for column '{}' at row {}", name, cursor.getRowOfSheet(), e);
            return null;
        }
    }

    /**
     * Returns the column name, used in the header row.
     */
    String getName() {
        return name;
    }
}
