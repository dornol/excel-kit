package io.github.dornol.excelkit.shared;

/**
 * A functional interface for mapping a row of data to a cell value, with optional cursor access.
 * <p>
 * Shared base for both Excel and CSV column value extraction.
 *
 * @param <T> The type of the row data
 * @param <R> The type of the value to be returned
 * @author dhkim
 * @since 2025-07-19
 */
@FunctionalInterface
public interface RowFunction<T, R> {

    /**
     * Applies this function to the given row data and cursor.
     *
     * @param rowData The data for a single row
     * @param cursor  The cursor tracking the current position
     * @return The computed value
     */
    R apply(T rowData, Cursor cursor);
}
