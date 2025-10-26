package io.github.dornol.excelkit.excel;


/**
 * A functional interface for mapping a row of data to a cell value, optionally using cursor information.
 * <p>
 * This is similar to {@link java.util.function.BiFunction}, but specialized for Excel export use cases.
 * It allows the column to compute a value based on both the row content and the current cursor position
 * (such as row number or sheet number).
 *
 * @param <T> The type of the row data
 * @param <R> The type of the value to be returned for the cell
 *
 * @author dhkim
 * @since 2025-07-19
 */
@FunctionalInterface
public interface ExcelRowFunction<T, R> {

    /**
     * Applies this function to the given row data and cursor.
     *
     * @param rowData The data for a single row
     * @param cursor  The cursor tracking the current sheet / row position
     * @return The value to write into the cell
     */
    R apply(T rowData, ExcelCursor cursor);
}
