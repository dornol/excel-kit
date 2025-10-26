package io.github.dornol.excelkit.csv;

/**
 * A functional interface for mapping a row of data to a CSV column value, with optional access to cursor position.
 * <p>
 * This is typically used when defining column mappings in {@code CsvWriter}, enabling both data extraction
 * and row-aware formatting (e.g. row number, alternating values, etc.).
 *
 * @param <T> The type of the row data
 * @param <R> The type of the value to be written to the CSV cell
 * @author dhkim
 * @since 2025-07-19
 */
@FunctionalInterface
public interface CsvRowFunction<T, R> {

    /**
     * Applies this function to the given row data and cursor.
     *
     * @param rowData The row object to extract data from
     * @param cursor  The current cursor indicating row position
     * @return The value to write into the CSV column
     */
    R apply(T rowData, CsvCursor cursor);
}
