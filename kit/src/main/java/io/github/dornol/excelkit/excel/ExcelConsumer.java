package io.github.dornol.excelkit.excel;

/**
 * A functional interface for consuming row data during Excel export.
 * <p>
 * Used as a callback during streaming write in {@link ExcelWriter#write(java.util.stream.Stream, ExcelConsumer)},
 * typically for tracking, logging, progress, or collecting metadata.
 *
 * @param <T> The type of the row data
 *
 * @author dhkim
 * @since 2025-07-19
 */
@FunctionalInterface
public interface ExcelConsumer<T> {

    /**
     * Called for each row processed during the Excel writing process.
     *
     * @param rowData The row data object
     * @param cursor  The current cursor with row/sheet position information
     */
    void accept(T rowData, ExcelCursor cursor);

}