package io.github.dornol.excelkit.excel;

/**
 * Function that determines the background color for an individual cell based on its value and row data.
 * <p>
 * When set on a column, this function is called for each cell. If it returns a non-null
 * {@link ExcelColor}, that color is applied as the cell's background, overriding both
 * column-level {@code backgroundColor} and row-level {@code rowColor}.
 *
 * @param <T> the row data type
 * @author dhkim
 */
@FunctionalInterface
public interface CellColorFunction<T> {

    /**
     * Determines the background color for a cell.
     *
     * @param cellValue the resolved cell value (may be null)
     * @param rowData   the full row data object
     * @return an {@link ExcelColor} to apply, or {@code null} for no override
     */
    ExcelColor apply(Object cellValue, T rowData);
}
