package io.github.dornol.excelkit.excel;

import io.github.dornol.excelkit.shared.RowFunction;

/**
 * A functional interface for mapping a row of data to a cell value in Excel exports.
 * <p>
 * Extends the shared {@link RowFunction} interface.
 *
 * @param <T> The type of the row data
 * @param <R> The type of the value to be returned for the cell
 * @author dhkim
 * @since 2025-07-19
 */
@FunctionalInterface
public interface ExcelRowFunction<T, R> extends RowFunction<T, R> {
}
