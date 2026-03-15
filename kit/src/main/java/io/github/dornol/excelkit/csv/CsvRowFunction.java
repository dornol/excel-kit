package io.github.dornol.excelkit.csv;

import io.github.dornol.excelkit.shared.RowFunction;

/**
 * A functional interface for mapping a row of data to a CSV column value.
 * <p>
 * Extends the shared {@link RowFunction} interface.
 *
 * @param <T> The type of the row data
 * @param <R> The type of the value to be written to the CSV cell
 * @author dhkim
 * @since 2025-07-19
 */
@FunctionalInterface
public interface CsvRowFunction<T, R> extends RowFunction<T, R> {
}
