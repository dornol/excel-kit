package io.github.dornol.excelkit.csv;

import io.github.dornol.excelkit.shared.CellData;
import org.jspecify.annotations.Nullable;

import java.util.function.BiConsumer;

/**
 * Represents a single CSV column binding for reading.
 * <p>
 * Supports three matching modes:
 * <ul>
 *     <li>Positional (default) — matched by insertion order</li>
 *     <li>Name-based — matched by header name via {@code headerName}</li>
 *     <li>Index-based — matched by explicit column index via {@code columnIndex}</li>
 * </ul>
 *
 * @param headerName  optional header name for name-based column matching (null for positional/index)
 * @param columnIndex explicit 0-based column index (-1 for positional/name-based)
 * @param setter      the setter function to bind a column value to a field
 * @param <T> The row data type
 * @author dhkim
 * @since 2025-07-19
 */
public record CsvReadColumn<T>(@Nullable String headerName, int columnIndex, BiConsumer<T, CellData> setter) {

    public CsvReadColumn(BiConsumer<T, CellData> setter) {
        this(null, -1, setter);
    }

    public CsvReadColumn(String headerName, BiConsumer<T, CellData> setter) {
        this(headerName, -1, setter);
    }
}
