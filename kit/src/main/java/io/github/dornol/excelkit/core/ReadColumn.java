package io.github.dornol.excelkit.core;

import org.jspecify.annotations.Nullable;

import java.util.function.BiConsumer;

/**
 * Represents a single column binding for reading (Excel or CSV).
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
 * @param isRequired  if true, blank/empty cells will produce a validation error
 * @param <T> the row data type
 * @author dhkim
 * @since 0.14.0
 */
public record ReadColumn<T>(@Nullable String headerName, int columnIndex,
                             BiConsumer<T, CellData> setter, boolean isRequired) {

    /** Creates a positional column binding (matched by column index order).
     * @param setter the setter function */
    public ReadColumn(BiConsumer<T, CellData> setter) {
        this(null, -1, setter, false);
    }

    /** Creates a name-based column binding.
     * @param headerName the header name to match
     * @param setter the setter function */
    public ReadColumn(String headerName, BiConsumer<T, CellData> setter) {
        this(headerName, -1, setter, false);
    }

    /** Creates a column binding with explicit index.
     * @param headerName the header name (nullable)
     * @param columnIndex 0-based column index (-1 for positional/name-based)
     * @param setter the setter function */
    public ReadColumn(@Nullable String headerName, int columnIndex, BiConsumer<T, CellData> setter) {
        this(headerName, columnIndex, setter, false);
    }

    /**
     * Returns a copy of this column marked as required.
     * When a required column has a blank/empty cell value, the row will be marked as failed.
     *
     * @return a new ReadColumn with isRequired=true
     */
    public ReadColumn<T> required() {
        return new ReadColumn<>(headerName, columnIndex, setter, true);
    }
}
