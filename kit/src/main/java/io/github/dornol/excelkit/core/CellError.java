package io.github.dornol.excelkit.core;

import org.jspecify.annotations.Nullable;

/**
 * Structured cell-level error information for row read failures.
 *
 * @param columnIndex 0-based column index in the source file
 * @param headerName header name resolved for the column, if available
 * @param cellValue formatted cell value that caused the error
 * @param message human-readable error message
 * @since 0.18.0
 */
public record CellError(
        int columnIndex,
        @Nullable String headerName,
        @Nullable String cellValue,
        String message
) {
}
