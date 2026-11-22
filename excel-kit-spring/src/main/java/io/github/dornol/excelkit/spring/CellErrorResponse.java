package io.github.dornol.excelkit.spring;

import io.github.dornol.excelkit.core.read.CellError;
import org.jspecify.annotations.Nullable;

/**
 * JSON-friendly cell-level read error.
 *
 * @since 0.18.4
 */
public record CellErrorResponse(
        int columnIndex,
        @Nullable String headerName,
        @Nullable String cellValue,
        String message
) {
    /** Converts a cell error from a read. */
    public static CellErrorResponse from(CellError error) {
        return new CellErrorResponse(
                error.columnIndex(),
                error.headerName(),
                error.cellValue(),
                error.message()
        );
    }
}
