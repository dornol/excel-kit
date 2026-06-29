package io.github.dornol.excelkit.spring;

import io.github.dornol.excelkit.core.CellError;
import org.jspecify.annotations.Nullable;

/**
 * JSON-friendly cell-level read error.
 */
public record CellErrorResponse(
        int columnIndex,
        @Nullable String headerName,
        @Nullable String cellValue,
        String message
) {
    public static CellErrorResponse from(CellError error) {
        return new CellErrorResponse(
                error.columnIndex(),
                error.headerName(),
                error.cellValue(),
                error.message()
        );
    }
}
