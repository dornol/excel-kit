package io.github.dornol.excelkit.spring;

import io.github.dornol.excelkit.core.ReadResult;
import io.github.dornol.excelkit.core.RowError;

import java.util.List;

/**
 * JSON-friendly row-level upload read error.
 */
public record UploadError(
        long rowNum,
        long fileRowNum,
        RowError.Type type,
        List<String> messages,
        List<CellErrorResponse> cellErrors
) {
    public UploadError {
        messages = messages == null ? List.of() : List.copyOf(messages);
        cellErrors = cellErrors == null ? List.of() : List.copyOf(cellErrors);
    }

    public static UploadError from(RowError error) {
        return new UploadError(
                error.rowNum(),
                error.fileRowNum(),
                error.type(),
                error.messages(),
                error.cellErrors().stream().map(CellErrorResponse::from).toList()
        );
    }

    static UploadError from(long rowNum, ReadResult<?> result) {
        RowError.Type type = result.cause() != null ? RowError.Type.MAPPING : RowError.Type.VALIDATION;
        return new UploadError(
                rowNum,
                result.fileRowNum(),
                type,
                result.messages(),
                result.cellErrors().stream().map(CellErrorResponse::from).toList()
        );
    }
}
