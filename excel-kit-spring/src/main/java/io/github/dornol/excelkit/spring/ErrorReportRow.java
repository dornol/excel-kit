package io.github.dornol.excelkit.spring;

import org.jspecify.annotations.Nullable;

/**
 * Flat row used for CSV/XLSX read-error report downloads.
 */
public record ErrorReportRow(
        long rowNum,
        long fileRowNum,
        @Nullable Integer columnIndex,
        @Nullable String headerName,
        @Nullable String cellValue,
        String message,
        String rawValues
) {
}
