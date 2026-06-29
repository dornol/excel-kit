package io.github.dornol.excelkit.example.app.showcase;

public record ErrorReportRow(
        long fileRowNum,
        Integer columnIndex,
        String headerName,
        String cellValue,
        String message
) {}
