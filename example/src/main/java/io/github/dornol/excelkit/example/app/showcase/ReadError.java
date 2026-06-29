package io.github.dornol.excelkit.example.app.showcase;

import io.github.dornol.excelkit.core.CellError;

import java.util.List;

public record ReadError(
        long fileRowNum,
        List<String> messages,
        List<CellError> cellErrors
) {}
