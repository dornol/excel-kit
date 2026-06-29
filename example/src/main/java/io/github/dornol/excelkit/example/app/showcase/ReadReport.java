package io.github.dornol.excelkit.example.app.showcase;

import io.github.dornol.excelkit.example.app.dto.ProductReadDto;

import java.util.List;

public record ReadReport(
        String type,
        int successCount,
        int errorCount,
        List<ProductReadDto> rows,
        List<ReadError> errors
) {}
