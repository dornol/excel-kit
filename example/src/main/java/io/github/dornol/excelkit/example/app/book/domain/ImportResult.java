package io.github.dornol.excelkit.example.app.book.domain;

import java.util.List;

/**
 * Domain-level result for a single imported row.
 * Decouples the application layer from excel-kit's ReadResult.
 */
public record ImportResult<T>(T data, boolean success, List<String> messages) {
}
