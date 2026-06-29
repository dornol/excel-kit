package io.github.dornol.excelkit.spring;

import io.github.dornol.excelkit.core.AbstractReadHandler;

import java.util.ArrayList;
import java.util.List;
import java.util.concurrent.atomic.AtomicLong;

/**
 * JSON-friendly upload read result containing successful rows and structured errors.
 *
 * @param <T> successful row type
 */
public record UploadResult<T>(
        String type,
        int successCount,
        int errorCount,
        List<T> rows,
        List<UploadError> errors
) {
    public UploadResult {
        rows = rows == null ? List.of() : List.copyOf(rows);
        errors = errors == null ? List.of() : List.copyOf(errors);
    }

    public static <T> UploadResult<T> read(String type, AbstractReadHandler<T> handler) {
        List<T> rows = new ArrayList<>();
        List<UploadError> errors = new ArrayList<>();
        AtomicLong rowNum = new AtomicLong(0);

        handler.read(result -> {
            long currentRowNum = rowNum.incrementAndGet();
            if (result.success()) {
                rows.add(result.data());
            } else {
                errors.add(UploadError.from(currentRowNum, result));
            }
        });

        return new UploadResult<>(type, rows.size(), errors.size(), rows, errors);
    }
}
