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
        List<UploadError> errors,
        UploadSummary summary
) {
    public UploadResult {
        rows = rows == null ? List.of() : List.copyOf(rows);
        errors = errors == null ? List.of() : List.copyOf(errors);
        summary = summary == null
                ? new UploadSummary(successCount + errorCount, successCount, errorCount, 0, null, -1)
                : summary;
    }

    public UploadResult(String type, int successCount, int errorCount, List<T> rows, List<UploadError> errors) {
        this(type, successCount, errorCount, rows, errors,
                new UploadSummary(successCount + errorCount, successCount, errorCount, 0, null, -1));
    }

    public static <T> UploadResult<T> read(String type, AbstractReadHandler<T> handler) {
        return read(type, handler, null, -1);
    }

    static <T> UploadResult<T> read(String type, AbstractReadHandler<T> handler, String filename, long fileSize) {
        long started = System.nanoTime();
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

        long durationMillis = java.util.concurrent.TimeUnit.NANOSECONDS.toMillis(System.nanoTime() - started);
        UploadSummary summary = new UploadSummary(rowNum.get(), rows.size(), errors.size(),
                durationMillis, filename, fileSize);
        return new UploadResult<>(type, rows.size(), errors.size(), rows, errors, summary);
    }
}
