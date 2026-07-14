package io.github.dornol.excelkit.spring;

import io.github.dornol.excelkit.core.ReadResult;

import java.util.ArrayList;
import java.util.List;
import java.util.concurrent.atomic.AtomicLong;
import java.util.function.Consumer;

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
        UploadSummary summary,
        boolean rowsTruncated,
        boolean errorsTruncated
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
                new UploadSummary(successCount + errorCount, successCount, errorCount, 0, null, -1), false, false);
    }

    public UploadResult(String type, int successCount, int errorCount, List<T> rows,
                        List<UploadError> errors, UploadSummary summary) {
        this(type, successCount, errorCount, rows, errors, summary, false, false);
    }

    public static <T> UploadResult<T> read(String type, java.util.function.Consumer<Consumer<ReadResult<T>>> reader) {
        return read(type, reader, null, -1);
    }

    static <T> UploadResult<T> read(String type, java.util.function.Consumer<Consumer<ReadResult<T>>> reader,
                                    String filename, long fileSize) {
        return read(type, reader, filename, fileSize, UploadCollectionLimits.UNLIMITED);
    }

    static <T> UploadResult<T> read(String type, java.util.function.Consumer<Consumer<ReadResult<T>>> reader,
                                    String filename, long fileSize, UploadCollectionLimits limits) {
        long started = System.nanoTime();
        List<T> rows = new ArrayList<>();
        List<UploadError> errors = new ArrayList<>();
        AtomicLong rowNum = new AtomicLong(0);
        AtomicLong successCount = new AtomicLong();
        AtomicLong errorCount = new AtomicLong();

        reader.accept(result -> {
            long currentRowNum = rowNum.incrementAndGet();
            if (result.success()) {
                successCount.incrementAndGet();
                if (limits.maxSuccessRows() < 0 || rows.size() < limits.maxSuccessRows()) rows.add(result.data());
            } else {
                errorCount.incrementAndGet();
                if (limits.maxErrors() < 0 || errors.size() < limits.maxErrors())
                    errors.add(UploadError.from(currentRowNum, result));
            }
        });

        long durationMillis = java.util.concurrent.TimeUnit.NANOSECONDS.toMillis(System.nanoTime() - started);
        UploadSummary summary = new UploadSummary(rowNum.get(), successCount.get(), errorCount.get(),
                durationMillis, filename, fileSize);
        return new UploadResult<>(type, Math.toIntExact(successCount.get()), Math.toIntExact(errorCount.get()),
                rows, errors, summary, successCount.get() > rows.size(), errorCount.get() > errors.size());
    }
}
