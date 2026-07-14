package io.github.dornol.excelkit.spring;

import org.jspecify.annotations.Nullable;

/**
 * Summary metadata for an upload read operation.
 *
 * @since 0.19.0
 */
public record UploadSummary(
        long totalRows,
        long successRows,
        long errorRows,
        long durationMillis,
        @Nullable String filename,
        long fileSize
) {
}
