package io.github.dornol.excelkit.core.read;

import java.time.Duration;

/**
 * Immutable progress snapshot for a reader execution.
 *
 * @since 0.21.0
 */
public record ReadProgress(long processedRows, long successRows, long errorRows,
                           int sheetIndex, long totalRows, Duration elapsed,
                           boolean completed, boolean cancelled) {}
