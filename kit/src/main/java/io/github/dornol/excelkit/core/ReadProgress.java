package io.github.dornol.excelkit.core;

import java.time.Duration;

/** Immutable progress snapshot for a reader execution. */
public record ReadProgress(long processedRows, long successRows, long errorRows,
                           int sheetIndex, long totalRows, Duration elapsed) {}
