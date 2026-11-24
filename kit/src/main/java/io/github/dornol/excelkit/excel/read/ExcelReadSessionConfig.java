package io.github.dornol.excelkit.excel.read;

import io.github.dornol.excelkit.core.read.ReadSessionConfig;
import org.jspecify.annotations.Nullable;

/** Immutable execution configuration consumed by the Excel engine. */
record ExcelReadSessionConfig<T>(ReadSessionConfig<T> common, int sheetIndex, int headerRows,
        @Nullable String password, boolean fillMergedCells) {}
