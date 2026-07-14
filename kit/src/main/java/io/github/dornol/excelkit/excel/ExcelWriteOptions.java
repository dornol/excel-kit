package io.github.dornol.excelkit.excel;

import java.util.List;

/** Immutable settings captured when a write execution starts. */
record ExcelWriteOptions<T>(List<ExcelColumn<T>> columns, int maxRows, SheetConfig<T> sheetConfig) {
}
