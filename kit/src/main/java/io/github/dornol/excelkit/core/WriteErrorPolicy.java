package io.github.dornol.excelkit.core;

/**
 * Controls how the writers ({@code ExcelWriter}, {@code ExcelSheetWriter}, {@code TemplateListWriter} and
 * {@code CsvWriter}) handle a value that cannot be taken from a row or written to a cell.
 *
 * @since 0.19.0 (as {@code WriteErrorPolicy}; shared by every writer since 0.29.0)
 */
public enum WriteErrorPolicy {
    /**
     * A value extraction failure produces a blank cell or empty field (or the column's {@code nullValue}), and a
     * cell type write failure (Excel only: a CSV has no cell types) falls back to writing the value as text.
     * Failures are logged.
     */
    LENIENT,

    /**
     * Abort writing by throwing the writer's write exception ({@code ExcelWriteException} or
     * {@code CsvWriteException}) on the first value extraction or cell type write failure. This is the default.
     */
    FAIL_FAST
}
