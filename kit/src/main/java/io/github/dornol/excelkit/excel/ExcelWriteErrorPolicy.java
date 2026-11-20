package io.github.dornol.excelkit.excel;

/**
 * Controls how Excel writers handle row value extraction and cell value write failures.
 *
 * @since 0.19.0
 */
public enum ExcelWriteErrorPolicy {
    /**
     * Value extraction failures produce a blank cell, and cell type write failures
     * fall back to writing the value as text. Failures are logged.
     */
    LENIENT,

    /**
     * Abort writing by throwing {@link ExcelWriteException} on the first value
     * extraction or cell type write failure. This is the default.
     */
    FAIL_FAST
}
