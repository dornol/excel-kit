package io.github.dornol.excelkit.excel.read;

/**
 * Controls whether a workbook read continues after one sheet fails structurally.
 *
 * @since 0.23.0
 */
public enum SheetFailurePolicy {
    CONTINUE,
    STOP
}
