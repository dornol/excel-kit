package io.github.dornol.excelkit.excel.read;

/** Controls whether a workbook read continues after one sheet fails structurally. */
public enum SheetFailurePolicy {
    CONTINUE,
    STOP
}
