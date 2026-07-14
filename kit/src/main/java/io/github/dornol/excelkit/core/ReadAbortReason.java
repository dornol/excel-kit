package io.github.dornol.excelkit.core;

/** Reason a read operation aborted before normal completion. */
public enum ReadAbortReason {
    STRICT_FAILURE,
    MAX_ERRORS_EXCEEDED
}
