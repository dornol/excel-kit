package io.github.dornol.excelkit.core.read;

/**
 * Reason a read operation aborted before normal completion.
 *
 * @since 0.20.0
 */
public enum ReadAbortReason {
    STRICT_FAILURE,
    MAX_ERRORS_EXCEEDED
}
