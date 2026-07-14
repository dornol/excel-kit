package io.github.dornol.excelkit.core;

/**
 * Exception thrown by {@link AbstractReadHandler#readStrict(java.util.function.Consumer)}
 * when a row fails validation or mapping.
 * <p>
 * This exception does <b>not</b> extend {@link ExcelKitException}, so it is propagated
 * separately through handler catch blocks.
 *
 * @author dhkim
 * @since 2025-07-19
 */
public class ReadAbortException extends RuntimeException {
    private final ReadAbortReason reason;
    private final long maxErrors;
    private final long errorCount;

    /** Creates an exception with the given message.
     * @param message the detail message */
    public ReadAbortException(String message) {
        this(message, ReadAbortReason.STRICT_FAILURE, -1, -1);
    }

    public ReadAbortException(String message, ReadAbortReason reason, long maxErrors, long errorCount) {
        super(message);
        this.reason = java.util.Objects.requireNonNull(reason, "reason cannot be null");
        this.maxErrors = maxErrors;
        this.errorCount = errorCount;
    }

    public ReadAbortReason reason() {
        return reason;
    }

    public long maxErrors() {
        return maxErrors;
    }

    public long errorCount() {
        return errorCount;
    }
}
