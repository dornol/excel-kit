package io.github.dornol.excelkit.core.read;

import io.github.dornol.excelkit.core.ExcelKitException;
import java.util.Objects;

/**
 * Exception thrown by {@link AbstractReadHandler#readStrict(java.util.function.Consumer)}
 * when a row fails validation or mapping.
 * <p>
 * This exception does <b>not</b> extend {@link ExcelKitException}, so it is propagated
 * separately through handler catch blocks.
 *
 * @author dhkim
 * @since 0.3.0
 */
public class ReadAbortException extends RuntimeException {
    private static final long serialVersionUID = 1L;

    private final ReadAbortReason reason;
    private final long maxErrors;
    private final long errorCount;

    /** Creates an exception with the given message.
     * @param message the detail message */
    public ReadAbortException(String message) {
        this(message, ReadAbortReason.STRICT_FAILURE, -1, -1);
    }

    /**
     * Creates an exception with a reason and, for the error limit, the limit and the number of errors
     * seen.
     *
     * @param message    the detail message
     * @param reason     why the read aborted
     * @param maxErrors  the configured error limit, or -1 when the abort was not caused by the limit
     * @param errorCount the number of errors seen when the read aborted, or -1 when the abort was not caused by the limit
     */
    public ReadAbortException(String message, ReadAbortReason reason, long maxErrors, long errorCount) {
        super(message);
        this.reason = Objects.requireNonNull(reason, "reason cannot be null");
        this.maxErrors = maxErrors;
        this.errorCount = errorCount;
    }

    /** Returns why the read aborted. */
    public ReadAbortReason reason() {
        return reason;
    }

    /** Returns the configured error limit, or -1 when the abort was not caused by the limit. */
    public long maxErrors() {
        return maxErrors;
    }

    /**
     * Returns the number of errors seen when the read aborted, or -1 when the abort was not caused by
     * the limit.
     */
    public long errorCount() {
        return errorCount;
    }
}
