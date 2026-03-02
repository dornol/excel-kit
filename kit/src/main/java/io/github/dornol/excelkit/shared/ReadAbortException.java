package io.github.dornol.excelkit.shared;

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

    public ReadAbortException(String message) {
        super(message);
    }
}
