package io.github.dornol.excelkit.excel;

import io.github.dornol.excelkit.shared.ExcelKitException;

/**
 * Exception thrown during Excel write operations.
 *
 * @author dhkim
 * @since 2025-07-19
 */
public class ExcelWriteException extends ExcelKitException {

    /** Creates an exception with the given message.
     * @param message the detail message */
    public ExcelWriteException(String message) {
        super(message);
    }

    /** Creates an exception with the given message and cause.
     * @param message the detail message
     * @param cause the underlying cause */
    public ExcelWriteException(String message, Throwable cause) {
        super(message, cause);
    }
}
