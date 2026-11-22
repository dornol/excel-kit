package io.github.dornol.excelkit.excel.read;

import io.github.dornol.excelkit.core.ExcelKitException;

/**
 * Exception thrown during Excel read operations.
 *
 * @author dhkim
 * @since 0.2.0
 */
public class ExcelReadException extends ExcelKitException {
    private static final long serialVersionUID = 1L;

    /** Creates an exception with the given message.
     * @param message the detail message */
    public ExcelReadException(String message) {
        super(message);
    }

    /** Creates an exception with the given message and cause.
     * @param message the detail message
     * @param cause the underlying cause */
    public ExcelReadException(String message, Throwable cause) {
        super(message, cause);
    }
}
