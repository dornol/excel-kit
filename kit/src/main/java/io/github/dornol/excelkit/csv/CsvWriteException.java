package io.github.dornol.excelkit.csv;

import io.github.dornol.excelkit.core.ExcelKitException;

/**
 * Exception thrown during CSV write operations.
 *
 * @author dhkim
 * @since 2025-07-19
 */
public class CsvWriteException extends ExcelKitException {

    /** Creates an exception with the given message.
     * @param message the detail message */
    public CsvWriteException(String message) {
        super(message);
    }

    /** Creates an exception with the given message and cause.
     * @param message the detail message
     * @param cause the underlying cause */
    public CsvWriteException(String message, Throwable cause) {
        super(message, cause);
    }
}
