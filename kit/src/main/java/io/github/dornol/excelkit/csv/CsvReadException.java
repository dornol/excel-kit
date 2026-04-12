package io.github.dornol.excelkit.csv;

import io.github.dornol.excelkit.core.ExcelKitException;

/**
 * Exception thrown during CSV read operations.
 *
 * @author dhkim
 * @since 2025-07-19
 */
public class CsvReadException extends ExcelKitException {

    /** Creates an exception with the given message.
     * @param message the detail message */
    public CsvReadException(String message) {
        super(message);
    }

    /** Creates an exception with the given message and cause.
     * @param message the detail message
     * @param cause the underlying cause */
    public CsvReadException(String message, Throwable cause) {
        super(message, cause);
    }
}
