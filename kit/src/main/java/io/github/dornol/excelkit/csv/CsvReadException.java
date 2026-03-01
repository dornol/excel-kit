package io.github.dornol.excelkit.csv;

import io.github.dornol.excelkit.shared.ExcelKitException;

/**
 * Exception thrown during CSV read operations.
 *
 * @author dhkim
 * @since 2025-07-19
 */
public class CsvReadException extends ExcelKitException {

    public CsvReadException(String message) {
        super(message);
    }

    public CsvReadException(String message, Throwable cause) {
        super(message, cause);
    }
}
