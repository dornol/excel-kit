package io.github.dornol.excelkit.csv;

import io.github.dornol.excelkit.shared.ExcelKitException;

/**
 * Exception thrown during CSV write operations.
 *
 * @author dhkim
 * @since 2025-07-19
 */
public class CsvWriteException extends ExcelKitException {

    public CsvWriteException(String message) {
        super(message);
    }

    public CsvWriteException(String message, Throwable cause) {
        super(message, cause);
    }
}
