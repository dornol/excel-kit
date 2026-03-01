package io.github.dornol.excelkit.excel;

import io.github.dornol.excelkit.shared.ExcelKitException;

/**
 * Exception thrown during Excel read operations.
 *
 * @author dhkim
 * @since 2025-07-19
 */
public class ExcelReadException extends ExcelKitException {

    public ExcelReadException(String message) {
        super(message);
    }

    public ExcelReadException(String message, Throwable cause) {
        super(message, cause);
    }
}
