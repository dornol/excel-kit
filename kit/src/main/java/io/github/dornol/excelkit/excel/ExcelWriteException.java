package io.github.dornol.excelkit.excel;

import io.github.dornol.excelkit.shared.ExcelKitException;

/**
 * Exception thrown during Excel write operations.
 *
 * @author dhkim
 * @since 2025-07-19
 */
public class ExcelWriteException extends ExcelKitException {

    public ExcelWriteException(String message) {
        super(message);
    }

    public ExcelWriteException(String message, Throwable cause) {
        super(message, cause);
    }
}
