package io.github.dornol.excelkit.spring;

import io.github.dornol.excelkit.core.ExcelKitException;

/**
 * Raised when a Spring upload cannot be opened for reading.
 */
public class ExcelKitUploadException extends ExcelKitException {

    public ExcelKitUploadException(String message) {
        super(message);
    }

    public ExcelKitUploadException(String message, Throwable cause) {
        super(message, cause);
    }
}
