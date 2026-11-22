package io.github.dornol.excelkit.spring;

import io.github.dornol.excelkit.core.ExcelKitException;

/**
 * Raised when a Spring upload cannot be opened for reading.
 *
 * @since 0.18.4
 */
public class ExcelKitUploadException extends ExcelKitException {
    private static final long serialVersionUID = 1L;

    /**
     * Creates an exception with a message.
     *
     * @param message the detail message
     */
    public ExcelKitUploadException(String message) {
        super(message);
    }

    /**
     * Creates an exception with a message and a cause.
     *
     * @param message the detail message
     * @param cause   the underlying failure
     */
    public ExcelKitUploadException(String message, Throwable cause) {
        super(message, cause);
    }
}
