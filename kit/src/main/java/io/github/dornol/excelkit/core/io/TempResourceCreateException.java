package io.github.dornol.excelkit.core.io;

import io.github.dornol.excelkit.core.ExcelKitException;

/**
 * Exception thrown when a temporary resource (file or directory) cannot be created.
 *
 * @author dhkim
 * @since 0.1.2
 */
public class TempResourceCreateException extends ExcelKitException {
    private static final long serialVersionUID = 1L;

    /**
     * Constructs a new TempResourceCreateException with the specified cause.
     *
     * @param cause The cause of the failure
     */
    public TempResourceCreateException(Throwable cause) {
        super(cause);
    }

}
