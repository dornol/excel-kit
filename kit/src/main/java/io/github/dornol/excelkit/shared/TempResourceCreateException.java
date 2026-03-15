package io.github.dornol.excelkit.shared;

/**
 * Exception thrown when a temporary resource (file or directory) cannot be created.
 *
 * @author dhkim
 * @since 2025-07-19
 */
public class TempResourceCreateException extends ExcelKitException {

    /**
     * Constructs a new TempResourceCreateException with the specified cause.
     *
     * @param cause The cause of the failure
     */
    public TempResourceCreateException(Throwable cause) {
        super(cause);
    }

}
