package io.github.dornol.excelkit.shared;

/**
 * Base exception for all excel-kit library errors.
 * <p>
 * This allows callers to catch all library-specific exceptions with a single catch block.
 *
 * @author dhkim
 * @since 2025-07-19
 */
public class ExcelKitException extends RuntimeException {

    /** Creates an exception with the given message.
     * @param message the detail message */
    public ExcelKitException(String message) {
        super(message);
    }

    /** Creates an exception with the given message and cause.
     * @param message the detail message
     * @param cause the underlying cause */
    public ExcelKitException(String message, Throwable cause) {
        super(message, cause);
    }

    /** Creates an exception with the given cause.
     * @param cause the underlying cause */
    public ExcelKitException(Throwable cause) {
        super(cause);
    }
}
