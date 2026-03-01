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

    public ExcelKitException(String message) {
        super(message);
    }

    public ExcelKitException(String message, Throwable cause) {
        super(message, cause);
    }

    public ExcelKitException(Throwable cause) {
        super(cause);
    }
}
