package io.github.dornol.excelkit.core.read;

import io.github.dornol.excelkit.core.ExcelKitException;
import java.util.Objects;

/**
 * Raised when workbook content violates a configured security policy.
 *
 * @since 0.21.0
 */
public final class ReadSecurityException extends ExcelKitException {
    private static final long serialVersionUID = 1L;

    /** Identifies the security rule that rejected the workbook. */
    public enum Reason {
        FORMULA, EXTERNAL_LINK, ENTRY_SIZE, TOTAL_SCAN_SIZE, COMPRESSION_RATIO, ZIP_ENTRIES
    }
    private final Reason reason;
    /** Creates an exception with a machine-readable rejection reason. */
    public ReadSecurityException(Reason reason, String message) {
        super(message);
        this.reason = Objects.requireNonNull(reason, "reason cannot be null");
    }
    /**
     * Creates the exception with the failure that made the scan fail closed.
     *
     * @param reason  why the read was refused
     * @param message the message
     * @param cause   what went wrong while checking, kept for diagnosis
     */
    public ReadSecurityException(Reason reason, String message, Throwable cause) {
        super(message, cause);
        this.reason = Objects.requireNonNull(reason, "reason cannot be null");
    }
    /** Returns the rule that rejected the workbook. */
    public Reason reason() { return reason; }
}
