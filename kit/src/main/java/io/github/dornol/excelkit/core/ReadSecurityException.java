package io.github.dornol.excelkit.core;

/** Raised when workbook content violates a configured security policy. */
public final class ReadSecurityException extends ExcelKitException {
    /** Identifies the security rule that rejected the workbook. */
    public enum Reason {
        FORMULA, EXTERNAL_LINK, ENTRY_SIZE, TOTAL_SCAN_SIZE, COMPRESSION_RATIO, ZIP_ENTRIES
    }
    private final Reason reason;
    /** Creates an exception with a machine-readable rejection reason. */
    public ReadSecurityException(Reason reason, String message) {
        super(message);
        this.reason = java.util.Objects.requireNonNull(reason, "reason cannot be null");
    }
    /** Returns the rule that rejected the workbook. */
    public Reason reason() { return reason; }
}
