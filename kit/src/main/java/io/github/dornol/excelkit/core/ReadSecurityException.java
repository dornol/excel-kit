package io.github.dornol.excelkit.core;

/** Raised when workbook content violates a configured security policy. */
public final class ReadSecurityException extends ExcelKitException {
    public enum Reason { FORMULA, EXTERNAL_LINK, ENTRY_SIZE, TOTAL_SCAN_SIZE, COMPRESSION_RATIO }
    private final Reason reason;
    public ReadSecurityException(Reason reason, String message) {
        super(message);
        this.reason = java.util.Objects.requireNonNull(reason, "reason cannot be null");
    }
    public Reason reason() { return reason; }
}
