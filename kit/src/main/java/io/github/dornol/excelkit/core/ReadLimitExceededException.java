package io.github.dornol.excelkit.core;

/** Raised when a defensive reader limit is exceeded. */
public final class ReadLimitExceededException extends ExcelKitException {
    public enum Limit { INPUT_BYTES, SHEETS, COLUMNS, CELL_CHARACTERS }
    private final Limit limit;
    private final long configured;
    private final long actual;
    public ReadLimitExceededException(Limit limit, long configured, long actual) {
        super("Read limit exceeded: " + limit + " configured=" + configured + ", actual=" + actual);
        this.limit = java.util.Objects.requireNonNull(limit, "limit cannot be null");
        this.configured = configured;
        this.actual = actual;
    }
    public Limit limit() { return limit; }
    public long configured() { return configured; }
    public long actual() { return actual; }
}
