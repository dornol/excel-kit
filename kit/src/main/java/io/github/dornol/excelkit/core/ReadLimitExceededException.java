package io.github.dornol.excelkit.core;

import java.util.Objects;

/** Raised when a defensive reader limit is exceeded. */
public final class ReadLimitExceededException extends ExcelKitException {
    private static final long serialVersionUID = 1L;

    public enum Limit { INPUT_BYTES, SHEETS, COLUMNS, CELL_CHARACTERS, TOTAL_CELLS }
    private final Limit limit;
    private final long configured;
    private final long actual;
    public ReadLimitExceededException(Limit limit, long configured, long actual) {
        super("Read limit exceeded: " + limit + " configured=" + configured + ", actual=" + actual);
        this.limit = Objects.requireNonNull(limit, "limit cannot be null");
        this.configured = configured;
        this.actual = actual;
    }
    public Limit limit() { return limit; }
    public long configured() { return configured; }
    /**
     * Returns the observed value, not necessarily the full size of a rejected input.
     * Stream byte limits report the first byte beyond the budget; byte counts
     * saturate at {@link Long#MAX_VALUE} if the observed count exceeds that value.
     */
    public long actual() { return actual; }
}
