package io.github.dornol.excelkit.core;

/**
 * Defensive limits for untrusted tabular input. A negative value disables that limit.
 * <p>
 * A reader applies no limits unless it is given some, and it copies its whole input to a temp file,
 * so a file from an untrusted source (an upload) should be read with {@link #UPLOAD} or limits of
 * your own, together with {@link ReadSecurityPolicy#STRICT} for Excel.
 */
public record ReadLimits(long maxInputBytes, int maxSheets, int maxColumns, int maxCellCharacters,
                         long maxTotalCells) {
    public static final ReadLimits UNLIMITED = new ReadLimits(-1, -1, -1, -1, -1);

    /**
     * A starting point for files uploaded by untrusted users: at most 50 MiB of input, 50 sheets,
     * 1,000 columns, 32,767 characters per cell (Excel's own cell limit) and 10 million cells in total.
     * Tighten it to what your application actually accepts.
     *
     * @since 0.27.0
     */
    public static final ReadLimits UPLOAD = new ReadLimits(50L * 1024 * 1024, 50, 1_000, 32_767, 10_000_000L);

    public ReadLimits(long maxInputBytes, int maxSheets, int maxColumns, int maxCellCharacters) {
        this(maxInputBytes, maxSheets, maxColumns, maxCellCharacters, -1);
    }

    public ReadLimits {
        if (maxInputBytes < -1 || maxSheets < -1 || maxColumns < -1 || maxCellCharacters < -1
                || maxTotalCells < -1)
            throw new IllegalArgumentException("read limits must be >= -1");
    }
}
