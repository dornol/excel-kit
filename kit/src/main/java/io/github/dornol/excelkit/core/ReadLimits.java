package io.github.dornol.excelkit.core;

/** Defensive limits for untrusted tabular input. A negative value disables that limit. */
public record ReadLimits(long maxInputBytes, int maxSheets, int maxColumns, int maxCellCharacters) {
    public static final ReadLimits UNLIMITED = new ReadLimits(-1, -1, -1, -1);
    public ReadLimits {
        if (maxInputBytes < -1 || maxSheets < -1 || maxColumns < -1 || maxCellCharacters < -1)
            throw new IllegalArgumentException("read limits must be >= -1");
    }
}
