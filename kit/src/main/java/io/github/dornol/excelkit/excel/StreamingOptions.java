package io.github.dornol.excelkit.excel;

/** Immutable SXSSF workbook-creation settings shared by writer entry points. */
public record StreamingOptions(int rowAccessWindowSize, boolean compressTempFiles, boolean useSharedStrings) {
    public static final StreamingOptions DEFAULT = new StreamingOptions(1000, false, false);
    public StreamingOptions {
        if (rowAccessWindowSize <= 0) throw new IllegalArgumentException("rowAccessWindowSize must be positive");
    }
}
