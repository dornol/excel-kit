package io.github.dornol.excelkit.core.detect;

/**
 * Supported tabular input container types.
 *
 * @since 0.21.0
 */
public enum TabularFileType {
    XLSX(true), XLS(false), CSV(true), UNKNOWN(false);
    private final boolean readable;
    TabularFileType(boolean readable) { this.readable = readable; }
    /** Returns whether excel-kit can read this type: true for {@link #XLSX} and {@link #CSV}. */
    public boolean isReadable() { return readable; }
}
