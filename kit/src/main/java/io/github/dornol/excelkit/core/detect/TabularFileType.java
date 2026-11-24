package io.github.dornol.excelkit.core.detect;

/**
 * Supported tabular input container types.
 * <p>
 * {@link #XLS} stands for every OLE2 container: an Excel 97-2003 workbook and also a password-protected {@code .xlsx}.
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
