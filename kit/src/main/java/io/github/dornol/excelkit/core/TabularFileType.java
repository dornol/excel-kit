package io.github.dornol.excelkit.core;

/** Supported tabular input container types. */
public enum TabularFileType {
    XLSX(true), XLS(false), CSV(true), UNKNOWN(false);
    private final boolean readable;
    TabularFileType(boolean readable) { this.readable = readable; }
    public boolean isReadable() { return readable; }
}
