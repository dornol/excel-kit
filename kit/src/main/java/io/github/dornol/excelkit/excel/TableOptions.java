package io.github.dornol.excelkit.excel;

/** Excel structured-table settings. */
public record TableOptions(String name, String style, boolean showRowStripes, boolean perRolloverSheet) {
    public TableOptions {
        java.util.Objects.requireNonNull(name, "name cannot be null");
        java.util.Objects.requireNonNull(style, "style cannot be null");
    }
    public static TableOptions defaults(String name) {
        return new TableOptions(name, "TableStyleMedium2", true, true);
    }
}
