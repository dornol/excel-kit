package io.github.dornol.excelkit.excel;

import java.util.Objects;

/** Excel structured-table settings. */
public record TableOptions(String name, String style, boolean showRowStripes, boolean perRolloverSheet) {
    public TableOptions {
        Objects.requireNonNull(name, "name cannot be null");
        Objects.requireNonNull(style, "style cannot be null");
    }
    public static TableOptions defaults(String name) {
        return new TableOptions(name, "TableStyleMedium2", true, true);
    }
}
