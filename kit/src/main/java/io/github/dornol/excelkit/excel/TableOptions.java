package io.github.dornol.excelkit.excel;

import java.util.Objects;

/**
 * Excel structured-table settings.
 *
 * @since 0.21.0
 */
public record TableOptions(String name, String style, boolean showRowStripes, boolean perRolloverSheet) {
    public TableOptions {
        Objects.requireNonNull(name, "name cannot be null");
        Objects.requireNonNull(style, "style cannot be null");
    }
    /**
     * Returns options for a table with the given name: style {@code TableStyleMedium2}, row stripes
     * on, and one table per rolled-over sheet.
     */
    public static TableOptions defaults(String name) {
        return new TableOptions(name, "TableStyleMedium2", true, true);
    }
}
