package io.github.dornol.excelkit.excel;

import org.apache.poi.ss.usermodel.CellStyle;
import org.jspecify.annotations.Nullable;

import java.util.Objects;

/**
 * The settings of one column, given to the {@code .column(name, fn, c -> ...)} overloads of {@link ExcelWriter},
 * {@link ExcelSheetWriter} and {@link TemplateListWriter}.
 * <p>
 * The styling methods (type, format, bold, color, width, border, validation, etc.) are inherited from
 * {@link ColumnStyleConfig}; this class closes its self type, so that chaining within the lambda returns the concrete
 * type, and adds {@link #style(CellStyle)}. One type serves every writer, so a setting kept in a variable of type
 * {@code Consumer<ColumnConfig<Row>>} can be given to any of them.
 *
 * @param <T> the row data type
 * @author dhkim
 * @since 0.13.0
 */
public class ColumnConfig<T> extends ColumnStyleConfig<T, ColumnConfig<T>> {
    @Nullable CellStyle customStyle;

    /** Creates a new column configuration with defaults. */
    public ColumnConfig() {}

    /**
     * Uses a cell style built by hand for the column, instead of the one made from the settings of this column.
     * The style must belong to the workbook the column is written to. The other settings (type, width, validation...)
     * still apply, but a format, font or color given here is not applied to the style: it is yours as it is.
     *
     * @param style the cell style to apply
     * @return this configuration for chaining
     */
    public ColumnConfig<T> style(CellStyle style) {
        this.customStyle = Objects.requireNonNull(style, "style cannot be null");
        return this;
    }
}
