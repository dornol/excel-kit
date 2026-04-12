package io.github.dornol.excelkit.excel;

/**
 * Concrete column styling configuration used by {@link ExcelSheetWriter} and
 * {@link TemplateListWriter} for their {@code .column(name, fn, cfg -> ...)} overloads.
 * <p>
 * All 47 styling methods (type, format, bold, color, width, border, validation, etc.)
 * are inherited from {@link ColumnStyleConfig}. This class exists solely to close the
 * generic self-type: {@code ColumnStyleConfig<T, ColumnConfig<T>>}, so that fluent
 * chaining within the configurer lambda returns the correct concrete type.
 * <p>
 * {@link ExcelColumn.ExcelColumnBuilder} is a separate subclass of
 * {@code ColumnStyleConfig} that adds {@code style(CellStyle)} and {@code build()} —
 * it is used by {@link ExcelWriter}'s column API and is unaffected by this class.
 *
 * @param <T> the row data type
 * @author dhkim
 * @since 0.13.0
 */
public class ColumnConfig<T> extends ColumnStyleConfig<T, ColumnConfig<T>> {
    /** Creates a new column configuration with defaults. */
    public ColumnConfig() {}
}
