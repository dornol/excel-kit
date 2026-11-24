package io.github.dornol.excelkit.excel;

import java.util.Map;

import org.apache.poi.xssf.streaming.SXSSFWorkbook;

import org.apache.poi.ss.usermodel.CellStyle;

import io.github.dornol.excelkit.core.Cursor;
import io.github.dornol.excelkit.core.read.RowFunction;
import org.jspecify.annotations.Nullable;

import java.util.Objects;
import java.util.function.Consumer;
import java.util.function.Function;

/**
 * The columns every Excel writer of rows has in common: {@link ExcelWriter} and {@link ExcelSheetWriter}, which are
 * {@link SheetWriter}s (a sheet of their own, with settings and rollover), and {@link TemplateListWriter}, which fills a
 * list into a sheet of a template that already exists and so has no sheet settings.
 * <p>
 * A column is configured with a {@code Consumer<ColumnConfig<T>>} on all of them, so a setting kept in a variable works on
 * every writer. The type cannot be extended outside this package.
 *
 * @param <T>    the row data type
 * @param <SELF> the concrete writer type, for fluent method chaining
 * @author dhkim
 * @since 0.29.0
 */
@SuppressWarnings("unchecked")
public abstract class ColumnsWriter<T, SELF extends ColumnsWriter<T, SELF>> {

    ColumnsWriter() { }

    SELF self() {
        return (SELF) this;
    }

    /** Builds the column the writer will write: its styles come from the writer's workbook. */
    abstract ExcelColumn<T> buildColumn(String name, RowFunction<T, @Nullable Object> function,
                                        @Nullable ColumnConfig<T> config);

    /** Adds a built column to the sheet. */
    abstract void addColumn(ExcelColumn<T> column);

    private SELF register(String name, RowFunction<T, @Nullable Object> function, @Nullable Consumer<ColumnConfig<T>> configurer) {
        Objects.requireNonNull(name, "name cannot be null");
        Objects.requireNonNull(function, "function cannot be null");
        ColumnConfig<T> config = null;
        if (configurer != null) {
            config = new ColumnConfig<>();
            configurer.accept(config);
        }
        addColumn(buildColumn(name, function, config));
        return self();
    }

    /**
     * Adds a column using a simple function.
     *
     * @param name     the column header
     * @param function function to extract the cell value
     * @return this writer for chaining
     */
    public SELF column(String name, Function<T, @Nullable Object> function) {
        return column(name, function, (Consumer<ColumnConfig<T>>) null);
    }

    /**
     * Adds a column using a simple function, with settings such as type, format, width or color.
     *
     * <pre>{@code
     * writer.column("Price", Book::price, c -> c.type(ExcelDataType.INTEGER).format("#,##0"));
     * }</pre>
     *
     * @param name       the column header
     * @param function   function to extract the cell value
     * @param configurer sets the column's properties, or {@code null} for none
     * @return this writer for chaining
     */
    public SELF column(String name, Function<T, @Nullable Object> function, @Nullable Consumer<ColumnConfig<T>> configurer) {
        Objects.requireNonNull(function, "function cannot be null");
        return register(name, (r, c) -> function.apply(r), configurer);
    }

    /**
     * Adds a column whose function also receives the row's {@link Cursor}, for a value that depends on the position.
     *
     * @param name     the column header
     * @param function function to extract the cell value
     * @return this writer for chaining
     */
    public SELF column(String name, RowFunction<T, @Nullable Object> function) {
        return column(name, function, (Consumer<ColumnConfig<T>>) null);
    }

    /**
     * Adds a column whose function also receives the row's {@link Cursor}, with settings.
     *
     * @param name       the column header
     * @param function   function to extract the cell value
     * @param configurer sets the column's properties, or {@code null} for none
     * @return this writer for chaining
     */
    public SELF column(String name, RowFunction<T, @Nullable Object> function, @Nullable Consumer<ColumnConfig<T>> configurer) {
        return register(name, function, configurer);
    }

    /**
     * Conditionally adds a column using a simple function. If {@code condition} is false, nothing is added.
     *
     * @param name      the column header
     * @param condition whether to include this column
     * @param function  function to extract the cell value
     * @return this writer for chaining
     */
    public SELF columnIf(String name, boolean condition, Function<T, @Nullable Object> function) {
        return condition ? column(name, function) : self();
    }

    /**
     * Conditionally adds a column using a simple function, with settings. If {@code condition} is false, nothing is added.
     *
     * @param name       the column header
     * @param condition  whether to include this column
     * @param function   function to extract the cell value
     * @param configurer sets the column's properties, or {@code null} for none
     * @return this writer for chaining
     */
    public SELF columnIf(String name, boolean condition, Function<T, @Nullable Object> function,
                         @Nullable Consumer<ColumnConfig<T>> configurer) {
        return condition ? column(name, function, configurer) : self();
    }

    /**
     * Conditionally adds a column whose function also receives the row's {@link Cursor}.
     *
     * @param name      the column header
     * @param condition whether to include this column
     * @param function  function to extract the cell value
     * @return this writer for chaining
     */
    public SELF columnIf(String name, boolean condition, RowFunction<T, @Nullable Object> function) {
        return condition ? column(name, function) : self();
    }

    /**
     * Conditionally adds a column whose function also receives the row's {@link Cursor}, with settings.
     *
     * @param name       the column header
     * @param condition  whether to include this column
     * @param function   function to extract the cell value
     * @param configurer sets the column's properties, or {@code null} for none
     * @return this writer for chaining
     */
    public SELF columnIf(String name, boolean condition, RowFunction<T, @Nullable Object> function,
                         @Nullable Consumer<ColumnConfig<T>> configurer) {
        return condition ? column(name, function, configurer) : self();
    }

    /**
     * Adds a column with a constant value for all rows.
     *
     * @param name  the column header
     * @param value the constant value written in every row
     * @return this writer for chaining
     */
    public SELF constColumn(String name, @Nullable Object value) {
        return constColumn(name, value, null);
    }

    /**
     * Adds a column with a constant value for all rows, with settings.
     *
     * @param name       the column header
     * @param value      the constant value written in every row
     * @param configurer sets the column's properties, or {@code null} for none
     * @return this writer for chaining
     */
    public SELF constColumn(String name, @Nullable Object value, @Nullable Consumer<ColumnConfig<T>> configurer) {
        return column(name, (RowFunction<T, @Nullable Object>) (r, c) -> value, configurer);
    }

    /**
     * Conditionally adds a column with a constant value for all rows. If {@code condition} is false, nothing is added.
     *
     * @param name      the column header
     * @param condition whether to include this column
     * @param value     the constant value written in every row
     * @return this writer for chaining
     */
    public SELF constColumnIf(String name, boolean condition, @Nullable Object value) {
        return condition ? constColumn(name, value) : self();
    }

    /**
     * Conditionally adds a column with a constant value for all rows, with settings.
     *
     * @param name       the column header
     * @param condition  whether to include this column
     * @param value      the constant value written in every row
     * @param configurer sets the column's properties, or {@code null} for none
     * @return this writer for chaining
     */
    public SELF constColumnIf(String name, boolean condition, @Nullable Object value,
                              @Nullable Consumer<ColumnConfig<T>> configurer) {
        return condition ? constColumn(name, value, configurer) : self();
    }

    /**
     * Adds a 1-based sequential row-number column (it spans all sheets when the writer rolls over to more).
     *
     * @param name the column header name (e.g. {@code "No."})
     * @return this writer for chaining
     * @since 0.16.11
     */
    public SELF rowNumberColumn(String name) {
        return column(name, (row, cursor) -> cursor.getCurrentTotal(), c -> c.type(ExcelDataType.LONG));
    }

    /**
     * Builds a column from a function and an optional configuration: applies the sheet's default style, resolves
     * the data format and the cell style (shared through {@code cellStyleCache}) and picks the setter of the data
     * type, which is plain text when none was given.
     */
    static <T> ExcelColumn<T> buildColumn(SXSSFWorkbook wb, Map<String, CellStyle> cellStyleCache,
                                          ColumnStyleConfig.@Nullable DefaultStyleConfig<T> defaultStyle, String name,
                                          RowFunction<T, @Nullable Object> function, @Nullable ColumnConfig<T> config) {
        Objects.requireNonNull(function, "function cannot be null");
        ColumnStyleConfig<T, ?> c = config != null ? config : new ColumnConfig<>();
        if (defaultStyle != null) {
            c.applyDefaults(defaultStyle);
        }
        ExcelDataType dataType = c.dataType;
        CellStyle style = config != null && config.customStyle != null
                ? config.customStyle
                : ExcelStyleSupporter.cellStyle(wb, c.styleParams(), cellStyleCache);

        return new ExcelColumn<>(name, function, style,
                dataType != null ? dataType.getSetter() : ExcelDataType.STRING.getSetter(), c)
                .sharingStylesOf(cellStyleCache);
    }
}
