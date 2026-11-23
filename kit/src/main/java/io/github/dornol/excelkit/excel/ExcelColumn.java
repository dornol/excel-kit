package io.github.dornol.excelkit.excel;

import io.github.dornol.excelkit.excel.feature.ExcelCellComment;
import io.github.dornol.excelkit.excel.style.CellColorFunction;
import io.github.dornol.excelkit.excel.style.ExcelBorderStyle;
import io.github.dornol.excelkit.core.read.RowFunction;
import io.github.dornol.excelkit.core.Cursor;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.xssf.streaming.SXSSFCell;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import org.jspecify.annotations.Nullable;

import java.util.HashMap;
import java.util.Map;
import java.util.Objects;
import java.util.function.Function;

/**
 * Represents a single Excel column and how its value is derived, styled, and rendered.
 * <p>
 * An {@code ExcelColumn} encapsulates:
 * - a name (used as the header),
 * - a value extractor function,
 * - a cell style,
 * - a column width calculator, and
 * - a setter function to write the value into a cell.
 *
 * @param <T> the row data type
 *
 * @author dhkim
 * @since 0.1.2
 */
public class ExcelColumn<T> {
    private static final Logger log = LoggerFactory.getLogger(ExcelColumn.class);
    private static final int MAX_COLUMN_WIDTH = 255 * 256;
    private static final int WIDTH_PER_CHAR = 250;
    private static final int WIDTH_BASE_PADDING = 1024;
    private final String name;
    private final RowFunction<T, @Nullable Object> function;
    private CellStyle style;
    private @Nullable Map<String, CellStyle> mixedStyles;
    private ExcelColumnSetter columnSetter;
    private @Nullable ExcelDataType dataType;
    private final boolean explicitDataType;
    private final boolean explicitDataFormat;
    private final int minWidth;
    private final int maxWidth;
    private final boolean fixedWidth;
    private final String @Nullable [] dropdownOptions;
    private final @Nullable CellColorFunction<T> cellColorFunction;
    private final String @Nullable [] groupNames;
    private final int outlineLevel;
    private final @Nullable Function<T, @Nullable String> commentFunction;
    private final @Nullable ExcelBorderStyle borderStyle;
    private final @Nullable Boolean locked;
    private final boolean hidden;
    private final @Nullable ExcelValidation validation;
    private final @Nullable DependentDropdown dependentDropdown;
    private final int @Nullable [] headerFontColor;
    private final int @Nullable [] headerBackgroundColor;
    private final @Nullable ExcelCellComment headerComment;
    private final int commentWidth;
    private final int commentHeight;
    private final @Nullable Object nullValue;
    private int columnWidth = 1;
    private boolean extractionFailureLogged;
    private boolean cellWriteFailureLogged;

    static <T> ExcelColumn<T> of(String name, RowFunction<T, @Nullable Object> function,
                                  @Nullable CellStyle style, ExcelColumnSetter columnSetter) {
        return new ExcelColumn<>(name, function, style, columnSetter, new ColumnStyleConfig.DefaultStyleConfig<>());
    }

    ExcelColumn(String name, RowFunction<T, @Nullable Object> function, @Nullable CellStyle style,
                ExcelColumnSetter columnSetter, ColumnStyleConfig<T, ?> config) {
        this.name = name;
        this.function = function;
        this.style = style;
        this.columnSetter = columnSetter;
        this.dataType = config.dataType;
        this.explicitDataType = config.dataType != null;
        this.explicitDataFormat = config.dataFormat != null;
        this.minWidth = config.minWidth;
        this.maxWidth = config.maxWidth;
        this.fixedWidth = config.fixedWidth;
        this.dropdownOptions = config.dropdownOptions;
        this.cellColorFunction = config.cellColorFunction;
        this.groupNames = config.groupNames;
        this.outlineLevel = config.outlineLevel;
        this.commentFunction = config.commentFunction;
        this.borderStyle = config.borderStyle;
        this.locked = config.locked;
        this.hidden = config.hidden;
        if (config.validation != null && config.dropdownOptions != null) {
            throw new IllegalArgumentException("Column '" + name + "' has both a dropdown and a validation; "
                    + "a cell has one data validation, so use one of them (listFromRange(...) makes a list)");
        }
        if (config.dependentDropdown != null && (config.dropdownOptions != null || config.validation != null)) {
            throw new IllegalArgumentException("Column '" + name + "' has a dependent dropdown and also a dropdown or a validation; "
                    + "a cell has one data validation, so use one of them");
        }
        this.validation = config.validation;
        this.dependentDropdown = config.dependentDropdown;
        this.headerFontColor = config.headerFontColor;
        this.headerBackgroundColor = config.headerBackgroundColor;
        this.headerComment = config.headerComment;
        this.commentWidth = config.commentWidth;
        this.commentHeight = config.commentHeight;
        this.nullValue = config.nullValue;
        this.columnWidth = config.fixedWidth ? config.minWidth : Math.max(getLogicalLength(name), config.minWidth);
    }

    /**
     * Applies the column's function leniently: an exception is logged and {@code null}
     * (an empty cell) is returned. Writers use {@link #applyFunction(Object, Cursor, ExcelWriteErrorPolicy)}
     * with their configured policy instead; this shortcut exists for callers that always want
     * {@link ExcelWriteErrorPolicy#LENIENT}.
     *
     * @param rowData the current row
     * @param cursor  the current cursor (position)
     * @return the cell value, or {@code null} if the function threw an exception
     */
    @Nullable Object applyFunction(T rowData, Cursor cursor) {
        return applyFunction(rowData, cursor, ExcelWriteErrorPolicy.LENIENT);
    }

    /**
     * Applies the column's function to extract a value from the row and cursor.
     * <p>
     * With {@link ExcelWriteErrorPolicy#FAIL_FAST} (the writers' default) an exception aborts the
     * write as an {@link ExcelWriteException} naming the column. With
     * {@link ExcelWriteErrorPolicy#LENIENT} the error is logged with the column name, row data, and
     * cursor, and {@code null} (an empty cell) is returned so one bad cell does not fail a bulk export.
     *
     * @param rowData the current row
     * @param cursor  the current cursor (position)
     * @param policy  how to handle an exception thrown by the function
     * @return the cell value, or {@code null} if the function threw and the policy is lenient
     * @throws ExcelWriteException if the function threw and the policy is fail-fast
     */
    @Nullable Object applyFunction(T rowData, Cursor cursor, ExcelWriteErrorPolicy policy) {
        try {
            return function.apply(rowData, cursor);
        } catch (RuntimeException e) {
            if (policy == ExcelWriteErrorPolicy.FAIL_FAST) {
                throw new ExcelWriteException("Failed to extract value for column '" + name + "'", e);
            }
            // The row itself is not logged: it may hold personal data, and a failing column repeats this per row.
            if (!extractionFailureLogged) {
                extractionFailureLogged = true;
                log.error("applyFunction exception caught for column '{}' at row {}; further failures of this column are logged at DEBUG",
                        name, cursor.getRowOfSheet(), e);
            } else {
                log.debug("applyFunction exception caught for column '{}' at row {}", name, cursor.getRowOfSheet(), e);
            }
            return null;
        }
    }

    /**
     * Sets the column's internal width value.
     */
    void setColumnWidth(int columnWidth) {
        this.columnWidth = columnWidth;
    }

    /**
     * Updates the column width based on the logical string length of a value.
     */
    void fitColumnWidthByValue(@Nullable Object value) {
        if (fixedWidth || value == null) {
            return;
        }
        int width = getLogicalLength(ExcelCellText.textOf(value));
        int candidate = Math.max(this.columnWidth, width);
        if (maxWidth > 0) candidate = Math.min(candidate, maxWidth);
        this.setColumnWidth(Math.max(candidate, minWidth));
    }

    /**
     * Writes a value into a given cell leniently: a value that cannot be stored as the column's
     * type is written as text. Writers call
     * {@link #setColumnData(SXSSFCell, Object, ExcelWriteErrorPolicy)} with their configured policy.
     */
    void setColumnData(SXSSFCell cell, @Nullable Object columnData) {
        setColumnData(cell, columnData, ExcelWriteErrorPolicy.LENIENT);
    }

    /**
     * Writes a value into a given cell using the column's setter logic.
     * With {@link ExcelWriteErrorPolicy#FAIL_FAST} a value that cannot be stored as the column's type
     * throws {@link ExcelWriteException}; with {@link ExcelWriteErrorPolicy#LENIENT} it is written as text.
     *
     * @return the style to apply to the cell
     */
    CellStyle setColumnData(SXSSFCell cell, @Nullable Object columnData, ExcelWriteErrorPolicy policy) {
        boolean placeholder = false;
        if (columnData == null) {
            if (nullValue != null) {
                columnData = nullValue;
                placeholder = true;
            } else {
                cell.setBlank();
                return style;
            }
        }
        if (!explicitDataType && dataType == null) {
            dataType = ExcelDataType.infer(columnData);
            if (dataType != null) {
                columnSetter = dataType.getSetter();
                if (!explicitDataFormat && dataType.getDefaultFormat() != null) {
                    var workbook = cell.getSheet().getWorkbook();
                    CellStyle inferredStyle = workbook.createCellStyle();
                    inferredStyle.cloneStyleFrom(style);
                    inferredStyle.setDataFormat(workbook.createDataFormat()
                            .getFormat(dataType.getDefaultFormat()));
                    style = inferredStyle;
                }
            }
        }
        if (columnData instanceof ExcelRichText) {
            cell.setCellStyle(style);   // rich text takes the column's font for the text it does not format
        }
        try {
            return writeValue(cell, columnData);
        } catch (RuntimeException e) {
            if (placeholder) {
                // An explicit nullValue (for example "-" in a numeric column) is always written, as text
                // when it does not fit the column type, whatever the error policy.
                ExcelCellText.write(cell, ExcelCellText.textOf(columnData));
                return style;
            }
            if (policy == ExcelWriteErrorPolicy.FAIL_FAST) {
                throw new ExcelWriteException("Failed to write cell value for column '" + name + "' (value "
                        + (columnData == null ? "null" : "of type " + columnData.getClass().getSimpleName())
                        + (dataType == null ? "" : ", column type " + dataType) + "): " + e.getMessage(), e);
            }
            String type = columnData.getClass().getSimpleName();
            if (!cellWriteFailureLogged) {
                cellWriteFailureLogged = true;
                log.warn("Writing value as text after a cell write failure for column '{}' (value of type {}); "
                        + "further failures of this column are logged at DEBUG", name, type, e);
            } else {
                log.debug("Writing value as text after a cell write failure for column '{}' (value of type {})", name, type, e);
            }
            ExcelCellText.write(cell, ExcelCellText.textOf(columnData));
            return style;
        }
    }

    private CellStyle writeValue(SXSSFCell cell, Object value) {
        if (!explicitDataType && dataType != null) {
            ExcelDataType actualType = ExcelDataType.infer(value);
            if (actualType != null && actualType != dataType
                    && !(dataType.isInferredNumber() && actualType.isInferredNumber())) {
                actualType.getSetter().set(cell, value);
                return styleForValue(cell, actualType, value);
            }
        }
        try {
            columnSetter.set(cell, value);
            return dataType == null ? style : styleForValue(cell, dataType, value);
        } catch (ArithmeticException e) {
            if (explicitDataType || (dataType != ExcelDataType.INTEGER && dataType != ExcelDataType.LONG)
                    || !(value instanceof Number number)) throw e;

            ExcelDataType.setNumber(cell, number);
            return mixedStyle(cell, null);
        }
    }

    private CellStyle styleForValue(SXSSFCell cell, ExcelDataType type, Object value) {
        if (explicitDataFormat) return style;
        String format = type.getDefaultFormat(value);
        if (type == dataType && Objects.equals(format, type.getDefaultFormat())) return style;
        return mixedStyle(cell, format);
    }

    private CellStyle mixedStyle(SXSSFCell cell, @Nullable String defaultFormat) {
        if (explicitDataFormat) return style;
        if (mixedStyles == null) mixedStyles = new HashMap<>();
        String format = defaultFormat == null ? "General" : defaultFormat;
        return mixedStyles.computeIfAbsent(format, key -> {
            var workbook = cell.getSheet().getWorkbook();
            CellStyle mixed = workbook.createCellStyle();
            mixed.cloneStyleFrom(style);
            mixed.setDataFormat(workbook.createDataFormat().getFormat(key));
            return mixed;
        });
    }

    /**
     * Calculates the display width for the column based on character content: an ASCII character counts as
     * one, any other character (CJK and the like) as two, and a surrogate pair (an emoji, for example) as the
     * two of one wide character.
     */
    private int getLogicalLength(String input) {
        int logicalLength = 0;
        int length = input.length();
        for (int i = 0; i < length; i++) {
            char ch = input.charAt(i);
            if (Character.isHighSurrogate(ch) && i + 1 < length && Character.isLowSurrogate(input.charAt(i + 1))) {
                i++;
                logicalLength += 2;
            } else {
                logicalLength += (ch <= 0x7F) ? 1 : 2;
            }
        }
        return Math.min(MAX_COLUMN_WIDTH, logicalLength * WIDTH_PER_CHAR + WIDTH_BASE_PADDING);
    }

    String getName() {
        return name;
    }


    CellStyle getStyle() {
        return style;
    }

    /** Whether the column's width was set with {@code width(n)} rather than fitted to its content. */
    boolean isFixedWidth() {
        return fixedWidth;
    }

    int getColumnWidth() {
        int w = columnWidth;
        if (maxWidth > 0) w = Math.min(w, maxWidth);
        return Math.max(w, minWidth);
    }

    String @Nullable [] getDropdownOptions() {
        return dropdownOptions != null ? dropdownOptions.clone() : null;
    }

    @Nullable DependentDropdown getDependentDropdown() {
        return dependentDropdown;
    }

    @Nullable CellColorFunction<T> getCellColorFunction() {
        return cellColorFunction;
    }

    /**
     * Returns group header levels (outermost first), or an empty array if this column
     * is not grouped. Never {@code null}.
     */
    String[] getGroupNames() {
        return groupNames != null ? groupNames.clone() : new String[0];
    }

    int getOutlineLevel() {
        return outlineLevel;
    }

    @Nullable Function<T, @Nullable String> getCommentFunction() {
        return commentFunction;
    }

    @Nullable ExcelBorderStyle getBorderStyle() {
        return borderStyle;
    }

    @Nullable Boolean getLocked() {
        return locked;
    }

    boolean isHidden() {
        return hidden;
    }

    @Nullable ExcelValidation getValidation() {
        return validation;
    }

    int @Nullable [] getHeaderFontColor() {
        return headerFontColor != null ? headerFontColor.clone() : null;
    }

    int @Nullable [] getHeaderBackgroundColor() {
        return headerBackgroundColor != null ? headerBackgroundColor.clone() : null;
    }

    @Nullable ExcelCellComment getHeaderComment() {
        return headerComment;
    }

    int getCommentWidth() {
        return commentWidth;
    }

    int getCommentHeight() {
        return commentHeight;
    }

    /**
     * Builder for constructing {@link ExcelColumn} instances using a fluent DSL-style API.
     *
     * @param <T> the row data type
     */
    public static class ExcelColumnBuilder<T> extends ColumnStyleConfig<T, ExcelColumnBuilder<T>> {
        private final ExcelWriter<T> writer;
        private final String name;
        private final RowFunction<T, @Nullable Object> function;
        private @Nullable CellStyle style;
        private @Nullable ExcelColumnSetter columnSetter;

        ExcelColumnBuilder(ExcelWriter<T> writer, String name, RowFunction<T, @Nullable Object> function) {
            this.writer = writer;
            this.name = name;
            this.function = Objects.requireNonNull(function, "function cannot be null");
        }

        /**
         * Sets a custom {@link CellStyle} for this column.
         *
         * @param style the cell style to apply
         * @return this instance for chaining
         */
        public ExcelColumnBuilder<T> style(CellStyle style) {
            this.style = style;
            return this;
        }

        /**
         * Builds the column definition with all current configurations.
         */
        ExcelColumn<T> build() {
            var defaults = writer.getDefaultStyleConfig();
            if (defaults != null) {
                this.applyDefaults(defaults);
            }
            if (this.style == null) {
                // The type's default format is not written back to dataFormat: the column must still know
                // whether the format was given explicitly.
                this.style = ExcelStyleSupporter.cellStyle(
                        writer.getWb(), styleParams(), writer.getCellStyleCache());
            }
            if (this.columnSetter == null) {
                this.columnSetter = this.dataType != null
                        ? this.dataType.getSetter() : ExcelDataType.STRING.getSetter();
            }
            return new ExcelColumn<>(this.name, this.function, this.style, this.columnSetter, this);
        }

    }

}
