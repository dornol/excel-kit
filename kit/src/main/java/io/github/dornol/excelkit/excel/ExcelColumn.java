package io.github.dornol.excelkit.excel;

import io.github.dornol.excelkit.core.RowFunction;
import io.github.dornol.excelkit.core.Cursor;
import io.github.dornol.excelkit.core.ProgressCallback;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.xssf.streaming.SXSSFCell;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import org.jspecify.annotations.Nullable;

import java.util.function.Function;
import java.util.stream.Stream;

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
 * @since 2025-07-19
 */
public class ExcelColumn<T> {
    private static final Logger log = LoggerFactory.getLogger(ExcelColumn.class);
    private static final int MAX_COLUMN_WIDTH = 255 * 256;
    private static final int WIDTH_PER_CHAR = 250;
    private static final int WIDTH_BASE_PADDING = 1024;
    private final String name;
    private final RowFunction<T, @Nullable Object> function;
    private final CellStyle style;
    private final ExcelColumnSetter columnSetter;
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
    private final int @Nullable [] headerFontColor;
    private final int @Nullable [] headerBackgroundColor;
    private final @Nullable ExcelCellComment headerComment;
    private final int commentWidth;
    private final int commentHeight;
    private final @Nullable Object nullValue;
    private int columnWidth = 1;

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
        this.validation = config.validation;
        this.headerFontColor = config.headerFontColor;
        this.headerBackgroundColor = config.headerBackgroundColor;
        this.headerComment = config.headerComment;
        this.commentWidth = config.commentWidth;
        this.commentHeight = config.commentHeight;
        this.nullValue = config.nullValue;
        this.columnWidth = config.fixedWidth ? config.minWidth : Math.max(getLogicalLength(name), config.minWidth);
    }

    /**
     * Applies the column's function to extract a value from the row and cursor.
     * <p>
     * Intentionally catches exceptions and returns {@code null} (empty cell) instead of
     * propagating. In bulk exports (100K+ rows), failing the entire export for one bad cell
     * is worse than leaving it blank. Errors are logged with column name, row data, and cursor
     * for debugging.
     *
     * @param rowData the current row
     * @param cursor  the current cursor (position)
     * @return the cell value, or {@code null} if the function threw an exception
     */
    @Nullable Object applyFunction(T rowData, Cursor cursor) {
        try {
            return function.apply(rowData, cursor);
        } catch (RuntimeException e) {
            log.error("applyFunction exception caught for column '{}': row={}, cursor={}", name, rowData, cursor, e);
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
        int width = getLogicalLength(String.valueOf(value));
        int candidate = Math.max(this.columnWidth, width);
        if (maxWidth > 0) candidate = Math.min(candidate, maxWidth);
        this.setColumnWidth(Math.max(candidate, minWidth));
    }

    /**
     * Writes a value into a given cell using the column's setter logic.
     */
    void setColumnData(SXSSFCell cell, @Nullable Object columnData) {
        if (columnData == null) {
            if (nullValue != null) {
                columnData = nullValue;
            } else {
                cell.setCellValue("");
                return;
            }
        }
        try {
            this.columnSetter.set(cell, columnData);
        } catch (RuntimeException e) {
            log.warn("Failed to set cell value for column '{}': expected type mismatch (value={})", name, columnData, e);
            cell.setCellValue(String.valueOf(columnData));
        }
    }

    /**
     * Calculates the display width for the column based on character content.
     * Treats non-ASCII as double width.
     */
    private int getLogicalLength(String input) {
        int logicalLength = 0;
        for (char ch : input.toCharArray()) {
            logicalLength += (ch <= 0x7F) ? 1 : 2; // ASCII: 1, CJK etc: 2
        }
        return Math.min(MAX_COLUMN_WIDTH, logicalLength * WIDTH_PER_CHAR + WIDTH_BASE_PADDING);
    }

    String getName() {
        return name;
    }


    CellStyle getStyle() {
        return style;
    }

    int getColumnWidth() {
        int w = columnWidth;
        if (maxWidth > 0) w = Math.min(w, maxWidth);
        return Math.max(w, minWidth);
    }

    String @Nullable [] getDropdownOptions() {
        return dropdownOptions != null ? dropdownOptions.clone() : null;
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
            this.function = function;
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
            if (this.dataType == null) {
                this.type(ExcelDataType.STRING);
            }
            if (this.dataFormat == null) {
                this.dataFormat = this.dataType.getDefaultFormat();
            }
            if (this.style == null) {
                CellStyleParams params = new CellStyleParams(
                        this.alignment, this.dataFormat,
                        this.backgroundColor, this.bold, this.fontSize,
                        this.borderStyle, this.locked,
                        this.rotation,
                        this.borderTop, this.borderBottom, this.borderLeft, this.borderRight,
                        this.fontColor, this.strikethrough, this.underline,
                        this.verticalAlignment, this.wrapText, this.fontName, this.indentation
                );
                this.style = ExcelStyleSupporter.cellStyle(
                        writer.getWb(), params, writer.getCellStyleCache());
            }
            if (this.columnSetter == null) {
                this.columnSetter = this.dataType.getSetter();
            }
            return new ExcelColumn<>(this.name, this.function, this.style, this.columnSetter, this);
        }

    }

}
