package io.github.dornol.excelkit.excel;

import io.github.dornol.excelkit.shared.Cursor;
import io.github.dornol.excelkit.shared.ProgressCallback;
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
    private final String name;
    private final ExcelRowFunction<T, @Nullable Object> function;
    private final CellStyle style;
    private final ExcelColumnSetter columnSetter;
    private final int minWidth;
    private final int maxWidth;
    private final boolean fixedWidth;
    private final String @Nullable [] dropdownOptions;
    private final @Nullable CellColorFunction<T> cellColorFunction;
    private final @Nullable String groupName;
    private final int outlineLevel;
    private final @Nullable Function<T, @Nullable String> commentFunction;
    private final @Nullable ExcelBorderStyle borderStyle;
    private final @Nullable Boolean locked;
    private final boolean hidden;
    private final @Nullable ExcelValidation validation;
    private int columnWidth = 1;

    ExcelColumn(String name, ExcelRowFunction<T, @Nullable Object> function, CellStyle style, ExcelColumnSetter columnSetter,
                int minWidth, int maxWidth, boolean fixedWidth, String @Nullable [] dropdownOptions,
                @Nullable CellColorFunction<T> cellColorFunction, @Nullable String groupName, int outlineLevel,
                @Nullable Function<T, @Nullable String> commentFunction, @Nullable ExcelBorderStyle borderStyle, @Nullable Boolean locked,
                boolean hidden, @Nullable ExcelValidation validation) {
        this.name = name;
        this.function = function;
        this.style = style;
        this.columnSetter = columnSetter;
        this.minWidth = minWidth;
        this.maxWidth = maxWidth;
        this.fixedWidth = fixedWidth;
        this.dropdownOptions = dropdownOptions;
        this.cellColorFunction = cellColorFunction;
        this.groupName = groupName;
        this.outlineLevel = outlineLevel;
        this.commentFunction = commentFunction;
        this.borderStyle = borderStyle;
        this.locked = locked;
        this.hidden = hidden;
        this.validation = validation;
        this.columnWidth = fixedWidth ? minWidth : Math.max(getLogicalLength(name), minWidth);
    }

    /**
     * Applies the column's function to extract a value from the row and cursor.
     *
     * @param rowData the current row
     * @param cursor  the current cursor (position)
     * @return the cell value
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
            cell.setCellValue("");
            return;
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
        return Math.min(MAX_COLUMN_WIDTH, logicalLength * 250 + 1024);
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
        return dropdownOptions;
    }

    @Nullable CellColorFunction<T> getCellColorFunction() {
        return cellColorFunction;
    }

    @Nullable String getGroupName() {
        return groupName;
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

    /**
     * Builder for constructing {@link ExcelColumn} instances using a fluent DSL-style API.
     *
     * @param <T> the row data type
     */
    public static class ExcelColumnBuilder<T> extends ColumnStyleConfig<T, ExcelColumnBuilder<T>> {
        private final ExcelWriter<T> writer;
        private final String name;
        private final ExcelRowFunction<T, @Nullable Object> function;
        private @Nullable CellStyle style;
        private @Nullable ExcelColumnSetter columnSetter;

        ExcelColumnBuilder(ExcelWriter<T> writer, String name, ExcelRowFunction<T, @Nullable Object> function) {
            this.writer = writer;
            this.name = name;
            this.function = function;
        }

        /**
         * Sets a custom {@link CellStyle} for this column.
         */
        public ExcelColumnBuilder<T> style(CellStyle style) {
            this.style = style;
            return this;
        }

        /**
         * Builds the column definition with all current configurations.
         */
        ExcelColumn<T> build() {
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
            return new ExcelColumn<>(this.name, this.function, this.style, this.columnSetter,
                    this.minWidth, this.maxWidth, this.fixedWidth, this.dropdownOptions,
                    this.cellColorFunction, this.groupName, this.outlineLevel,
                    this.commentFunction, this.borderStyle, this.locked, this.hidden, this.validation);
        }

        /**
         * Finalizes the current column and returns a new builder for the next column.
         */
        public ExcelColumnBuilder<T> column(String name, ExcelRowFunction<T, @Nullable Object> function) {
            this.writer.addColumn(this.build());
            return new ExcelColumnBuilder<>(writer, name, function);
        }


        /**
         * Conditionally finalizes the current column and adds a new column if the condition is true.
         *
         * @param name      the name of the new column
         * @param condition the condition that determines if the column should be added
         * @param function  the function to extract values for the new column
         * @return a new builder for the next column, or the same builder if condition is false
         */
        public ExcelColumnBuilder<T> columnIf(String name, boolean condition, ExcelRowFunction<T, @Nullable Object> function) {
            if (!condition) {
                return this;
            }
            this.writer.addColumn(this.build());
            return new ExcelColumnBuilder<>(writer, name, function);
        }

        /**
         * Finalizes the current column and adds a new column using a simple Function.
         *
         * @param name     the name of the new column
         * @param function the function to extract values for the new column
         * @return a new builder for the next column
         */
        public ExcelColumnBuilder<T> column(String name, Function<T, @Nullable Object> function) {
            this.writer.addColumn(this.build());
            return new ExcelColumnBuilder<>(writer, name, (r, c) -> function.apply(r));
        }

        /**
         * Conditionally finalizes the current column and adds a new column using a simple Function if the condition is true.
         *
         * @param name      the name of the new column
         * @param condition the condition that determines if the column should be added
         * @param function  the function to extract values for the new column
         * @return a new builder for the next column, or the same builder if condition is false
         */
        public ExcelColumnBuilder<T> columnIf(String name, boolean condition, Function<T, @Nullable Object> function) {
            if (!condition) {
                return this;
            }
            this.writer.addColumn(this.build());
            return new ExcelColumnBuilder<>(writer, name, (r, c) -> function.apply(r));
        }

        /**
         * Finalizes the current column and adds a new column with a constant value.
         *
         * @param name  the name of the new column
         * @param value the constant value for all cells in this column
         * @return a new builder for the next column
         */
        public ExcelColumnBuilder<T> constColumn(String name, @Nullable Object value) {
            this.writer.addColumn(this.build());
            return new ExcelColumnBuilder<>(writer, name, (r, c) -> value);
        }

        /**
         * Finalizes the current column and registers a progress callback on the writer.
         */
        public ExcelWriter<T> onProgress(int interval, ProgressCallback callback) {
            this.writer.addColumn(this.build());
            return this.writer.onProgress(interval, callback);
        }

        /**
         * Finalizes the current column and registers a beforeHeader callback on the writer.
         */
        public ExcelWriter<T> beforeHeader(BeforeHeaderWriter beforeHeaderWriter) {
            this.writer.addColumn(this.build());
            return this.writer.beforeHeader(beforeHeaderWriter);
        }

        /**
         * Finalizes the current column and registers an afterData callback on the writer.
         */
        public ExcelWriter<T> afterData(AfterDataWriter afterDataWriter) {
            this.writer.addColumn(this.build());
            return this.writer.afterData(afterDataWriter);
        }

        /**
         * Finalizes the current column and registers an afterAll callback on the writer.
         */
        public ExcelWriter<T> afterAll(AfterDataWriter afterAllWriter) {
            this.writer.addColumn(this.build());
            return this.writer.afterAll(afterAllWriter);
        }

        /**
         * Finalizes the column definition and writes the Excel stream with row-level post-processing.
         */
        public ExcelHandler write(Stream<T> stream, ExcelConsumer<T> consumer) {
            this.writer.addColumn(this.build());
            return this.writer.write(stream, consumer);
        }

        /**
         * Finalizes the column definition and writes the Excel stream.
         */
        public ExcelHandler write(Stream<T> stream) {
            this.writer.addColumn(this.build());
            return this.writer.write(stream);
        }

    }

}
