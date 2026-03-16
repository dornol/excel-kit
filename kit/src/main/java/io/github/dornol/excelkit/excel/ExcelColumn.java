package io.github.dornol.excelkit.excel;

import io.github.dornol.excelkit.shared.Cursor;
import io.github.dornol.excelkit.shared.ProgressCallback;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.xssf.streaming.SXSSFCell;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import org.jspecify.annotations.Nullable;

import java.util.function.Function;
import java.util.stream.Stream;
import java.util.function.Consumer;

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
    private int columnWidth = 1;

    ExcelColumn(String name, ExcelRowFunction<T, @Nullable Object> function, CellStyle style, ExcelColumnSetter columnSetter,
                int minWidth, int maxWidth, boolean fixedWidth, String @Nullable [] dropdownOptions,
                @Nullable CellColorFunction<T> cellColorFunction, @Nullable String groupName, int outlineLevel,
                @Nullable Function<T, @Nullable String> commentFunction, @Nullable ExcelBorderStyle borderStyle, @Nullable Boolean locked,
                boolean hidden) {
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

    /**
     * Builder for constructing {@link ExcelColumn} instances using a fluent DSL-style API.
     *
     * @param <T> the row data type
     */
    public static class ExcelColumnBuilder<T> {
        private final ExcelWriter<T> writer;
        private final String name;
        private final ExcelRowFunction<T, @Nullable Object> function;
        private @Nullable ExcelDataType dataType;
        private @Nullable String dataFormat;
        private HorizontalAlignment alignment = HorizontalAlignment.CENTER;
        private @Nullable CellStyle style;
        private @Nullable ExcelColumnSetter columnSetter;
        private int @Nullable [] backgroundColor;
        private @Nullable Boolean bold;
        private @Nullable Integer fontSize;
        private int minWidthValue;
        private int maxWidthValue;
        private boolean fixedWidthValue;
        private String @Nullable [] dropdownOptions;
        private @Nullable CellColorFunction<T> cellColorFunction;
        private @Nullable String groupName;
        private int outlineLevel;
        private @Nullable Function<T, @Nullable String> commentFunction;
        private @Nullable ExcelBorderStyle borderStyle;
        private @Nullable Boolean locked;
        private boolean hiddenValue;

        ExcelColumnBuilder(ExcelWriter<T> writer, String name, ExcelRowFunction<T, @Nullable Object> function) {
            this.writer = writer;
            this.name = name;
            this.function = function;
        }

        /**
         * Sets the column's data type (used for styling and value conversion).
         */
        public ExcelColumnBuilder<T> type(ExcelDataType dataType) {
            this.dataType = dataType;
            return this;
        }

        /**
         * Sets the column's Excel cell data format.
         */
        public ExcelColumnBuilder<T> format(String dataFormat) {
            this.dataFormat = dataFormat;
            return this;
        }

        /**
         * Sets the column's horizontal text alignment.
         */
        public ExcelColumnBuilder<T> alignment(HorizontalAlignment alignment) {
            this.alignment = alignment;
            return this;
        }

        /**
         * Sets a custom {@link CellStyle} for this column.
         */
        public ExcelColumnBuilder<T> style(CellStyle style) {
            this.style = style;
            return this;
        }

        /**
         * Sets the background color for this column's cells.
         *
         * @param r Red component (0–255)
         * @param g Green component (0–255)
         * @param b Blue component (0–255)
         */
        public ExcelColumnBuilder<T> backgroundColor(int r, int g, int b) {
            this.backgroundColor = new int[]{r, g, b};
            return this;
        }

        /**
         * Sets the background color for this column's cells using a preset color.
         *
         * @param color Preset color
         */
        public ExcelColumnBuilder<T> backgroundColor(ExcelColor color) {
            return backgroundColor(color.getR(), color.getG(), color.getB());
        }

        /**
         * Sets whether this column's font should be bold.
         */
        public ExcelColumnBuilder<T> bold(boolean bold) {
            this.bold = bold;
            return this;
        }

        /**
         * Sets the font size for this column's cells.
         *
         * @param fontSize Font size in points (must be positive)
         */
        public ExcelColumnBuilder<T> fontSize(int fontSize) {
            if (fontSize <= 0) {
                throw new IllegalArgumentException("fontSize must be positive");
            }
            this.fontSize = fontSize;
            return this;
        }

        /**
         * Sets a fixed column width. The column will not auto-resize.
         *
         * @param fixedWidth Fixed width value (in Excel internal units)
         */
        public ExcelColumnBuilder<T> width(int fixedWidth) {
            this.fixedWidthValue = true;
            this.minWidthValue = fixedWidth;
            return this;
        }

        /**
         * Sets the minimum column width. Auto-resize will not shrink below this value.
         *
         * @param minWidth Minimum width value (in Excel internal units)
         */
        public ExcelColumnBuilder<T> minWidth(int minWidth) {
            this.minWidthValue = minWidth;
            return this;
        }

        /**
         * Sets the maximum column width. Auto-resize will not grow beyond this value.
         *
         * @param maxWidth Maximum width value (in Excel internal units)
         */
        public ExcelColumnBuilder<T> maxWidth(int maxWidth) {
            this.maxWidthValue = maxWidth;
            return this;
        }

        /**
         * Sets dropdown validation options for this column's cells.
         *
         * @param options The list of allowed values for the dropdown
         */
        public ExcelColumnBuilder<T> dropdown(String... options) {
            this.dropdownOptions = options;
            return this;
        }

        /**
         * Sets a per-cell conditional color function.
         * <p>
         * The function receives the resolved cell value and the row data, and returns
         * an {@link ExcelColor} to apply as the cell background, or {@code null} for no override.
         * Cell-level color takes precedence over row-level {@code rowColor}.
         *
         * @param cellColorFunction function to determine per-cell background color
         */
        public ExcelColumnBuilder<T> cellColor(CellColorFunction<T> cellColorFunction) {
            this.cellColorFunction = cellColorFunction;
            return this;
        }

        /**
         * Sets the group header name for this column.
         * <p>
         * Adjacent columns with the same group name will share a merged group header row
         * above the regular column header row.
         *
         * @param groupName the group header label
         */
        public ExcelColumnBuilder<T> group(String groupName) {
            this.groupName = groupName;
            return this;
        }

        /**
         * Sets the outline (grouping) level for this column.
         * <p>
         * Columns with an outline level > 0 can be collapsed/expanded in Excel.
         * Adjacent columns with the same outline level are grouped together.
         *
         * @param level the outline level (1-7, 0 = no outline)
         */
        public ExcelColumnBuilder<T> outline(int level) {
            if (level < 0 || level > 7) {
                throw new IllegalArgumentException("outline level must be between 0 and 7");
            }
            this.outlineLevel = level;
            return this;
        }

        /**
         * Sets a function that generates a cell comment (note) for each row.
         * <p>
         * The function receives the row data and returns the comment text,
         * or {@code null} if no comment should be added.
         *
         * @param commentFunction function to generate comment text per row
         */
        public ExcelColumnBuilder<T> comment(Function<T, @Nullable String> commentFunction) {
            this.commentFunction = commentFunction;
            return this;
        }

        /**
         * Sets the border style for this column's cells.
         * <p>
         * Overrides the default THIN border.
         *
         * @param borderStyle the border style to apply
         */
        public ExcelColumnBuilder<T> border(ExcelBorderStyle borderStyle) {
            this.borderStyle = borderStyle;
            return this;
        }

        /**
         * Sets whether this column's cells should be locked when sheet protection is enabled.
         * <p>
         * By default, all cells are locked when sheet protection is active.
         * Set to {@code false} to allow editing of this column's cells even when the sheet is protected.
         *
         * @param locked whether cells should be locked
         */
        public ExcelColumnBuilder<T> locked(boolean locked) {
            this.locked = locked;
            return this;
        }

        /**
         * Marks this column as hidden in the Excel output.
         */
        public ExcelColumnBuilder<T> hidden() {
            this.hiddenValue = true;
            return this;
        }

        /**
         * Sets whether this column should be hidden in the Excel output.
         *
         * @param hidden whether the column should be hidden
         */
        public ExcelColumnBuilder<T> hidden(boolean hidden) {
            this.hiddenValue = hidden;
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
                this.dataFormat = this.dataType.getDefaultFormat(); // apply format first
            }
            if (this.style == null) {
                this.style = ExcelStyleSupporter.cellStyle(
                        writer.getWb(), this.alignment, this.dataFormat,
                        this.backgroundColor, this.bold, this.fontSize,
                        this.borderStyle, this.locked,
                        writer.getCellStyleCache());
            }
            if (this.columnSetter == null) {
                this.columnSetter = this.dataType.getSetter();
            }
            return new ExcelColumn<>(this.name, this.function, this.style, this.columnSetter,
                    this.minWidthValue, this.maxWidthValue, this.fixedWidthValue, this.dropdownOptions,
                    this.cellColorFunction, this.groupName, this.outlineLevel,
                    this.commentFunction, this.borderStyle, this.locked, this.hiddenValue);
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
