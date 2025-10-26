package io.github.dornol.excelkit.excel;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.xssf.streaming.SXSSFCell;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

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
class ExcelColumn<T> {
    private static final Logger log = LoggerFactory.getLogger(ExcelColumn.class);
    private static final int MAX_COLUMN_WIDTH = 255 * 256;
    private final String name;
    private final ExcelRowFunction<T, Object> function;
    private final CellStyle style;
    private final ExcelColumnSetter columnSetter;
    private int columnWidth = 1;

    ExcelColumn(String name, ExcelRowFunction<T, Object> function, CellStyle style, ExcelColumnSetter columnSetter) {
        this.name = name;
        this.function = function;
        this.style = style;
        this.columnSetter = columnSetter;
        this.columnWidth = getLogicalLength(name);
    }

    /**
     * Applies the column's function to extract a value from the row and cursor.
     *
     * @param rowData the current row
     * @param cursor  the current cursor (position)
     * @return the cell value
     */
    Object applyFunction(T rowData, ExcelCursor cursor) {
        try {
            return function.apply(rowData, cursor);
        } catch (Exception e) {
            log.error("applyFunction exception caught : {}, {} \n", rowData, cursor, e);
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
    void fitColumnWidthByValue(Object value) {
        int width = getLogicalLength(String.valueOf(value));
        this.setColumnWidth(Math.max(this.columnWidth, width));
    }

    /**
     * Writes a value into a given cell using the column's setter logic.
     */
    void setColumnData(SXSSFCell cell, Object columnData) {
        if (columnData == null) {
            cell.setCellValue("");
            return;
        }
        try {
            this.columnSetter.set(cell, columnData);
        } catch (Exception e) {
            log.warn("cast error: {}", e.getMessage());
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
            logicalLength += (ch <= 0x7F) ? 1 : 2; // ASCII: 1, 한글 등: 2
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
        return columnWidth;
    }

    /**
     * Builder for constructing {@link ExcelColumn} instances using a fluent DSL-style API.
     *
     * @param <T> the row data type
     */
    public static class ExcelColumnBuilder<T> {
        private final ExcelWriter<T> writer;
        private final String name;
        private final ExcelRowFunction<T, Object> function;
        private ExcelDataType dataType;
        private String dataFormat;
        private HorizontalAlignment alignment = HorizontalAlignment.CENTER;
        private CellStyle style;
        private ExcelColumnSetter columnSetter;

        ExcelColumnBuilder(ExcelWriter<T> writer, String name, ExcelRowFunction<T, Object> function) {
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
         * Builds the column definition with all current configurations.
         */
        private ExcelColumn<T> build() {
            if (this.dataType == null) {
                this.type(ExcelDataType.STRING);
            }
            if (this.dataFormat == null) {
                this.dataFormat = this.dataType.getDefaultFormat(); // format 먼저
            }
            if (this.style == null) {
                this.style = ExcelStyleSupporter.cellStyle(writer.getWb(), this.alignment, this.dataFormat); // format 반영됨
            }
            if (this.columnSetter == null) {
                this.columnSetter = this.dataType.getSetter();
            }
            return new ExcelColumn<>(this.name, this.function, this.style, this.columnSetter);
        }

        /**
         * Finalizes the current column and returns a new builder for the next column.
         */
        public ExcelColumnBuilder<T> column(String name, ExcelRowFunction<T, Object> function) {
            this.writer.addColumn(this.build());
            return new ExcelColumnBuilder<>(writer, name, function);
        }

        public ExcelColumnBuilder<T> column(String name, Function<T, Object> function) {
            this.writer.addColumn(this.build());
            return new ExcelColumnBuilder<>(writer, name, (r, c) -> function.apply(r));
        }

        public ExcelColumnBuilder<T> constColumn(String name, Object value) {
            this.writer.addColumn(this.build());
            return new ExcelColumnBuilder<>(writer, name, (r, c) -> value);
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
