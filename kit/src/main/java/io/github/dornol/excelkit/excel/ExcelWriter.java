package io.github.dornol.excelkit.excel;

import io.github.dornol.excelkit.shared.Cursor;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.streaming.SXSSFCell;
import org.apache.poi.xssf.streaming.SXSSFRow;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFColor;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.function.Function;
import java.util.stream.Stream;

import static io.github.dornol.excelkit.excel.ExcelStyleSupporter.titleStyle;

/**
 * ExcelWriter is a utility class for generating large Excel files using Apache POI's SXSSFWorkbook.
 * Supports streaming writes, column configuration, style customization, and sheet auto-splitting.
 *
 * @author dhkim
 * @param <T> The data type of each row to be written into the Excel file
 * @since 2025-07-19
 */
public class ExcelWriter<T> implements AutoCloseable {
    private final SXSSFWorkbook wb;
    private final List<ExcelColumn<T>> columns = new ArrayList<>();
    private final int maxRowsOfSheet;
    private final CellStyle headerStyle;
    private final Map<String, CellStyle> cellStyleCache = new HashMap<>();
    private String title;
    private CellStyle titleStyle;

    private SXSSFSheet sheet;
    private Cursor cursor;


    /**
     * Constructs an ExcelWriter with a custom header color and maximum rows per sheet.
     *
     * @param r               Red component of the header color (0–255)
     * @param g               Green component of the header color (0–255)
     * @param b               Blue component of the header color (0–255)
     * @param maxRowsOfSheet  Maximum number of rows allowed per sheet before creating a new one
     */
    public ExcelWriter(int r, int g, int b, int maxRowsOfSheet) {
        this.wb = new SXSSFWorkbook(1000);
        this.maxRowsOfSheet = maxRowsOfSheet;
        this.headerStyle = ExcelStyleSupporter.headerStyle(wb, new XSSFColor(new byte[]{(byte) r, (byte) g, (byte) b}));
    }

    /**
     * Constructs an ExcelWriter with white header color and custom sheet row limit.
     *
     * @param maxRowsOfSheet Maximum number of rows per sheet
     */
    public ExcelWriter(int maxRowsOfSheet) {
        this(255, 255, 255, maxRowsOfSheet);
    }

    /**
     * Constructs an ExcelWriter with custom header color and default max 1,000,000 rows per sheet.
     *
     * @param r Red component of header color
     * @param g Green component of header color
     * @param b Blue component of header color
     */
    public ExcelWriter(int r, int g, int b) {
        this(r, g, b, 1_000_000);
    }

    /**
     * Constructs an ExcelWriter with a default white header and default max 1,000,000 rows per sheet.
     */
    public ExcelWriter() {
        this(255, 255, 255, 1_000_000);
    }

    /**
     * Sets the title for the Excel sheet with default font size and color.
     *
     * @param title The title text to display at the top
     * @return Current ExcelWriter instance for chaining
     */
    public ExcelWriter<T> title(String title) {
        return title(title, 0, IndexedColors.BLACK);
    }

    /**
     * Sets the title for the Excel sheet with a specified font size.
     *
     * @param title    The title text to display at the top
     * @param fontSize Font size in points
     * @return Current ExcelWriter instance for chaining
     */
    public ExcelWriter<T> title(String title, int fontSize) {
        return title(title, fontSize, IndexedColors.BLACK);
    }

    /**
     * Sets the title for the Excel sheet with a specified font size and color.
     *
     * @param title    The title text to display at the top
     * @param fontSize Font size in points
     * @param color    The text color
     * @return Current ExcelWriter instance for chaining
     */
    public ExcelWriter<T> title(String title, int fontSize, IndexedColors color) {
        if (this.title != null) {
            throw new ExcelWriteException("title setting already exists");
        }

        this.title = title;
        this.titleStyle = titleStyle(
                this.wb,
                HorizontalAlignment.CENTER,
                color,
                fontSize);

        return this;
    }

    /**
     * Adds an already-built column to the column list.
     *
     * @param column The ExcelColumn to add
     */
    void addColumn(ExcelColumn<T> column) {
        this.columns.add(column);
    }

    /**
     * Begins building a new column using a custom row function with cursor.
     *
     * @param name     Column header name
     * @param function Function to extract cell value from row with cursor
     * @return Column builder
     */
    public ExcelColumn.ExcelColumnBuilder<T> column(String name, ExcelRowFunction<T, Object> function) {
        return new ExcelColumn.ExcelColumnBuilder<>(this, name, function);
    }

    /**
     * Begins building a new column using a basic row-mapping function.
     *
     * @param name     Column header name
     * @param function Function to extract cell value from row
     * @return Column builder
     */
    public ExcelColumn.ExcelColumnBuilder<T> column(String name, Function<T, Object> function) {
        return new ExcelColumn.ExcelColumnBuilder<>(this, name, (r, c) -> function.apply(r));
    }

    /**
     * Begins building a new column with constant value for all rows.
     *
     * @param name  Column header name
     * @param value Constant value to be used in all rows
     * @return Column builder
     */
    public ExcelColumn.ExcelColumnBuilder<T> constColumn(String name, Object value) {
        return new ExcelColumn.ExcelColumnBuilder<>(this, name, (r, c) -> value);
    }

    /**
     * Writes the stream of row data into an Excel file using custom row-level callback.
     *
     * @param stream   The data stream
     * @param consumer Custom consumer for post-processing row with cursor
     * @return ExcelHandler wrapping the workbook
     */
    ExcelHandler write(Stream<T> stream, ExcelConsumer<T> consumer) {
        if (this.columns.isEmpty()) {
            throw new ExcelWriteException("columns setting required");
        }

        this.sheet = wb.createSheet();
        this.cursor = new Cursor(this.title != null ? 2 : 0);

        if (this.title != null) {
            setSheetTitle();
        }

        setColumnHeaders();

        try (stream) {
            stream.forEach(rowData -> {
                this.handleRowData(rowData);
                consumer.accept(rowData, cursor);
            });
        }
        applyColumWidthAllSheets();
        return new ExcelHandler(this.wb);
    }

    /**
     * Writes the stream of row data into Excel file without row-level callback.
     *
     * @param stream The data stream
     * @return ExcelHandler wrapping the workbook
     */
    ExcelHandler write(Stream<T> stream) {
        return this.write(stream, (rowData, consumer) -> {});
    }

    /**
     * Internal method to set the sheet title and merge cells across columns.
     */
    private void setSheetTitle() {
        CellRangeAddress region = new CellRangeAddress(0, 1, 0, this.columns.size() - 1);

        sheet.addMergedRegion(region);

        SXSSFRow titleRow = sheet.createRow(0);
        SXSSFCell cell = titleRow.createCell(0);
        cell.setCellValue(title);
        cell.setCellStyle(titleStyle);
    }

    /**
     * Writes column headers to the current sheet using the predefined header style.
     */
    private void setColumnHeaders() {
        SXSSFRow headRow = sheet.createRow(cursor.getRowOfSheet());
        cursor.plusRow();
        for (int j = 0; j < this.columns.size(); j++) {
            SXSSFCell cell = headRow.createCell(j);
            ExcelColumn<T> column = columns.get(j);
            cell.setCellValue(column.getName());
            cell.setCellStyle(headerStyle);
        }
    }

    /**
     * Handles the logic of writing a single row to the sheet, including value mapping and style.
     *
     * @param rowData A row of data
     */
    void handleRowData(T rowData) {
        cursor.plusTotal();
        if (isOverMaxRows()) {
            turnOverSheet();
            if (this.title != null) {
                setSheetTitle();
            }
            setColumnHeaders();
        }
        SXSSFRow row = sheet.createRow(cursor.getRowOfSheet());
        row.setHeightInPoints(20);
        cursor.plusRow();

        for (int j = 0; j < this.columns.size(); j++) {
            SXSSFCell cell = row.createCell(j);
            ExcelColumn<T> column = columns.get(j);
            Object columnData = column.applyFunction(rowData, cursor);
            column.setColumnData(cell, columnData);
            cell.setCellStyle(column.getStyle());
            if (cursor.getRowOfSheet() < 100) {
                column.fitColumnWidthByValue(columnData);
            }
        }
    }

    /**
     * Creates a new sheet and resets the row index when the current sheet exceeds row limit.
     */
    private void turnOverSheet() {
        this.sheet = wb.createSheet();
        this.cursor.initRow();
    }

    /**
     * Checks whether the current sheet has exceeded its max row limit.
     *
     * @return true if a sheet needs to turn over; otherwise false
     */
    private boolean isOverMaxRows() {
        return cursor.getCurrentTotal() >= maxRowsOfSheet && cursor.getCurrentTotal() % maxRowsOfSheet == 1;
    }

    /**
     * Applies the calculated column widths to all sheets after writing is complete.
     */
    private void applyColumWidthAllSheets() {
        int numberOfSheets = wb.getNumberOfSheets();
        for (int i = 0; i < numberOfSheets; i++) {
            SXSSFSheet s = wb.getSheetAt(i);
            for (int j = 0; j < columns.size(); j++) {
                s.setColumnWidth(j, columns.get(j).getColumnWidth());
            }
        }
    }

    /**
     * Returns the underlying streaming workbook instance.
     *
     * @return SXSSFWorkbook instance
     */
    SXSSFWorkbook getWb() {
        return wb;
    }

    Map<String, CellStyle> getCellStyleCache() {
        return cellStyleCache;
    }

    /**
     * Closes the underlying workbook, releasing any resources.
     * <p>
     * This is a safety net for cases where {@link #write(Stream)} is never called.
     * If the workbook has already been consumed via {@link ExcelHandler}, this is a no-op.
     */
    @Override
    public void close() {
        try {
            wb.close();
        } catch (Exception e) {
            // already closed or disposed — safe to ignore
        }
    }
}
