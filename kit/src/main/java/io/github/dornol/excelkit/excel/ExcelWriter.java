package io.github.dornol.excelkit.excel;

import io.github.dornol.excelkit.shared.Cursor;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DataValidation;
import org.apache.poi.ss.usermodel.DataValidationConstraint;
import org.apache.poi.ss.usermodel.DataValidationHelper;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.jspecify.annotations.NonNull;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellRangeAddressList;
import org.apache.poi.xssf.streaming.SXSSFCell;
import org.apache.poi.xssf.streaming.SXSSFRow;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

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
    private static final Logger log = LoggerFactory.getLogger(ExcelWriter.class);
    private static final int AUTO_WIDTH_SAMPLE_ROWS = 100;
    private static final int DEFAULT_ROW_ACCESS_WINDOW_SIZE = 1000;
    private static final int EXCEL_MAX_ROWS = 1_048_575;

    private final SXSSFWorkbook wb;
    private final List<ExcelColumn<T>> columns = new ArrayList<>();
    private final int maxRowsOfSheet;
    private final CellStyle headerStyle;
    private final Map<String, CellStyle> cellStyleCache = new HashMap<>();
    private String title;
    private CellStyle titleStyle;
    private float rowHeightInPoints = 20;
    private boolean autoFilter = false;
    private int freezePaneRows = 0;
    private BeforeHeaderWriter beforeHeaderWriter;
    private AfterDataWriter afterDataWriter;
    private AfterDataWriter afterAllWriter;
    private Function<Integer, String> sheetNameFunction;
    private int sheetCount = 0;
    private Function<T, ExcelColor> rowColorFunction;
    private final Map<String, CellStyle> rowStyleCache = new HashMap<>();
    private int headerRowIndex;

    private SXSSFSheet sheet;
    private Cursor cursor;


    /**
     * Constructs an ExcelWriter with a custom header color, maximum rows per sheet, and row access window size.
     *
     * @param r                    Red component of the header color (0–255)
     * @param g                    Green component of the header color (0–255)
     * @param b                    Blue component of the header color (0–255)
     * @param maxRowsOfSheet       Maximum number of rows allowed per sheet before creating a new one
     * @param rowAccessWindowSize  Number of rows kept in memory by SXSSFWorkbook (higher = more memory, lower = less memory)
     */
    public ExcelWriter(int r, int g, int b, int maxRowsOfSheet, int rowAccessWindowSize) {
        this.wb = new SXSSFWorkbook(rowAccessWindowSize);
        this.maxRowsOfSheet = maxRowsOfSheet;
        this.headerStyle = ExcelStyleSupporter.headerStyle(wb, new XSSFColor(new byte[]{(byte) r, (byte) g, (byte) b}));
    }

    /**
     * Constructs an ExcelWriter with a custom header color and maximum rows per sheet.
     *
     * @param r               Red component of the header color (0–255)
     * @param g               Green component of the header color (0–255)
     * @param b               Blue component of the header color (0–255)
     * @param maxRowsOfSheet  Maximum number of rows allowed per sheet before creating a new one
     */
    public ExcelWriter(int r, int g, int b, int maxRowsOfSheet) {
        this(r, g, b, maxRowsOfSheet, DEFAULT_ROW_ACCESS_WINDOW_SIZE);
    }

    /**
     * Constructs an ExcelWriter with a preset header color, maximum rows per sheet, and row access window size.
     *
     * @param color              Preset header color
     * @param maxRowsOfSheet     Maximum number of rows allowed per sheet before creating a new one
     * @param rowAccessWindowSize Number of rows kept in memory by SXSSFWorkbook
     */
    public ExcelWriter(ExcelColor color, int maxRowsOfSheet, int rowAccessWindowSize) {
        this(color.getR(), color.getG(), color.getB(), maxRowsOfSheet, rowAccessWindowSize);
    }

    /**
     * Constructs an ExcelWriter with a preset header color and maximum rows per sheet.
     *
     * @param color          Preset header color
     * @param maxRowsOfSheet Maximum number of rows allowed per sheet before creating a new one
     */
    public ExcelWriter(ExcelColor color, int maxRowsOfSheet) {
        this(color.getR(), color.getG(), color.getB(), maxRowsOfSheet);
    }

    /**
     * Constructs an ExcelWriter with a preset header color and default max 1,000,000 rows per sheet.
     *
     * @param color Preset header color
     */
    public ExcelWriter(ExcelColor color) {
        this(color.getR(), color.getG(), color.getB());
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
    public ExcelWriter<T> title(@NonNull String title) {
        return title(title, 0, IndexedColors.BLACK);
    }

    /**
     * Sets the title for the Excel sheet with a specified font size.
     *
     * @param title    The title text to display at the top
     * @param fontSize Font size in points
     * @return Current ExcelWriter instance for chaining
     */
    public ExcelWriter<T> title(@NonNull String title, int fontSize) {
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
    public ExcelWriter<T> title(@NonNull String title, int fontSize, @NonNull IndexedColors color) {
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
     * Sets the row height for data rows in points.
     * Defaults to 20 points.
     *
     * @param rowHeightInPoints Row height in points (must be positive)
     * @return Current ExcelWriter instance for chaining
     */
    public ExcelWriter<T> rowHeight(float rowHeightInPoints) {
        if (rowHeightInPoints <= 0) {
            throw new IllegalArgumentException("rowHeightInPoints must be positive");
        }
        this.rowHeightInPoints = rowHeightInPoints;
        return this;
    }

    /**
     * Enables or disables auto-filter on the header row.
     *
     * @param autoFilter Whether to apply auto-filter
     * @return Current ExcelWriter instance for chaining
     */
    public ExcelWriter<T> autoFilter(boolean autoFilter) {
        this.autoFilter = autoFilter;
        return this;
    }

    /**
     * Sets the number of rows to freeze below the header row.
     * The freeze pane is created starting from the header row position.
     *
     * @param rows Number of rows to freeze (must be non-negative)
     * @return Current ExcelWriter instance for chaining
     */
    public ExcelWriter<T> freezePane(int rows) {
        if (rows < 0) {
            throw new IllegalArgumentException("freezePaneRows must be non-negative");
        }
        this.freezePaneRows = rows;
        return this;
    }

    /**
     * Registers a callback that writes custom content before the column header row.
     * <p>
     * The callback receives the current sheet, the workbook, and the first available row index,
     * and must return the next available row index (where the header will start).
     * When a title is set, {@code startRow} will be 2 (after the title rows); otherwise 0.
     * <p>
     * The callback is invoked on every sheet, including rollover sheets,
     * so it must always produce the same number of rows.
     *
     * @param beforeHeaderWriter the callback to invoke before writing column headers
     * @return Current ExcelWriter instance for chaining
     */
    public ExcelWriter<T> beforeHeader(BeforeHeaderWriter beforeHeaderWriter) {
        this.beforeHeaderWriter = beforeHeaderWriter;
        return this;
    }

    /**
     * Registers a callback that writes custom content after all data rows on each sheet.
     * <p>
     * Called on every sheet (including rollover sheets) after its data rows are written.
     * On the last sheet, this is called before {@code afterAll}.
     *
     * @param afterDataWriter the callback to invoke after data rows
     * @return Current ExcelWriter instance for chaining
     */
    public ExcelWriter<T> afterData(AfterDataWriter afterDataWriter) {
        this.afterDataWriter = afterDataWriter;
        return this;
    }

    /**
     * Registers a callback that writes custom content once on the last sheet after all data.
     * <p>
     * Called only once, on the last sheet, after {@code afterData} (if set).
     * Useful for writing grand totals or summary rows.
     *
     * @param afterAllWriter the callback to invoke after all data is written
     * @return Current ExcelWriter instance for chaining
     */
    public ExcelWriter<T> afterAll(AfterDataWriter afterAllWriter) {
        this.afterAllWriter = afterAllWriter;
        return this;
    }

    /**
     * Sets a function that generates sheet names based on the sheet index (0-based).
     *
     * @param sheetNameFunction a function that takes the sheet index and returns the sheet name
     * @return Current ExcelWriter instance for chaining
     */
    public ExcelWriter<T> sheetName(Function<Integer, String> sheetNameFunction) {
        this.sheetNameFunction = sheetNameFunction;
        return this;
    }

    /**
     * Sets a fixed sheet name. When sheets roll over, subsequent sheets are named
     * "{name} (2)", "{name} (3)", etc.
     *
     * @param name the base sheet name
     * @return Current ExcelWriter instance for chaining
     */
    public ExcelWriter<T> sheetName(String name) {
        this.sheetNameFunction = index -> index == 0 ? name : name + " (" + (index + 1) + ")";
        return this;
    }

    /**
     * Sets a function that determines the background color for each row.
     * <p>
     * When set, the function is called for each row of data. If it returns a non-null
     * {@link ExcelColor}, that color is applied as the background to all cells in the row,
     * overriding any column-level background color.
     *
     * @param rowColorFunction function that takes row data and returns an ExcelColor (or null for no override)
     * @return Current ExcelWriter instance for chaining
     */
    public ExcelWriter<T> rowColor(Function<T, ExcelColor> rowColorFunction) {
        this.rowColorFunction = rowColorFunction;
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
    public ExcelColumn.ExcelColumnBuilder<T> column(@NonNull String name, @NonNull ExcelRowFunction<T, Object> function) {
        return new ExcelColumn.ExcelColumnBuilder<>(this, name, function);
    }

    /**
     * Begins building a new column using a basic row-mapping function.
     *
     * @param name     Column header name
     * @param function Function to extract cell value from row
     * @return Column builder
     */
    public ExcelColumn.ExcelColumnBuilder<T> column(@NonNull String name, @NonNull Function<T, Object> function) {
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
     * Adds a column with default STRING type using a simple Function.
     * Useful for schema-based column registration.
     *
     * @param name     Column header name
     * @param function Function to extract cell value from row
     * @return Current ExcelWriter instance for chaining
     */
    public ExcelWriter<T> addColumn(String name, Function<T, Object> function) {
        ExcelColumn.ExcelColumnBuilder<T> builder =
                new ExcelColumn.ExcelColumnBuilder<>(this, name, (r, c) -> function.apply(r));
        this.columns.add(builder.build());
        return this;
    }

    /**
     * Writes the stream of row data into an Excel file using custom row-level callback.
     *
     * @param stream   The data stream
     * @param consumer Custom consumer for post-processing row with cursor
     * @return ExcelHandler wrapping the workbook
     */
    public ExcelHandler write(Stream<T> stream, ExcelConsumer<T> consumer) {
        if (this.columns.isEmpty()) {
            throw new ExcelWriteException("columns setting required");
        }

        this.sheet = createNamedSheet();
        int headerStartRow = initSheetPreamble();
        this.cursor = new Cursor(headerStartRow);
        this.headerRowIndex = headerStartRow;

        setColumnHeaders();
        applySheetOptions();

        try (stream) {
            stream.forEach(rowData -> {
                this.handleRowData(rowData);
                consumer.accept(rowData, cursor);
            });
        }

        int nextRow = cursor.getRowOfSheet();
        if (this.afterDataWriter != null) {
            nextRow = this.afterDataWriter.write(this.sheet, this.wb, nextRow);
        }
        if (this.afterAllWriter != null) {
            this.afterAllWriter.write(this.sheet, this.wb, nextRow);
        }

        applyDataValidations();
        applyColumnWidthAllSheets();
        return new ExcelHandler(this.wb);
    }

    /**
     * Writes the stream of row data into Excel file without row-level callback.
     *
     * @param stream The data stream
     * @return ExcelHandler wrapping the workbook
     */
    public ExcelHandler write(Stream<T> stream) {
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
     * Writes the title (if set) and invokes the beforeHeader callback (if set).
     *
     * @return the row index where the column header should be written
     */
    private int initSheetPreamble() {
        int currentRow = 0;
        if (this.title != null) {
            setSheetTitle();
            currentRow = 2;
        }
        if (this.beforeHeaderWriter != null) {
            currentRow = this.beforeHeaderWriter.write(this.sheet, this.wb, currentRow);
        }
        return currentRow;
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
     * Applies optional sheet-level settings such as auto-filter and freeze panes.
     */
    private void applySheetOptions() {
        int headerRowIdx = cursor.getRowOfSheet() - 1;
        if (this.autoFilter) {
            sheet.setAutoFilter(new CellRangeAddress(headerRowIdx, headerRowIdx, 0, columns.size() - 1));
        }
        if (this.freezePaneRows > 0) {
            sheet.createFreezePane(0, headerRowIdx + this.freezePaneRows);
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
            if (this.afterDataWriter != null) {
                this.afterDataWriter.write(this.sheet, this.wb, cursor.getRowOfSheet());
            }
            turnOverSheet();
            initSheetPreamble();
            setColumnHeaders();
            applySheetOptions();
        }
        SXSSFRow row = sheet.createRow(cursor.getRowOfSheet());
        row.setHeightInPoints(rowHeightInPoints);
        cursor.plusRow();

        ExcelColor rowColor = (rowColorFunction != null) ? rowColorFunction.apply(rowData) : null;

        for (int j = 0; j < this.columns.size(); j++) {
            SXSSFCell cell = row.createCell(j);
            ExcelColumn<T> column = columns.get(j);
            Object columnData = column.applyFunction(rowData, cursor);
            column.setColumnData(cell, columnData);
            if (rowColor != null) {
                cell.setCellStyle(getRowColorStyle(column.getStyle(), rowColor));
            } else {
                cell.setCellStyle(column.getStyle());
            }
            if (cursor.getRowOfSheet() < AUTO_WIDTH_SAMPLE_ROWS) {
                column.fitColumnWidthByValue(columnData);
            }
        }
    }

    /**
     * Creates a new sheet with a name determined by the sheet name function (if set).
     *
     * @return the newly created sheet
     */
    private SXSSFSheet createNamedSheet() {
        int index = sheetCount++;
        if (sheetNameFunction != null) {
            return wb.createSheet(sheetNameFunction.apply(index));
        }
        return wb.createSheet();
    }

    /**
     * Creates a new sheet and resets the row index when the current sheet exceeds row limit.
     */
    private void turnOverSheet() {
        this.sheet = createNamedSheet();
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
    private void applyColumnWidthAllSheets() {
        int numberOfSheets = wb.getNumberOfSheets();
        for (int i = 0; i < numberOfSheets; i++) {
            SXSSFSheet s = wb.getSheetAt(i);
            for (int j = 0; j < columns.size(); j++) {
                s.setColumnWidth(j, columns.get(j).getColumnWidth());
            }
        }
    }

    /**
     * Applies dropdown data validations to all sheets for columns that have dropdownOptions.
     */
    private void applyDataValidations() {
        for (int i = 0; i < wb.getNumberOfSheets(); i++) {
            SXSSFSheet s = wb.getSheetAt(i);
            DataValidationHelper helper = s.getDataValidationHelper();
            for (int j = 0; j < columns.size(); j++) {
                String[] options = columns.get(j).getDropdownOptions();
                if (options != null) {
                    DataValidationConstraint constraint = helper.createExplicitListConstraint(options);
                    CellRangeAddressList range = new CellRangeAddressList(
                            headerRowIndex + 1, EXCEL_MAX_ROWS, j, j);
                    DataValidation validation = helper.createValidation(constraint, range);
                    validation.setSuppressDropDownArrow(false);
                    validation.setShowErrorBox(true);
                    s.addValidationData(validation);
                }
            }
        }
    }

    /**
     * Returns a CellStyle that clones the base style but overrides the background color.
     * Results are cached by base style index + RGB to avoid creating excessive styles.
     */
    private CellStyle getRowColorStyle(CellStyle baseStyle, ExcelColor color) {
        String key = baseStyle.getIndex() + "_" + color.getR() + "_" + color.getG() + "_" + color.getB();
        return rowStyleCache.computeIfAbsent(key, k -> {
            CellStyle style = wb.createCellStyle();
            style.cloneStyleFrom(baseStyle);
            style.setFillForegroundColor(new XSSFColor(new byte[]{
                    (byte) color.getR(), (byte) color.getG(), (byte) color.getB()}));
            style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            return style;
        });
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
            log.debug("ExcelWriter.close() caught exception (likely already closed)", e);
        }
    }
}
