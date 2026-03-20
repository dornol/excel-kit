package io.github.dornol.excelkit.excel;

import io.github.dornol.excelkit.shared.Cursor;
import io.github.dornol.excelkit.shared.ProgressCallback;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import org.jspecify.annotations.Nullable;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.function.Consumer;
import java.util.function.Function;
import java.util.stream.Stream;

/**
 * ExcelWriter is a utility class for generating large Excel files using Apache POI's SXSSFWorkbook.
 * Supports streaming writes, column configuration, style customization, and sheet auto-splitting.
 *
 * @author dhkim
 * @param <T> The data type of each row to be written into the Excel file
 * @since 2025-07-19
 */
public class ExcelWriter<T> {
    private static final Logger log = LoggerFactory.getLogger(ExcelWriter.class);
    private static final int DEFAULT_ROW_ACCESS_WINDOW_SIZE = 1000;

    private final SXSSFWorkbook wb;
    private final List<ExcelColumn<T>> columns = new ArrayList<>();
    private final int maxRowsOfSheet;
    private CellStyle headerStyle;
    private final XSSFColor headerColor;
    private final Map<String, CellStyle> cellStyleCache = new HashMap<>();
    private final SheetConfig<T> cfg = new SheetConfig<>();
    private @Nullable AfterDataWriter afterAllWriter;
    private int sheetCount = 0;
    private final Map<String, CellStyle> rowStyleCache = new HashMap<>();
    private int headerRowIndex;
    private @Nullable String workbookPassword;
    private @Nullable String headerFontName;
    private @Nullable Integer headerFontSize;

    private @Nullable SXSSFSheet sheet;
    private @Nullable Cursor cursor;


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
        this.headerColor = new XSSFColor(new byte[]{(byte) r, (byte) g, (byte) b});
        this.headerStyle = ExcelStyleSupporter.headerStyle(wb, headerColor);
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
        this.cfg.rowHeightInPoints = rowHeightInPoints;
        return this;
    }

    /**
     * Enables or disables auto-filter on the header row.
     *
     * @param autoFilter Whether to apply auto-filter
     * @return Current ExcelWriter instance for chaining
     */
    public ExcelWriter<T> autoFilter(boolean autoFilter) {
        this.cfg.autoFilter = autoFilter;
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
        this.cfg.freezePaneRows = rows;
        return this;
    }

    /**
     * Registers a callback that writes custom content before the column header row.
     * <p>
     * The callback is invoked on every sheet, including rollover sheets,
     * so it must always produce the same number of rows.
     *
     * @param beforeHeaderWriter the callback to invoke before writing column headers
     * @return Current ExcelWriter instance for chaining
     */
    public ExcelWriter<T> beforeHeader(BeforeHeaderWriter beforeHeaderWriter) {
        this.cfg.beforeHeaderWriter = beforeHeaderWriter;
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
        this.cfg.afterDataWriter = afterDataWriter;
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
        this.cfg.sheetNameFunction = sheetNameFunction;
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
        this.cfg.sheetNameFunction = index -> index == 0 ? name : name + " (" + (index + 1) + ")";
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
    public ExcelWriter<T> rowColor(Function<T, @Nullable ExcelColor> rowColorFunction) {
        this.cfg.rowColorFunction = rowColorFunction;
        return this;
    }

    /**
     * Registers a progress callback that fires every {@code interval} rows.
     *
     * @param interval the number of rows between each callback invocation (must be positive)
     * @param callback the callback to invoke
     * @return Current ExcelWriter instance for chaining
     */
    public ExcelWriter<T> onProgress(int interval, ProgressCallback callback) {
        if (interval <= 0) {
            throw new IllegalArgumentException("progress interval must be positive");
        }
        this.cfg.progressInterval = interval;
        this.cfg.progressCallback = callback;
        return this;
    }

    /**
     * Sets the number of rows sampled for auto column width calculation.
     * <p>
     * Only the first N data rows are measured to determine column widths.
     * Set to 0 to disable auto-width (all columns use minimum width).
     * Defaults to 100.
     *
     * @param rows number of rows to sample (0 to disable)
     * @return Current ExcelWriter instance for chaining
     */
    public ExcelWriter<T> autoWidthSampleRows(int rows) {
        if (rows < 0) {
            throw new IllegalArgumentException("autoWidthSampleRows must be non-negative");
        }
        this.cfg.autoWidthSampleRows = rows;
        return this;
    }

    /**
     * Protects all sheets with the given password.
     * <p>
     * When sheet protection is enabled, cells are locked by default.
     * Use {@link ExcelColumn.ExcelColumnBuilder#locked(boolean)} to allow editing specific columns.
     *
     * @param password the protection password
     * @return Current ExcelWriter instance for chaining
     */
    public ExcelWriter<T> protectSheet(String password) {
        this.cfg.sheetPassword = password;
        return this;
    }

    /**
     * Adds a conditional formatting rule.
     *
     * @param configurer consumer to configure the rule
     * @return Current ExcelWriter instance for chaining
     */
    public ExcelWriter<T> conditionalFormatting(Consumer<ExcelConditionalRule> configurer) {
        cfg.addConditionalRule(configurer);
        return this;
    }

    /**
     * Configures a chart to be added after all data is written.
     *
     * @param configurer consumer to configure the chart
     * @return Current ExcelWriter instance for chaining
     */
    public ExcelWriter<T> chart(Consumer<ExcelChartConfig> configurer) {
        ExcelChartConfig config = new ExcelChartConfig();
        configurer.accept(config);
        this.cfg.chartConfig = config;
        return this;
    }

    /**
     * Configures print setup (page layout) for all sheets.
     * <p>
     * Controls orientation, paper size, margins, headers/footers, repeat rows, and fit-to-page.
     *
     * @param configurer consumer to configure the print setup
     * @return Current ExcelWriter instance for chaining
     */
    public ExcelWriter<T> printSetup(Consumer<ExcelPrintSetup> configurer) {
        ExcelPrintSetup config = new ExcelPrintSetup();
        configurer.accept(config);
        this.cfg.printSetup = config;
        return this;
    }

    /**
     * Sets the sheet tab color using RGB values.
     *
     * @param r Red component (0–255)
     * @param g Green component (0–255)
     * @param b Blue component (0–255)
     * @return Current ExcelWriter instance for chaining
     * @since 0.7.0
     */
    public ExcelWriter<T> tabColor(int r, int g, int b) {
        this.cfg.tabColor = new int[]{r, g, b};
        return this;
    }

    /**
     * Sets the sheet tab color using a preset color.
     *
     * @param color Preset color
     * @return Current ExcelWriter instance for chaining
     * @since 0.7.0
     */
    public ExcelWriter<T> tabColor(ExcelColor color) {
        return tabColor(color.getR(), color.getG(), color.getB());
    }

    /**
     * Protects the workbook structure with the given password.
     * <p>
     * When enabled, users cannot add, delete, rename, or reorder sheets.
     *
     * @param password the protection password
     * @return Current ExcelWriter instance for chaining
     */
    public ExcelWriter<T> protectWorkbook(String password) {
        this.workbookPassword = password;
        return this;
    }

    /**
     * Sets the header font name.
     *
     * @param fontName the font name (e.g., "Arial", "맑은 고딕")
     * @return Current ExcelWriter instance for chaining
     */
    public ExcelWriter<T> headerFontName(String fontName) {
        this.headerFontName = fontName;
        return this;
    }

    /**
     * Sets the header font size in points.
     *
     * @param fontSize font size in points (must be positive)
     * @return Current ExcelWriter instance for chaining
     */
    public ExcelWriter<T> headerFontSize(int fontSize) {
        if (fontSize <= 0) {
            throw new IllegalArgumentException("fontSize must be positive");
        }
        this.headerFontSize = fontSize;
        return this;
    }

    /**
     * Sets default column styles that apply to all columns unless overridden per-column.
     *
     * @param configurer consumer to configure default style properties
     * @return Current ExcelWriter instance for chaining
     */
    public ExcelWriter<T> defaultStyle(Consumer<ColumnStyleConfig.DefaultStyleConfig<T>> configurer) {
        ColumnStyleConfig.DefaultStyleConfig<T> config = new ColumnStyleConfig.DefaultStyleConfig<>();
        configurer.accept(config);
        this.cfg.defaultStyleConfig = config;
        return this;
    }

    /**
     * Configures summary (footer) rows with formulas such as SUM, AVERAGE, COUNT, MIN, MAX.
     * <p>
     * Summary rows are appended after data rows on each sheet.
     *
     * @param configurer consumer to configure the summary
     * @return Current ExcelWriter instance for chaining
     */
    public ExcelWriter<T> summary(Consumer<ExcelSummary> configurer) {
        ExcelSummary summary = new ExcelSummary();
        configurer.accept(summary);
        this.cfg.summaryConfig = summary;
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
    public ExcelColumn.ExcelColumnBuilder<T> column(String name, ExcelRowFunction<T, @Nullable Object> function) {
        return new ExcelColumn.ExcelColumnBuilder<>(this, name, function);
    }

    /**
     * Begins building a new column using a basic row-mapping function.
     *
     * @param name     Column header name
     * @param function Function to extract cell value from row
     * @return Column builder
     */
    public ExcelColumn.ExcelColumnBuilder<T> column(String name, Function<T, @Nullable Object> function) {
        return new ExcelColumn.ExcelColumnBuilder<>(this, name, (r, c) -> function.apply(r));
    }

    /**
     * Begins building a new column with constant value for all rows.
     *
     * @param name  Column header name
     * @param value Constant value to be used in all rows
     * @return Column builder
     */
    public ExcelColumn.ExcelColumnBuilder<T> constColumn(String name, @Nullable Object value) {
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
    public ExcelWriter<T> addColumn(String name, Function<T, @Nullable Object> function) {
        ExcelColumn.ExcelColumnBuilder<T> builder =
                new ExcelColumn.ExcelColumnBuilder<>(this, name, (r, c) -> function.apply(r));
        this.columns.add(builder.build());
        return this;
    }

    /**
     * Adds a column with additional configuration using a configurer consumer.
     * <p>
     * The configurer receives an {@link ExcelColumn.ExcelColumnBuilder} to set
     * column properties such as type, format, alignment, width, etc.
     *
     * <pre>{@code
     * writer.addColumn("Price", Book::getPrice, c -> c.type(ExcelDataType.INTEGER).format("#,##0"));
     * }</pre>
     *
     * @param name        Column header name
     * @param function    Function to extract cell value from row
     * @param configurer  Consumer to configure column properties
     * @return Current ExcelWriter instance for chaining
     */
    public ExcelWriter<T> addColumn(String name, Function<T, @Nullable Object> function,
                                     Consumer<ExcelColumn.ExcelColumnBuilder<T>> configurer) {
        ExcelColumn.ExcelColumnBuilder<T> builder =
                new ExcelColumn.ExcelColumnBuilder<>(this, name, (r, c) -> function.apply(r));
        if (configurer != null) {
            configurer.accept(builder);
        }
        this.columns.add(builder.build());
        return this;
    }

    /**
     * Adds a column with cursor access using an ExcelRowFunction.
     * Useful when the column value depends on row position (e.g., row number).
     *
     * @param name     Column header name
     * @param function Function to extract cell value from row with cursor access
     * @return Current ExcelWriter instance for chaining
     */
    public ExcelWriter<T> addColumn(String name, ExcelRowFunction<T, @Nullable Object> function) {
        ExcelColumn.ExcelColumnBuilder<T> builder =
                new ExcelColumn.ExcelColumnBuilder<>(this, name, function);
        this.columns.add(builder.build());
        return this;
    }

    /**
     * Adds a column with cursor access and additional configuration.
     *
     * @param name        Column header name
     * @param function    Function to extract cell value from row with cursor access
     * @param configurer  Consumer to configure column properties
     * @return Current ExcelWriter instance for chaining
     */
    public ExcelWriter<T> addColumn(String name, ExcelRowFunction<T, @Nullable Object> function,
                                     Consumer<ExcelColumn.ExcelColumnBuilder<T>> configurer) {
        ExcelColumn.ExcelColumnBuilder<T> builder =
                new ExcelColumn.ExcelColumnBuilder<>(this, name, function);
        if (configurer != null) {
            configurer.accept(builder);
        }
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
        ExcelWriteSupport.validateUniqueColumnNames(columns);

        if (headerFontName != null || headerFontSize != null) {
            this.headerStyle = ExcelStyleSupporter.headerStyle(wb, headerColor, headerFontName, headerFontSize);
        }

        this.sheet = createNamedSheet();
        int headerStartRow = ExcelWriteSupport.initSheetPreamble(sheet, wb, columns, cfg.beforeHeaderWriter);
        this.cursor = new Cursor(headerStartRow);
        this.headerRowIndex = headerStartRow;

        ExcelWriteSupport.writeColumnHeaders(sheet, cursor, columns, headerStyle);
        applySheetOptions();

        try (stream) {
            stream.forEach(rowData -> {
                this.handleRowData(rowData);
                consumer.accept(rowData, cursor);
            });
        }

        int nextRow = cursor.getRowOfSheet();
        if (this.cfg.afterDataWriter != null) {
            nextRow = this.cfg.afterDataWriter.write(new SheetContext(sheet, wb, nextRow, columns, headerRowIndex));
        }
        if (this.cfg.summaryConfig != null) {
            nextRow = this.cfg.summaryConfig.toAfterDataWriter().write(new SheetContext(sheet, wb, nextRow, columns, headerRowIndex));
        }
        if (this.afterAllWriter != null) {
            this.afterAllWriter.write(new SheetContext(sheet, wb, nextRow, columns, headerRowIndex));
        }

        applyPostProcessingAllSheets();
        ExcelWriteSupport.applyWorkbookProtection(wb, workbookPassword);

        // Apply chart on last sheet
        if (cfg.chartConfig != null) {
            ExcelWriteSupport.applyChart(sheet, cfg.chartConfig, headerRowIndex, cursor.getRowOfSheet() - 1);
        }

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
     * Applies optional sheet-level settings such as auto-filter and freeze panes.
     */
    private void applySheetOptions() {
        int headerRowIdx = cursor.getRowOfSheet() - 1;
        ExcelWriteSupport.applySheetOptions(sheet, headerRowIdx, cfg.autoFilter, cfg.freezePaneRows, columns.size());
    }

    /**
     * Handles the logic of writing a single row to the sheet, including value mapping and style.
     *
     * @param rowData A row of data
     */
    void handleRowData(T rowData) {
        cursor.plusTotal();
        if (isOverMaxRows()) {
            int rolloverRow = cursor.getRowOfSheet();
            if (this.cfg.afterDataWriter != null) {
                rolloverRow = this.cfg.afterDataWriter.write(new SheetContext(sheet, wb, rolloverRow, columns, headerRowIndex));
            }
            if (this.cfg.summaryConfig != null) {
                this.cfg.summaryConfig.toAfterDataWriter().write(new SheetContext(sheet, wb, rolloverRow, columns, headerRowIndex));
            }
            turnOverSheet();
            ExcelWriteSupport.initSheetPreamble(sheet, wb, columns, cfg.beforeHeaderWriter);
            ExcelWriteSupport.writeColumnHeaders(sheet, cursor, columns, headerStyle);
            applySheetOptions();
        }
        ExcelWriteSupport.writeRowCells(sheet, cursor, rowData, columns, cfg.rowHeightInPoints,
                cfg.rowColorFunction, rowStyleCache, wb, cfg.autoWidthSampleRows);
        ExcelWriteSupport.checkProgress(cursor, cfg.progressInterval, cfg.progressCallback);
    }

    /**
     * Creates a new sheet with a name determined by the sheet name function (if set).
     *
     * @return the newly created sheet
     */
    private SXSSFSheet createNamedSheet() {
        int index = sheetCount++;
        if (cfg.sheetNameFunction != null) {
            return wb.createSheet(cfg.sheetNameFunction.apply(index));
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
     * Applies all post-processing steps (column widths, validations, outlines, hiding,
     * protection, conditional formatting, print setup, tab color) to every sheet.
     */
    private void applyPostProcessingAllSheets() {
        for (int i = 0; i < wb.getNumberOfSheets(); i++) {
            SXSSFSheet s = wb.getSheetAt(i);
            ExcelWriteSupport.applyColumnWidths(s, columns);
            ExcelWriteSupport.applyDataValidations(s, columns, headerRowIndex);
            ExcelWriteSupport.applyColumnOutline(s, columns);
            ExcelWriteSupport.applyColumnHidden(s, columns);
            ExcelWriteSupport.applySheetProtection(s, cfg.sheetPassword);
            ExcelWriteSupport.applyConditionalFormatting(s, cfg.conditionalRules, headerRowIndex, columns.size(), s.getLastRowNum());
            ExcelWriteSupport.applyPrintSetup(s, cfg.printSetup, headerRowIndex);
            ExcelWriteSupport.applyTabColor(s, cfg.tabColor);
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

    ColumnStyleConfig.@Nullable DefaultStyleConfig<T> getDefaultStyleConfig() {
        return cfg.defaultStyleConfig;
    }

}
