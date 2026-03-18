package io.github.dornol.excelkit.excel;

import io.github.dornol.excelkit.shared.Cursor;
import io.github.dornol.excelkit.shared.ProgressCallback;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

import org.jspecify.annotations.Nullable;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Set;
import java.util.function.Consumer;
import java.util.function.Function;
import java.util.stream.Stream;


/**
 * Writes data of a specific type to one or more sheets within an {@link ExcelWorkbook}.
 * <p>
 * Supports optional auto-rollover via {@link #maxRows(int)} — when set, the writer
 * automatically creates additional sheets when the row limit is reached.
 *
 * @param <T> the row data type for this sheet
 * @author dhkim
 */
public class ExcelSheetWriter<T> {

    // Shared resources from ExcelWorkbook
    private final SXSSFWorkbook wb;
    private SXSSFSheet sheet;
    private final String baseName;
    private final CellStyle headerStyle;
    private final Map<String, CellStyle> cellStyleCache;
    private final Set<String> usedSheetNames;

    // Per-sheet settings
    private final List<ExcelColumn<T>> columns = new ArrayList<>();
    private final SheetConfig<T> cfg = new SheetConfig<>();
    private final Map<String, CellStyle> rowStyleCache = new HashMap<>();
    private int maxRows = Integer.MAX_VALUE;

    ExcelSheetWriter(SXSSFWorkbook wb, SXSSFSheet sheet, String baseName,
                     CellStyle headerStyle, Map<String, CellStyle> cellStyleCache,
                     Set<String> usedSheetNames) {
        this.wb = wb;
        this.sheet = sheet;
        this.baseName = baseName;
        this.headerStyle = headerStyle;
        this.cellStyleCache = cellStyleCache;
        this.usedSheetNames = usedSheetNames;
    }

    /**
     * Adds a column using a simple function.
     */
    public ExcelSheetWriter<T> column(String name, Function<T, @Nullable Object> function) {
        columns.add(buildColumn(name, (r, c) -> function.apply(r), null));
        return this;
    }

    /**
     * Adds a column with additional configuration.
     */
    public ExcelSheetWriter<T> column(String name, Function<T, @Nullable Object> function, Consumer<ColumnConfig<T>> cfg) {
        ColumnConfig<T> config = new ColumnConfig<>();
        cfg.accept(config);
        columns.add(buildColumn(name, (r, c) -> function.apply(r), config));
        return this;
    }

    /**
     * Adds a column using a row function with cursor support.
     */
    public ExcelSheetWriter<T> column(String name, ExcelRowFunction<T, @Nullable Object> function) {
        columns.add(buildColumn(name, function, null));
        return this;
    }

    /**
     * Adds a column using a row function with cursor support and additional configuration.
     */
    public ExcelSheetWriter<T> column(String name, ExcelRowFunction<T, @Nullable Object> function, Consumer<ColumnConfig<T>> cfg) {
        ColumnConfig<T> config = new ColumnConfig<>();
        cfg.accept(config);
        columns.add(buildColumn(name, function, config));
        return this;
    }

    /**
     * Adds a column with a constant value for all rows.
     */
    public ExcelSheetWriter<T> constColumn(String name, @Nullable Object value) {
        columns.add(buildColumn(name, (r, c) -> value, null));
        return this;
    }

    /**
     * Sets the row height for data rows in points.
     *
     * @param rowHeightInPoints Row height in points
     * @return this writer for chaining
     */
    public ExcelSheetWriter<T> rowHeight(float rowHeightInPoints) {
        this.cfg.rowHeightInPoints = rowHeightInPoints;
        return this;
    }

    /**
     * Enables auto-filter on the header row.
     *
     * @return this writer for chaining
     */
    public ExcelSheetWriter<T> autoFilter() {
        this.cfg.autoFilter = true;
        return this;
    }

    /**
     * Enables or disables auto-filter on the header row.
     *
     * @param autoFilter Whether to apply auto-filter
     * @return this writer for chaining
     */
    public ExcelSheetWriter<T> autoFilter(boolean autoFilter) {
        this.cfg.autoFilter = autoFilter;
        return this;
    }

    /**
     * Conditionally adds a column using a simple function.
     *
     * @param name      Column header name
     * @param condition Whether to include this column
     * @param function  Function to extract cell value from row
     * @return this writer for chaining
     */
    public ExcelSheetWriter<T> columnIf(String name, boolean condition, Function<T, @Nullable Object> function) {
        if (condition) {
            column(name, function);
        }
        return this;
    }

    /**
     * Conditionally adds a column with additional configuration.
     *
     * @param name      Column header name
     * @param condition Whether to include this column
     * @param function  Function to extract cell value from row
     * @param cfg       Consumer to configure column options
     * @return this writer for chaining
     */
    public ExcelSheetWriter<T> columnIf(String name, boolean condition, Function<T, @Nullable Object> function, Consumer<ColumnConfig<T>> cfg) {
        if (condition) {
            column(name, function, cfg);
        }
        return this;
    }

    public ExcelSheetWriter<T> freezePane(int rows) {
        this.cfg.freezePaneRows = rows;
        return this;
    }

    public ExcelSheetWriter<T> beforeHeader(BeforeHeaderWriter writer) {
        this.cfg.beforeHeaderWriter = writer;
        return this;
    }

    public ExcelSheetWriter<T> afterData(AfterDataWriter writer) {
        this.cfg.afterDataWriter = writer;
        return this;
    }

    public ExcelSheetWriter<T> rowColor(Function<T, @Nullable ExcelColor> fn) {
        this.cfg.rowColorFunction = fn;
        return this;
    }

    /**
     * Registers a progress callback that fires every {@code interval} rows.
     *
     * @param interval the number of rows between each callback invocation
     * @param callback the callback to invoke
     * @return this writer for chaining
     */
    public ExcelSheetWriter<T> onProgress(int interval, ProgressCallback callback) {
        if (interval <= 0) {
            throw new IllegalArgumentException("progress interval must be positive");
        }
        this.cfg.progressInterval = interval;
        this.cfg.progressCallback = callback;
        return this;
    }

    /**
     * Sets the maximum number of rows per sheet before auto-rollover.
     * <p>
     * When set, the writer automatically creates additional sheets within the workbook
     * when the row limit is reached. Use {@link #sheetName(Function)} to control
     * rollover sheet naming.
     *
     * @param maxRows maximum rows per sheet (must be positive)
     * @return this writer for chaining
     */
    public ExcelSheetWriter<T> maxRows(int maxRows) {
        if (maxRows <= 0) {
            throw new IllegalArgumentException("maxRows must be positive");
        }
        this.maxRows = maxRows;
        return this;
    }

    /**
     * Sets a function that generates sheet names for rollover sheets.
     * The function receives the 0-based rollover index (0 = first rollover sheet, i.e., the second sheet).
     * <p>
     * If not set, rollover sheets are named "{baseName} (2)", "{baseName} (3)", etc.
     *
     * @param sheetNameFunction function to generate sheet names
     * @return this writer for chaining
     */
    public ExcelSheetWriter<T> sheetName(Function<Integer, String> sheetNameFunction) {
        this.cfg.sheetNameFunction = sheetNameFunction;
        return this;
    }

    /**
     * Sets the number of rows sampled for auto column width calculation.
     * Defaults to 100. Set to 0 to disable.
     *
     * @param rows number of rows to sample
     * @return this writer for chaining
     */
    public ExcelSheetWriter<T> autoWidthSampleRows(int rows) {
        if (rows < 0) {
            throw new IllegalArgumentException("autoWidthSampleRows must be non-negative");
        }
        this.cfg.autoWidthSampleRows = rows;
        return this;
    }

    /**
     * Protects the sheet(s) with the given password.
     *
     * @param password the protection password
     * @return this writer for chaining
     */
    public ExcelSheetWriter<T> protectSheet(String password) {
        this.cfg.sheetPassword = password;
        return this;
    }

    /**
     * Adds a conditional formatting rule to the sheet(s).
     *
     * @param configurer consumer to configure the rule
     * @return this writer for chaining
     */
    public ExcelSheetWriter<T> conditionalFormatting(java.util.function.Consumer<ExcelConditionalRule> configurer) {
        cfg.addConditionalRule(configurer);
        return this;
    }

    /**
     * Configures a chart to be added after data is written.
     *
     * @param configurer consumer to configure the chart
     * @return this writer for chaining
     */
    public ExcelSheetWriter<T> chart(java.util.function.Consumer<ExcelChartConfig> configurer) {
        ExcelChartConfig config = new ExcelChartConfig();
        configurer.accept(config);
        this.cfg.chartConfig = config;
        return this;
    }

    /**
     * Configures print setup (page layout) for the sheet(s).
     * <p>
     * Controls orientation, paper size, margins, headers/footers, repeat rows, and fit-to-page.
     *
     * @param configurer consumer to configure the print setup
     * @return this writer for chaining
     */
    public ExcelSheetWriter<T> printSetup(Consumer<ExcelPrintSetup> configurer) {
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
     * @return this writer for chaining
     * @since 0.7.0
     */
    public ExcelSheetWriter<T> tabColor(int r, int g, int b) {
        this.cfg.tabColor = new int[]{r, g, b};
        return this;
    }

    /**
     * Sets the sheet tab color using a preset color.
     *
     * @param color Preset color
     * @return this writer for chaining
     * @since 0.7.0
     */
    public ExcelSheetWriter<T> tabColor(ExcelColor color) {
        return tabColor(color.getR(), color.getG(), color.getB());
    }

    /**
     * Sets default column styles that apply to all columns unless overridden per-column.
     *
     * @param configurer consumer to configure default style properties
     * @return this writer for chaining
     */
    public ExcelSheetWriter<T> defaultStyle(Consumer<ColumnStyleConfig.DefaultStyleConfig<T>> configurer) {
        ColumnStyleConfig.DefaultStyleConfig<T> config = new ColumnStyleConfig.DefaultStyleConfig<>();
        configurer.accept(config);
        this.cfg.defaultStyleConfig = config;
        return this;
    }

    /**
     * Configures summary (footer) rows with formulas such as SUM, AVERAGE, COUNT, MIN, MAX.
     *
     * @param configurer consumer to configure the summary
     * @return this writer for chaining
     */
    public ExcelSheetWriter<T> summary(Consumer<ExcelSummary> configurer) {
        ExcelSummary summary = new ExcelSummary();
        configurer.accept(summary);
        this.cfg.summaryConfig = summary;
        return this;
    }

    /**
     * Writes the data stream to this sheet (with optional auto-rollover).
     */
    public void write(Stream<T> stream) {
        if (columns.isEmpty()) {
            throw new ExcelWriteException("columns setting required");
        }
        ExcelWriteSupport.validateUniqueColumnNames(columns);

        List<SXSSFSheet> allSheets = new ArrayList<>();
        allSheets.add(this.sheet);

        int currentRow = ExcelWriteSupport.initSheetPreamble(sheet, wb, columns, cfg.beforeHeaderWriter);
        Cursor cursor = new Cursor(currentRow);
        int headerRowIndex = currentRow;

        ExcelWriteSupport.writeColumnHeaders(sheet, cursor, columns, headerStyle);
        int headerRowIdx = cursor.getRowOfSheet() - 1;
        ExcelWriteSupport.applySheetOptions(sheet, headerRowIdx, cfg.autoFilter, cfg.freezePaneRows, columns.size());

        // Mutable holder for current sheet in lambda
        SXSSFSheet[] currentSheet = {this.sheet};

        try (stream) {
            stream.forEach(rowData -> {
                cursor.plusTotal();
                if (maxRows != Integer.MAX_VALUE && cursor.getCurrentTotal() >= maxRows
                        && cursor.getCurrentTotal() % maxRows == 1) {
                    // afterData on current sheet
                    int rolloverRow = cursor.getRowOfSheet();
                    if (cfg.afterDataWriter != null) {
                        rolloverRow = cfg.afterDataWriter.write(new SheetContext(currentSheet[0], wb, rolloverRow, columns, headerRowIndex));
                    }
                    if (cfg.summaryConfig != null) {
                        cfg.summaryConfig.toAfterDataWriter().write(new SheetContext(currentSheet[0], wb, rolloverRow, columns, headerRowIndex));
                    }
                    // Create rollover sheet
                    currentSheet[0] = createRolloverSheet(allSheets.size());
                    allSheets.add(currentSheet[0]);
                    cursor.initRow();
                    ExcelWriteSupport.initSheetPreamble(currentSheet[0], wb, columns, cfg.beforeHeaderWriter);
                    ExcelWriteSupport.writeColumnHeaders(currentSheet[0], cursor, columns, headerStyle);
                    int hdrIdx = cursor.getRowOfSheet() - 1;
                    ExcelWriteSupport.applySheetOptions(currentSheet[0], hdrIdx, cfg.autoFilter, cfg.freezePaneRows, columns.size());
                }
                ExcelWriteSupport.writeRowCells(currentSheet[0], cursor, rowData, columns, cfg.rowHeightInPoints,
                        cfg.rowColorFunction, rowStyleCache, wb, cfg.autoWidthSampleRows);
                ExcelWriteSupport.checkProgress(cursor, cfg.progressInterval, cfg.progressCallback);
            });
        }

        int nextRow = cursor.getRowOfSheet();
        if (this.cfg.afterDataWriter != null) {
            nextRow = this.cfg.afterDataWriter.write(new SheetContext(currentSheet[0], wb, nextRow, columns, headerRowIndex));
        }
        if (this.cfg.summaryConfig != null) {
            this.cfg.summaryConfig.toAfterDataWriter().write(new SheetContext(currentSheet[0], wb, nextRow, columns, headerRowIndex));
        }

        for (SXSSFSheet s : allSheets) {
            ExcelWriteSupport.applyColumnWidths(s, columns);
            ExcelWriteSupport.applyDataValidations(s, columns, headerRowIndex);
            ExcelWriteSupport.applyColumnOutline(s, columns);
            ExcelWriteSupport.applyColumnHidden(s, columns);
            ExcelWriteSupport.applySheetProtection(s, cfg.sheetPassword);
            ExcelWriteSupport.applyConditionalFormatting(s, cfg.conditionalRules, headerRowIndex, columns.size());
            ExcelWriteSupport.applyPrintSetup(s, cfg.printSetup, headerRowIndex);
            ExcelWriteSupport.applyTabColor(s, cfg.tabColor);
        }

        // Apply chart on last sheet
        if (cfg.chartConfig != null) {
            SXSSFSheet lastSheet = allSheets.get(allSheets.size() - 1);
            ExcelWriteSupport.applyChart(lastSheet, cfg.chartConfig, headerRowIndex, cursor.getRowOfSheet() - 1);
        }
    }

    private SXSSFSheet createRolloverSheet(int rolloverIndex) {
        String name;
        if (cfg.sheetNameFunction != null) {
            name = cfg.sheetNameFunction.apply(rolloverIndex);
        } else {
            name = baseName + " (" + (rolloverIndex + 1) + ")";
        }
        if (usedSheetNames != null && !usedSheetNames.add(name)) {
            throw new ExcelWriteException("Duplicate sheet name during rollover: " + name);
        }
        return wb.createSheet(name);
    }

    private ExcelColumn<T> buildColumn(String name, ExcelRowFunction<T, @Nullable Object> function, @Nullable ColumnConfig<T> config) {
        ColumnStyleConfig<T, ?> c = config != null ? config : new ColumnConfig<>();
        if (cfg.defaultStyleConfig != null) {
            c.applyDefaults(cfg.defaultStyleConfig);
        }
        ExcelDataType dataType = c.dataType != null ? c.dataType : ExcelDataType.STRING;
        String dataFormat = c.dataFormat != null ? c.dataFormat : dataType.getDefaultFormat();

        CellStyleParams params = new CellStyleParams(c.alignment, dataFormat,
                c.backgroundColor, c.bold, c.fontSize, c.borderStyle, c.locked,
                c.rotation, c.borderTop, c.borderBottom, c.borderLeft, c.borderRight,
                c.fontColor, c.strikethrough, c.underline,
                c.verticalAlignment, c.wrapText, c.fontName, c.indentation);
        CellStyle style = ExcelStyleSupporter.cellStyle(wb, params, cellStyleCache);

        return new ExcelColumn<>(name, function, style, dataType.getSetter(),
                c.minWidth, c.maxWidth, c.fixedWidth, c.dropdownOptions,
                c.cellColorFunction, c.groupName, c.outlineLevel,
                c.commentFunction, c.borderStyle, c.locked, c.hidden, c.validation);
    }

    /**
     * Configuration class for column options.
     *
     * @param <T> the row data type
     */
    public static class ColumnConfig<T> extends ColumnStyleConfig<T, ColumnConfig<T>> {
    }
}
