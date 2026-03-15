package io.github.dornol.excelkit.excel;

import io.github.dornol.excelkit.shared.Cursor;
import io.github.dornol.excelkit.shared.ProgressCallback;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

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
    private float rowHeightInPoints = 20;
    private boolean autoFilter = false;
    private int freezePaneRows = 0;
    private BeforeHeaderWriter beforeHeaderWriter;
    private AfterDataWriter afterDataWriter;
    private Function<T, ExcelColor> rowColorFunction;
    private final Map<String, CellStyle> rowStyleCache = new HashMap<>();
    private ProgressCallback progressCallback;
    private int progressInterval;
    private int maxRows = Integer.MAX_VALUE;
    private Function<Integer, String> sheetNameFunction;
    private int autoWidthSampleRows = ExcelWriteSupport.AUTO_WIDTH_SAMPLE_ROWS;

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
    public ExcelSheetWriter<T> column(String name, Function<T, Object> function) {
        columns.add(buildColumn(name, (r, c) -> function.apply(r), null));
        return this;
    }

    /**
     * Adds a column with additional configuration.
     */
    public ExcelSheetWriter<T> column(String name, Function<T, Object> function, Consumer<ColumnConfig<T>> cfg) {
        ColumnConfig<T> config = new ColumnConfig<>();
        cfg.accept(config);
        columns.add(buildColumn(name, (r, c) -> function.apply(r), config));
        return this;
    }

    /**
     * Adds a column using a row function with cursor support.
     */
    public ExcelSheetWriter<T> column(String name, ExcelRowFunction<T, Object> function) {
        columns.add(buildColumn(name, function, null));
        return this;
    }

    /**
     * Adds a column using a row function with cursor support and additional configuration.
     */
    public ExcelSheetWriter<T> column(String name, ExcelRowFunction<T, Object> function, Consumer<ColumnConfig<T>> cfg) {
        ColumnConfig<T> config = new ColumnConfig<>();
        cfg.accept(config);
        columns.add(buildColumn(name, function, config));
        return this;
    }

    /**
     * Adds a column with a constant value for all rows.
     */
    public ExcelSheetWriter<T> constColumn(String name, Object value) {
        columns.add(buildColumn(name, (r, c) -> value, null));
        return this;
    }

    public ExcelSheetWriter<T> rowHeight(float rowHeightInPoints) {
        this.rowHeightInPoints = rowHeightInPoints;
        return this;
    }

    public ExcelSheetWriter<T> autoFilter() {
        this.autoFilter = true;
        return this;
    }

    /**
     * Enables or disables auto-filter on the header row.
     *
     * @param autoFilter Whether to apply auto-filter
     * @return this writer for chaining
     */
    public ExcelSheetWriter<T> autoFilter(boolean autoFilter) {
        this.autoFilter = autoFilter;
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
    public ExcelSheetWriter<T> columnIf(String name, boolean condition, Function<T, Object> function) {
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
    public ExcelSheetWriter<T> columnIf(String name, boolean condition, Function<T, Object> function, Consumer<ColumnConfig<T>> cfg) {
        if (condition) {
            column(name, function, cfg);
        }
        return this;
    }

    public ExcelSheetWriter<T> freezePane(int rows) {
        this.freezePaneRows = rows;
        return this;
    }

    public ExcelSheetWriter<T> beforeHeader(BeforeHeaderWriter writer) {
        this.beforeHeaderWriter = writer;
        return this;
    }

    public ExcelSheetWriter<T> afterData(AfterDataWriter writer) {
        this.afterDataWriter = writer;
        return this;
    }

    public ExcelSheetWriter<T> rowColor(Function<T, ExcelColor> fn) {
        this.rowColorFunction = fn;
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
        this.progressInterval = interval;
        this.progressCallback = callback;
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
        this.sheetNameFunction = sheetNameFunction;
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
        this.autoWidthSampleRows = rows;
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

        int currentRow = ExcelWriteSupport.initSheetPreamble(sheet, wb, columns, beforeHeaderWriter);
        Cursor cursor = new Cursor(currentRow);
        int headerRowIndex = currentRow;

        ExcelWriteSupport.writeColumnHeaders(sheet, cursor, columns, headerStyle);
        int headerRowIdx = cursor.getRowOfSheet() - 1;
        ExcelWriteSupport.applySheetOptions(sheet, headerRowIdx, autoFilter, freezePaneRows, columns.size());

        // Mutable holder for current sheet in lambda
        SXSSFSheet[] currentSheet = {this.sheet};

        try (stream) {
            stream.forEach(rowData -> {
                cursor.plusTotal();
                if (maxRows != Integer.MAX_VALUE && cursor.getCurrentTotal() >= maxRows
                        && cursor.getCurrentTotal() % maxRows == 1) {
                    // afterData on current sheet
                    if (afterDataWriter != null) {
                        afterDataWriter.write(new SheetContext(currentSheet[0], wb, cursor.getRowOfSheet(), columns));
                    }
                    // Create rollover sheet
                    currentSheet[0] = createRolloverSheet(allSheets.size());
                    allSheets.add(currentSheet[0]);
                    cursor.initRow();
                    ExcelWriteSupport.initSheetPreamble(currentSheet[0], wb, columns, beforeHeaderWriter);
                    ExcelWriteSupport.writeColumnHeaders(currentSheet[0], cursor, columns, headerStyle);
                    int hdrIdx = cursor.getRowOfSheet() - 1;
                    ExcelWriteSupport.applySheetOptions(currentSheet[0], hdrIdx, autoFilter, freezePaneRows, columns.size());
                }
                ExcelWriteSupport.writeRowCells(currentSheet[0], cursor, rowData, columns, rowHeightInPoints,
                        rowColorFunction, rowStyleCache, wb, autoWidthSampleRows);
                ExcelWriteSupport.checkProgress(cursor, progressInterval, progressCallback);
            });
        }

        int nextRow = cursor.getRowOfSheet();
        if (this.afterDataWriter != null) {
            this.afterDataWriter.write(new SheetContext(currentSheet[0], wb, nextRow, columns));
        }

        for (SXSSFSheet s : allSheets) {
            ExcelWriteSupport.applyColumnWidths(s, columns);
            ExcelWriteSupport.applyDataValidations(s, columns, headerRowIndex);
            ExcelWriteSupport.applyColumnOutline(s, columns);
        }
    }

    private SXSSFSheet createRolloverSheet(int rolloverIndex) {
        String name;
        if (sheetNameFunction != null) {
            name = sheetNameFunction.apply(rolloverIndex);
        } else {
            name = baseName + " (" + (rolloverIndex + 1) + ")";
        }
        if (usedSheetNames != null && !usedSheetNames.add(name)) {
            throw new ExcelWriteException("Duplicate sheet name during rollover: " + name);
        }
        return wb.createSheet(name);
    }

    private ExcelColumn<T> buildColumn(String name, ExcelRowFunction<T, Object> function, ColumnConfig<T> config) {
        ExcelDataType dataType = ExcelDataType.STRING;
        String dataFormat = null;
        HorizontalAlignment alignment = HorizontalAlignment.CENTER;
        int[] backgroundColor = null;
        Boolean bold = null;
        Integer fontSize = null;
        int minWidth = 0;
        int maxWidth = 0;
        boolean fixedWidth = false;
        String[] dropdownOptions = null;
        CellColorFunction<T> cellColorFunction = null;
        String groupName = null;
        int outlineLevel = 0;

        if (config != null) {
            if (config.dataType != null) dataType = config.dataType;
            dataFormat = config.dataFormat;
            if (config.alignment != null) alignment = config.alignment;
            backgroundColor = config.backgroundColor;
            bold = config.bold;
            fontSize = config.fontSize;
            minWidth = config.minWidth;
            maxWidth = config.maxWidth;
            fixedWidth = config.fixedWidth;
            dropdownOptions = config.dropdownOptions;
            cellColorFunction = config.cellColorFunction;
            groupName = config.groupName;
            outlineLevel = config.outlineLevel;
        }

        if (dataFormat == null) {
            dataFormat = dataType.getDefaultFormat();
        }

        CellStyle style = ExcelStyleSupporter.cellStyle(wb, alignment, dataFormat,
                backgroundColor, bold, fontSize, cellStyleCache);
        ExcelColumnSetter setter = dataType.getSetter();

        return new ExcelColumn<>(name, function, style, setter, minWidth, maxWidth, fixedWidth,
                dropdownOptions, cellColorFunction, groupName, outlineLevel);
    }

    /**
     * Configuration class for column options.
     *
     * @param <T> the row data type
     */
    public static class ColumnConfig<T> {
        private ExcelDataType dataType;
        private String dataFormat;
        private HorizontalAlignment alignment;
        private int[] backgroundColor;
        private Boolean bold;
        private Integer fontSize;
        private int minWidth;
        private int maxWidth;
        private boolean fixedWidth;
        private String[] dropdownOptions;
        private CellColorFunction<T> cellColorFunction;
        private String groupName;
        private int outlineLevel;

        public ColumnConfig<T> type(ExcelDataType dataType) {
            this.dataType = dataType;
            return this;
        }

        public ColumnConfig<T> format(String dataFormat) {
            this.dataFormat = dataFormat;
            return this;
        }

        public ColumnConfig<T> alignment(HorizontalAlignment alignment) {
            this.alignment = alignment;
            return this;
        }

        public ColumnConfig<T> backgroundColor(int r, int g, int b) {
            this.backgroundColor = new int[]{r, g, b};
            return this;
        }

        public ColumnConfig<T> backgroundColor(ExcelColor color) {
            return backgroundColor(color.getR(), color.getG(), color.getB());
        }

        public ColumnConfig<T> bold(boolean bold) {
            this.bold = bold;
            return this;
        }

        public ColumnConfig<T> fontSize(int fontSize) {
            if (fontSize <= 0) {
                throw new IllegalArgumentException("fontSize must be positive");
            }
            this.fontSize = fontSize;
            return this;
        }

        public ColumnConfig<T> width(int fixedWidth) {
            this.fixedWidth = true;
            this.minWidth = fixedWidth;
            return this;
        }

        public ColumnConfig<T> minWidth(int minWidth) {
            this.minWidth = minWidth;
            return this;
        }

        public ColumnConfig<T> maxWidth(int maxWidth) {
            this.maxWidth = maxWidth;
            return this;
        }

        public ColumnConfig<T> dropdown(String... options) {
            this.dropdownOptions = options;
            return this;
        }

        /**
         * Sets a per-cell conditional color function.
         */
        public ColumnConfig<T> cellColor(CellColorFunction<T> cellColorFunction) {
            this.cellColorFunction = cellColorFunction;
            return this;
        }

        /**
         * Sets the group header name for this column.
         */
        public ColumnConfig<T> group(String groupName) {
            this.groupName = groupName;
            return this;
        }

        /**
         * Sets the outline (grouping) level for this column.
         * Columns with outline level > 0 can be collapsed/expanded in Excel.
         *
         * @param level the outline level (1-7, 0 = no outline)
         */
        public ColumnConfig<T> outline(int level) {
            if (level < 0 || level > 7) {
                throw new IllegalArgumentException("outline level must be between 0 and 7");
            }
            this.outlineLevel = level;
            return this;
        }
    }
}
