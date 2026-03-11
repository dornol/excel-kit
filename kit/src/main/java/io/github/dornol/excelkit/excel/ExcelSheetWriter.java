package io.github.dornol.excelkit.excel;

import io.github.dornol.excelkit.shared.Cursor;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.function.Consumer;
import java.util.function.Function;
import java.util.stream.Stream;


/**
 * Writes data of a specific type to a single sheet within an {@link ExcelWorkbook}.
 * <p>
 * Unlike {@link ExcelWriter}, this class does not perform automatic sheet rollover.
 * It is designed for explicit multi-sheet workbooks where each sheet has its own data type.
 *
 * @param <T> the row data type for this sheet
 * @author dhkim
 */
public class ExcelSheetWriter<T> {

    // Shared resources from ExcelWorkbook
    private final SXSSFWorkbook wb;
    private final SXSSFSheet sheet;
    private final CellStyle headerStyle;
    private final Map<String, CellStyle> cellStyleCache;

    // Per-sheet settings
    private final List<ExcelColumn<T>> columns = new ArrayList<>();
    private float rowHeightInPoints = 20;
    private boolean autoFilter = false;
    private int freezePaneRows = 0;
    private BeforeHeaderWriter beforeHeaderWriter;
    private AfterDataWriter afterDataWriter;
    private Function<T, ExcelColor> rowColorFunction;
    private final Map<String, CellStyle> rowStyleCache = new HashMap<>();

    ExcelSheetWriter(SXSSFWorkbook wb, SXSSFSheet sheet, CellStyle headerStyle, Map<String, CellStyle> cellStyleCache) {
        this.wb = wb;
        this.sheet = sheet;
        this.headerStyle = headerStyle;
        this.cellStyleCache = cellStyleCache;
    }

    /**
     * Adds a column using a simple function.
     *
     * @param name     Column header name
     * @param function Function to extract cell value from row
     * @return this writer for chaining
     */
    public ExcelSheetWriter<T> column(String name, Function<T, Object> function) {
        columns.add(buildColumn(name, (r, c) -> function.apply(r), null));
        return this;
    }

    /**
     * Adds a column with additional configuration.
     *
     * @param name     Column header name
     * @param function Function to extract cell value from row
     * @param cfg      Consumer to configure column options
     * @return this writer for chaining
     */
    public ExcelSheetWriter<T> column(String name, Function<T, Object> function, Consumer<ColumnConfig<T>> cfg) {
        ColumnConfig<T> config = new ColumnConfig<>();
        cfg.accept(config);
        columns.add(buildColumn(name, (r, c) -> function.apply(r), config));
        return this;
    }

    /**
     * Adds a column using a row function with cursor support.
     *
     * @param name     Column header name
     * @param function Function to extract cell value from row with cursor
     * @return this writer for chaining
     */
    public ExcelSheetWriter<T> column(String name, ExcelRowFunction<T, Object> function) {
        columns.add(buildColumn(name, function, null));
        return this;
    }

    /**
     * Adds a column using a row function with cursor support and additional configuration.
     *
     * @param name     Column header name
     * @param function Function to extract cell value from row with cursor
     * @param cfg      Consumer to configure column options
     * @return this writer for chaining
     */
    public ExcelSheetWriter<T> column(String name, ExcelRowFunction<T, Object> function, Consumer<ColumnConfig<T>> cfg) {
        ColumnConfig<T> config = new ColumnConfig<>();
        cfg.accept(config);
        columns.add(buildColumn(name, function, config));
        return this;
    }

    /**
     * Adds a column with a constant value for all rows.
     *
     * @param name  Column header name
     * @param value Constant value
     * @return this writer for chaining
     */
    public ExcelSheetWriter<T> constColumn(String name, Object value) {
        columns.add(buildColumn(name, (r, c) -> value, null));
        return this;
    }

    /**
     * Sets the row height for data rows.
     *
     * @param rowHeightInPoints Row height in points
     * @return this writer for chaining
     */
    public ExcelSheetWriter<T> rowHeight(float rowHeightInPoints) {
        this.rowHeightInPoints = rowHeightInPoints;
        return this;
    }

    /**
     * Enables auto-filter on the header row.
     *
     * @return this writer for chaining
     */
    public ExcelSheetWriter<T> autoFilter() {
        this.autoFilter = true;
        return this;
    }

    /**
     * Freezes the specified number of rows below the header.
     *
     * @param rows Number of rows to freeze
     * @return this writer for chaining
     */
    public ExcelSheetWriter<T> freezePane(int rows) {
        this.freezePaneRows = rows;
        return this;
    }

    /**
     * Registers a callback to write content before the header row.
     *
     * @param writer the callback
     * @return this writer for chaining
     */
    public ExcelSheetWriter<T> beforeHeader(BeforeHeaderWriter writer) {
        this.beforeHeaderWriter = writer;
        return this;
    }

    /**
     * Registers a callback to write content after all data rows.
     *
     * @param writer the callback
     * @return this writer for chaining
     */
    public ExcelSheetWriter<T> afterData(AfterDataWriter writer) {
        this.afterDataWriter = writer;
        return this;
    }

    /**
     * Sets a function that determines the background color for each row.
     *
     * @param fn function that takes row data and returns an ExcelColor (or null for no override)
     * @return this writer for chaining
     */
    public ExcelSheetWriter<T> rowColor(Function<T, ExcelColor> fn) {
        this.rowColorFunction = fn;
        return this;
    }

    /**
     * Writes the data stream to this sheet.
     * <p>
     * This method configures the sheet (title, headers, options),
     * writes all data rows, applies callbacks, column widths, and data validations.
     *
     * @param stream the data stream to write
     */
    public void write(Stream<T> stream) {
        if (columns.isEmpty()) {
            throw new ExcelWriteException("columns setting required");
        }

        int currentRow = ExcelWriteSupport.initSheetPreamble(sheet, wb, columns, beforeHeaderWriter);
        Cursor cursor = new Cursor(currentRow);
        int headerRowIndex = currentRow;

        ExcelWriteSupport.writeColumnHeaders(sheet, cursor, columns, headerStyle);
        int headerRowIdx = cursor.getRowOfSheet() - 1;
        ExcelWriteSupport.applySheetOptions(sheet, headerRowIdx, autoFilter, freezePaneRows, columns.size());

        try (stream) {
            stream.forEach(rowData -> {
                cursor.plusTotal();
                ExcelWriteSupport.writeRowCells(sheet, cursor, rowData, columns, rowHeightInPoints,
                        rowColorFunction, rowStyleCache, wb);
            });
        }

        int nextRow = cursor.getRowOfSheet();
        if (this.afterDataWriter != null) {
            this.afterDataWriter.write(new SheetContext(sheet, wb, nextRow, columns));
        }

        ExcelWriteSupport.applyColumnWidths(sheet, columns);
        ExcelWriteSupport.applyDataValidations(sheet, columns, headerRowIndex);
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
        }

        if (dataFormat == null) {
            dataFormat = dataType.getDefaultFormat();
        }

        CellStyle style = ExcelStyleSupporter.cellStyle(wb, alignment, dataFormat,
                backgroundColor, bold, fontSize, cellStyleCache);
        ExcelColumnSetter setter = dataType.getSetter();

        return new ExcelColumn<>(name, function, style, setter, minWidth, maxWidth, fixedWidth, dropdownOptions);
    }

    /**
     * Configuration class for column options, used with {@code Consumer<ColumnConfig>} pattern.
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

        /**
         * Sets the data type for the column.
         */
        public ColumnConfig<T> type(ExcelDataType dataType) {
            this.dataType = dataType;
            return this;
        }

        /**
         * Sets the Excel data format string.
         */
        public ColumnConfig<T> format(String dataFormat) {
            this.dataFormat = dataFormat;
            return this;
        }

        /**
         * Sets the horizontal alignment.
         */
        public ColumnConfig<T> alignment(HorizontalAlignment alignment) {
            this.alignment = alignment;
            return this;
        }

        /**
         * Sets the background color using RGB values.
         */
        public ColumnConfig<T> backgroundColor(int r, int g, int b) {
            this.backgroundColor = new int[]{r, g, b};
            return this;
        }

        /**
         * Sets the background color using a preset color.
         */
        public ColumnConfig<T> backgroundColor(ExcelColor color) {
            return backgroundColor(color.getR(), color.getG(), color.getB());
        }

        /**
         * Sets whether the font should be bold.
         */
        public ColumnConfig<T> bold(boolean bold) {
            this.bold = bold;
            return this;
        }

        /**
         * Sets the font size in points.
         */
        public ColumnConfig<T> fontSize(int fontSize) {
            if (fontSize <= 0) {
                throw new IllegalArgumentException("fontSize must be positive");
            }
            this.fontSize = fontSize;
            return this;
        }

        /**
         * Sets a fixed column width.
         */
        public ColumnConfig<T> width(int fixedWidth) {
            this.fixedWidth = true;
            this.minWidth = fixedWidth;
            return this;
        }

        /**
         * Sets the minimum column width.
         */
        public ColumnConfig<T> minWidth(int minWidth) {
            this.minWidth = minWidth;
            return this;
        }

        /**
         * Sets the maximum column width.
         */
        public ColumnConfig<T> maxWidth(int maxWidth) {
            this.maxWidth = maxWidth;
            return this;
        }

        /**
         * Sets dropdown validation options.
         */
        public ColumnConfig<T> dropdown(String... options) {
            this.dropdownOptions = options;
            return this;
        }
    }
}
