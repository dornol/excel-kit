package io.github.dornol.excelkit.excel;

import io.github.dornol.excelkit.core.RowFunction;
import io.github.dornol.excelkit.core.Cursor;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

import org.jspecify.annotations.Nullable;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.Iterator;
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
 * <p>
 * Column configuration uses {@link ColumnConfig} (via {@code Consumer<ColumnConfig<T>>})
 * rather than {@link ExcelColumn.ExcelColumnBuilder}. Unlike ExcelWriter's builder, there
 * is no {@code style(CellStyle)} method — styles are derived from the declarative config
 * properties (type, bold, color, borders, etc.).
 *
 * @param <T> the row data type for this sheet
 * @author dhkim
 */
public class ExcelSheetWriter<T> extends AbstractSheetWriter<T, ExcelSheetWriter<T>> {

    // Shared resources from ExcelWorkbook
    private final SXSSFWorkbook wb;
    private SXSSFSheet sheet;
    private final String baseName;
    private final CellStyle headerStyle;
    private final Map<String, CellStyle> cellStyleCache;
    private final Set<String> usedSheetNames;

    // Per-sheet settings
    private final List<ExcelColumn<T>> columns = new ArrayList<>();
    private final Map<String, CellStyle> rowStyleCache = new HashMap<>();
    private final Map<String, CellStyle> headerStyleCache = new HashMap<>();
    private int maxRows = Integer.MAX_VALUE;
    private boolean written = false;

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
     *
     * @param name the column header
     * @param function function to extract the cell value
     * @return this writer for chaining
     */
    public ExcelSheetWriter<T> column(String name, Function<T, @Nullable Object> function) {
        columns.add(buildColumn(name, (r, c) -> function.apply(r), null));
        return this;
    }

    /**
     * Adds a column with additional configuration.
     *
     * @param name the column header
     * @param function function to extract the cell value
     * @param cfg consumer to configure column styling
     * @return this writer for chaining
     */
    public ExcelSheetWriter<T> column(String name, Function<T, @Nullable Object> function, Consumer<ColumnConfig<T>> cfg) {
        ColumnConfig<T> config = new ColumnConfig<>();
        cfg.accept(config);
        columns.add(buildColumn(name, (r, c) -> function.apply(r), config));
        return this;
    }

    /**
     * Adds a column using a row function with cursor support.
     *
     * @param name the column header
     * @param function function to extract the cell value
     * @return this writer for chaining
     */
    public ExcelSheetWriter<T> column(String name, RowFunction<T, @Nullable Object> function) {
        columns.add(buildColumn(name, function, null));
        return this;
    }

    /**
     * Adds a column using a row function with cursor support and additional configuration.
     *
     * @param name the column header
     * @param function function to extract the cell value
     * @param cfg consumer to configure column styling
     * @return this writer for chaining
     */
    public ExcelSheetWriter<T> column(String name, RowFunction<T, @Nullable Object> function, Consumer<ColumnConfig<T>> cfg) {
        ColumnConfig<T> config = new ColumnConfig<>();
        cfg.accept(config);
        columns.add(buildColumn(name, function, config));
        return this;
    }

    /**
     * Adds a column with a constant value for all rows.
     *
     * @param name the column header
     * @param value the constant value
     * @return this writer for chaining
     */
    public ExcelSheetWriter<T> constColumn(String name, @Nullable Object value) {
        columns.add(buildColumn(name, (r, c) -> value, null));
        return this;
    }

    /**
     * Enables auto-filter on the header row (no-arg convenience).
     *
     * @return this writer for chaining
     */
    public ExcelSheetWriter<T> autoFilter() {
        return autoFilter(true);
    }

    /**
     * Adds a 1-based sequential row-number column.
     *
     * @since 0.16.11
     */
    public ExcelSheetWriter<T> rowNumberColumn(String name) {
        return column(name, (row, cursor) -> cursor.getCurrentTotal(),
                c -> c.type(ExcelDataType.LONG));
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
     * Writes the data stream to this sheet (with optional auto-rollover).
     *
     * @param stream the data stream to write
     */
    public void write(Stream<T> stream) {
        if (written) {
            throw new ExcelWriteException("write() has already been called on this sheet");
        }
        written = true;
        if (columns.isEmpty()) {
            throw new ExcelWriteException("columns setting required");
        }
        ExcelWriteSupport.validateUniqueColumnNames(columns);

        List<SXSSFSheet> allSheets = new ArrayList<>();
        allSheets.add(this.sheet);

        int currentRow = ExcelWriteSupport.initSheetPreamble(sheet, wb, columns, cfg.beforeHeaderWriter);
        Cursor cursor = new Cursor(currentRow);
        int headerRowIndex = currentRow;

        ExcelWriteSupport.writeColumnHeaders(sheet, cursor, columns, headerStyle, wb, headerStyleCache, cfg.groupComments, cfg.headerRowHeightInPoints);
        int headerRowIdx = cursor.getRowOfSheet() - 1;
        ExcelWriteSupport.applySheetOptions(sheet, headerRowIdx, cfg.autoFilter, cfg.freezePaneCols, cfg.freezePaneRows, columns.size());

        SXSSFSheet activeSheet = this.sheet;

        {
            Iterator<T> it = stream.iterator();
            while (it.hasNext()) {
                T rowData = it.next();
                cursor.plusTotal();
                if (maxRows != Integer.MAX_VALUE && cursor.getCurrentTotal() >= maxRows
                        && cursor.getCurrentTotal() % maxRows == 1) {
                    ExcelWriteSupport.writeAfterDataAndSummary(activeSheet, wb, cursor.getRowOfSheet(), columns, headerRowIndex, cfg);
                    activeSheet = createRolloverSheet(allSheets.size());
                    allSheets.add(activeSheet);
                    cursor.initRow();
                    int preambleRow = ExcelWriteSupport.initSheetPreamble(activeSheet, wb, columns, cfg.beforeHeaderWriter);
                    cursor.setRowOfSheet(preambleRow);
                    headerRowIndex = preambleRow;
                    ExcelWriteSupport.writeColumnHeaders(activeSheet, cursor, columns, headerStyle, wb, headerStyleCache, cfg.groupComments, cfg.headerRowHeightInPoints);
                    int hdrIdx = cursor.getRowOfSheet() - 1;
                    ExcelWriteSupport.applySheetOptions(activeSheet, hdrIdx, cfg.autoFilter, cfg.freezePaneCols, cfg.freezePaneRows, columns.size());
                }
                ExcelWriteSupport.writeRowCells(activeSheet, cursor, rowData, columns, cfg, rowStyleCache, wb);
                ExcelWriteSupport.checkProgress(cursor, cfg.progressInterval, cfg.progressCallback);
            }
        }

        ExcelWriteSupport.writeAfterDataAndSummary(activeSheet, wb, cursor.getRowOfSheet(), columns, headerRowIndex, cfg);

        for (SXSSFSheet s : allSheets) {
            ExcelWriteSupport.applyPostProcessing(s, columns, headerRowIndex, cfg);
        }

        // Apply chart on last sheet
        if (cfg.chartConfig != null) {
            SXSSFSheet lastSheet = allSheets.get(allSheets.size() - 1);
            ExcelWriteSupport.applyChart(lastSheet, cfg.chartConfig, headerRowIndex, cursor.getRowOfSheet() - 1);
        }
    }

    /** Writes rows from an Iterable without copying them. */
    public void write(Iterable<T> rows) {
        java.util.Objects.requireNonNull(rows, "rows cannot be null");
        write(java.util.stream.StreamSupport.stream(rows.spliterator(), false));
    }

    private SXSSFSheet createRolloverSheet(int rolloverIndex) {
        String name;
        if (cfg.sheetNameFunction != null) {
            name = cfg.sheetNameFunction.apply(rolloverIndex);
        } else {
            name = baseName + " (" + (rolloverIndex + 1) + ")";
        }
        if (usedSheetNames != null && !usedSheetNames.add(name)) {
            throw new ExcelWriteException("Duplicate sheet name: '" + name + "'");
        }
        return wb.createSheet(name);
    }

    private ExcelColumn<T> buildColumn(String name, RowFunction<T, @Nullable Object> function, @Nullable ColumnConfig<T> config) {
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

        return new ExcelColumn<>(name, function, style, dataType.getSetter(), c);
    }

}
