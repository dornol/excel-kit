package io.github.dornol.excelkit.excel;

import io.github.dornol.excelkit.core.RowFunction;
import io.github.dornol.excelkit.core.Cursor;
import io.github.dornol.excelkit.core.ProgressCallback;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFColor;

import org.jspecify.annotations.Nullable;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.function.Consumer;
import java.util.function.Function;
import java.util.stream.Stream;

/**
 * Writes tabular (list) data into an existing template sheet starting at a given row.
 * <p>
 * Created by {@link ExcelTemplateWriter#list(int)} or {@link ExcelTemplateWriter#list(int, int)}.
 * Reuses the same column definition and write utilities as {@link ExcelSheetWriter}.
 *
 * <pre>{@code
 * writer.<Item>list(5)
 *     .column("A", Item::getName)
 *     .column("B", Item::getQty, c -> c.type(ExcelDataType.INTEGER))
 *     .afterData(ctx -> {
 *         ctx.getSheet().createRow(ctx.getCurrentRow())
 *            .createCell(0).setCellValue("Total");
 *         return ctx.getCurrentRow() + 1;
 *     })
 *     .write(itemStream);
 * }</pre>
 *
 * @param <T> the row data type
 * @author dhkim
 * @since 0.8.2
 */
public class TemplateListWriter<T> {

    private final ExcelTemplateWriter parent;
    private final SXSSFWorkbook wb;
    private final SXSSFSheet sheet;
    private final int startRow;
    private final Map<String, CellStyle> cellStyleCache;
    private final int sheetIndex;

    private final List<ExcelColumn<T>> columns = new ArrayList<>();
    private final Map<String, CellStyle> rowStyleCache = new HashMap<>();
    private final SheetConfig<T> cfg = new SheetConfig<>();
    private @Nullable TableOptions tableOptions;

    TemplateListWriter(ExcelTemplateWriter parent, SXSSFWorkbook wb, SXSSFSheet sheet,
                       int startRow, Map<String, CellStyle> cellStyleCache, int sheetIndex) {
        this.parent = parent;
        this.wb = wb;
        this.sheet = sheet;
        this.startRow = startRow;
        this.cellStyleCache = cellStyleCache;
        this.sheetIndex = sheetIndex;
    }

    /**
     * Adds a column using a simple function.
     *
     * @param name the column header
     * @param function function to extract the cell value
     * @return this writer for chaining
     */
    public TemplateListWriter<T> column(String name, Function<T, @Nullable Object> function) {
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
    public TemplateListWriter<T> column(String name, Function<T, @Nullable Object> function,
                                         Consumer<ColumnConfig<T>> cfg) {
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
    public TemplateListWriter<T> column(String name, RowFunction<T, @Nullable Object> function) {
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
    public TemplateListWriter<T> column(String name, RowFunction<T, @Nullable Object> function,
                                         Consumer<ColumnConfig<T>> cfg) {
        ColumnConfig<T> config = new ColumnConfig<>();
        cfg.accept(config);
        columns.add(buildColumn(name, function, config));
        return this;
    }

    /**
     * Sets the row height for data rows in points.
     *
     * @param rowHeightInPoints row height in points
     * @return this writer for chaining
     */
    public TemplateListWriter<T> rowHeight(float rowHeightInPoints) {
        this.cfg.rowHeightInPoints = rowHeightInPoints;
        return this;
    }

    /**
     * Sets a function that determines the background color for each row.
     *
     * @param fn function returning a color per row, or null
     * @return this writer for chaining
     */
    public TemplateListWriter<T> rowColor(Function<T, @Nullable ExcelColor> fn) {
        this.cfg.rowColorFunction = fn;
        return this;
    }

    /**
     * Adds a conditional row style that applies to all cells in a row when the predicate matches.
     *
     * @param predicate  condition to test each row
     * @param configurer style configuration to apply when the condition is true
     * @return this writer for chaining
     */
    public TemplateListWriter<T> rowStyle(java.util.function.Predicate<T> predicate,
                                           java.util.function.Consumer<RowStyleConfig> configurer) {
        RowStyleConfig style = new RowStyleConfig();
        configurer.accept(style);
        cfg.rowStyleEntries.add(new SheetConfig.RowStyleEntry<>(predicate, style));
        return this;
    }

    /**
     * Registers a progress callback that fires every {@code interval} rows.
     *
     * @param interval rows between each callback
     * @param callback the callback to invoke
     * @return this writer for chaining
     */
    public TemplateListWriter<T> onProgress(int interval, ProgressCallback callback) {
        if (interval <= 0) {
            throw new IllegalArgumentException("progress interval must be positive");
        }
        this.cfg.progressInterval = interval;
        this.cfg.progressCallback = callback;
        return this;
    }

    /**
     * Sets the number of rows sampled for auto column width calculation.
     *
     * @param rows number of rows to sample
     * @return this writer for chaining
     */
    public TemplateListWriter<T> autoWidthSampleRows(int rows) {
        if (rows < 0) {
            throw new IllegalArgumentException("autoWidthSampleRows must be non-negative");
        }
        this.cfg.autoWidthSampleRows = rows;
        return this;
    }

    /**
     * Registers a callback that writes content after all data rows.
     *
     * @param writer the after-data writer callback
     * @return this writer for chaining
     */
    public TemplateListWriter<T> afterData(AfterDataWriter writer) {
        this.cfg.afterDataWriter = writer;
        return this;
    }

    /**
     * Configures summary (footer) rows with formulas.
     *
     * @param configurer consumer to configure the summary
     * @return this writer for chaining
     */
    public TemplateListWriter<T> summary(Consumer<ExcelSummary> configurer) {
        ExcelSummary summary = new ExcelSummary();
        configurer.accept(summary);
        this.cfg.summaryConfig = summary;
        return this;
    }

    public TemplateListWriter<T> table(String name) {
        return table(TableOptions.defaults(name));
    }

    public TemplateListWriter<T> table(TableOptions options) {
        StructuredTableWriter.validateName(java.util.Objects.requireNonNull(options, "options cannot be null").name());
        this.tableOptions = options;
        return this;
    }

    /**
     * Sets default column styles that apply to all columns unless overridden per-column.
     *
     * @param configurer consumer to configure default styles
     * @return this writer for chaining
     */
    public TemplateListWriter<T> defaultStyle(Consumer<ColumnStyleConfig.DefaultStyleConfig<T>> configurer) {
        ColumnStyleConfig.DefaultStyleConfig<T> config = new ColumnStyleConfig.DefaultStyleConfig<>();
        configurer.accept(config);
        this.cfg.defaultStyleConfig = config;
        return this;
    }

    /**
     * Writes the data stream to this sheet starting at the configured start row.
     * <p>
     * Does <b>not</b> write column headers — the template is expected to have them already.
     * Use {@link #writeWithHeaders(Stream)} if headers should be written.
     *
     * @param stream the data stream
     * @return the parent {@link ExcelTemplateWriter} for further chaining
     */
    public ExcelTemplateWriter write(Stream<T> stream) {
        return writeInternal(stream, false);
    }

    public ExcelTemplateWriter write(Iterable<T> rows) {
        java.util.Objects.requireNonNull(rows, "rows cannot be null");
        return write(java.util.stream.StreamSupport.stream(rows.spliterator(), false));
    }

    /**
     * Writes column headers at the start row, followed by data rows.
     * <p>
     * Use this when the template does not have pre-existing column headers.
     *
     * @param stream the data stream
     * @return the parent {@link ExcelTemplateWriter} for further chaining
     */
    public ExcelTemplateWriter writeWithHeaders(Stream<T> stream) {
        return writeInternal(stream, true);
    }

    public ExcelTemplateWriter writeWithHeaders(Iterable<T> rows) {
        java.util.Objects.requireNonNull(rows, "rows cannot be null");
        return writeWithHeaders(java.util.stream.StreamSupport.stream(rows.spliterator(), false));
    }

    private ExcelTemplateWriter writeInternal(Stream<T> stream, boolean writeHeaders) {
        if (columns.isEmpty()) {
            throw new ExcelWriteException("columns setting required");
        }
        ExcelWriteSupport.validateUniqueColumnNames(columns);

        Cursor cursor = new Cursor(startRow);
        int headerRowIndex = startRow;

        if (tableOptions != null && !writeHeaders) {
            if (startRow == 0) throw new ExcelWriteException("Template table requires a header row before startRow");
            StructuredTableWriter.validateExistingHeaders(sheet, startRow - 1, columns.size());
        }

        if (writeHeaders) {
            CellStyle headerStyle = ExcelStyleSupporter.headerStyle(wb,
                    new XSSFColor(
                            new byte[]{(byte) 255, (byte) 255, (byte) 255}));
            ExcelHeaderWriter.write(sheet, cursor, columns, headerStyle);
        }

        stream.sequential().forEach(rowData -> {
            cursor.plusTotal();
            ExcelRowWriter.write(sheet, cursor, rowData, columns, cfg, rowStyleCache, wb);
            ExcelWriteSupport.checkProgress(cursor, cfg.progressInterval, cfg.progressCallback);
        });

        int nextRow = ExcelWriteSupport.writeAfterDataAndSummary(sheet, wb, cursor.getRowOfSheet(), columns, headerRowIndex, cfg);

        if (tableOptions != null) {
            int tableHeaderRow = writeHeaders ? startRow : startRow - 1;
            if (tableHeaderRow < 0) {
                throw new ExcelWriteException("Template table requires an existing header row before startRow");
            }
            int lastDataRow = cursor.getRowOfSheet() - 1;
            if (lastDataRow > tableHeaderRow) {
                StructuredTableWriter.apply(sheet, tableOptions.name(), tableHeaderRow, lastDataRow,
                        columns.size(), tableOptions.style(), tableOptions.showRowStripes());
            }
        }

        ExcelWriteSupport.applyColumnWidths(sheet, columns);

        parent.updateLastWrittenRow(sheetIndex, nextRow - 1);
        return parent;
    }

    private ExcelColumn<T> buildColumn(String name, RowFunction<T, @Nullable Object> function,
                                        @Nullable ColumnConfig<T> config) {
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
