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
    private float rowHeightInPoints = 20;
    private @Nullable Function<T, @Nullable ExcelColor> rowColorFunction;
    private @Nullable ProgressCallback progressCallback;
    private int progressInterval;
    private int autoWidthSampleRows = ExcelWriteSupport.AUTO_WIDTH_SAMPLE_ROWS;
    private @Nullable AfterDataWriter afterDataWriter;
    private @Nullable ExcelSummary summaryConfig;
    private ColumnStyleConfig.@Nullable DefaultStyleConfig<T> defaultStyleConfig;

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
     */
    public TemplateListWriter<T> column(String name, Function<T, @Nullable Object> function) {
        columns.add(buildColumn(name, (r, c) -> function.apply(r), null));
        return this;
    }

    /**
     * Adds a column with additional configuration.
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
     */
    public TemplateListWriter<T> column(String name, ExcelRowFunction<T, @Nullable Object> function) {
        columns.add(buildColumn(name, function, null));
        return this;
    }

    /**
     * Adds a column using a row function with cursor support and additional configuration.
     */
    public TemplateListWriter<T> column(String name, ExcelRowFunction<T, @Nullable Object> function,
                                         Consumer<ColumnConfig<T>> cfg) {
        ColumnConfig<T> config = new ColumnConfig<>();
        cfg.accept(config);
        columns.add(buildColumn(name, function, config));
        return this;
    }

    /**
     * Sets the row height for data rows in points.
     */
    public TemplateListWriter<T> rowHeight(float rowHeightInPoints) {
        this.rowHeightInPoints = rowHeightInPoints;
        return this;
    }

    /**
     * Sets a function that determines the background color for each row.
     */
    public TemplateListWriter<T> rowColor(Function<T, @Nullable ExcelColor> fn) {
        this.rowColorFunction = fn;
        return this;
    }

    /**
     * Registers a progress callback that fires every {@code interval} rows.
     */
    public TemplateListWriter<T> onProgress(int interval, ProgressCallback callback) {
        if (interval <= 0) {
            throw new IllegalArgumentException("progress interval must be positive");
        }
        this.progressInterval = interval;
        this.progressCallback = callback;
        return this;
    }

    /**
     * Sets the number of rows sampled for auto column width calculation.
     */
    public TemplateListWriter<T> autoWidthSampleRows(int rows) {
        if (rows < 0) {
            throw new IllegalArgumentException("autoWidthSampleRows must be non-negative");
        }
        this.autoWidthSampleRows = rows;
        return this;
    }

    /**
     * Registers a callback that writes content after all data rows.
     */
    public TemplateListWriter<T> afterData(AfterDataWriter writer) {
        this.afterDataWriter = writer;
        return this;
    }

    /**
     * Configures summary (footer) rows with formulas.
     */
    public TemplateListWriter<T> summary(Consumer<ExcelSummary> configurer) {
        ExcelSummary summary = new ExcelSummary();
        configurer.accept(summary);
        this.summaryConfig = summary;
        return this;
    }

    /**
     * Sets default column styles that apply to all columns unless overridden per-column.
     */
    public TemplateListWriter<T> defaultStyle(Consumer<ColumnStyleConfig.DefaultStyleConfig<T>> configurer) {
        ColumnStyleConfig.DefaultStyleConfig<T> config = new ColumnStyleConfig.DefaultStyleConfig<>();
        configurer.accept(config);
        this.defaultStyleConfig = config;
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

    private ExcelTemplateWriter writeInternal(Stream<T> stream, boolean writeHeaders) {
        if (columns.isEmpty()) {
            throw new ExcelWriteException("columns setting required");
        }
        ExcelWriteSupport.validateUniqueColumnNames(columns);

        Cursor cursor = new Cursor(startRow);
        int headerRowIndex = startRow;

        if (writeHeaders) {
            CellStyle headerStyle = ExcelStyleSupporter.headerStyle(wb,
                    new org.apache.poi.xssf.usermodel.XSSFColor(
                            new byte[]{(byte) 255, (byte) 255, (byte) 255}));
            ExcelWriteSupport.writeColumnHeaders(sheet, cursor, columns, headerStyle);
        }

        try (stream) {
            stream.forEach(rowData -> {
                cursor.plusTotal();
                ExcelWriteSupport.writeRowCells(sheet, cursor, rowData, columns, rowHeightInPoints,
                        rowColorFunction, rowStyleCache, wb, autoWidthSampleRows);
                ExcelWriteSupport.checkProgress(cursor, progressInterval, progressCallback);
            });
        }

        int nextRow = cursor.getRowOfSheet();
        if (afterDataWriter != null) {
            nextRow = afterDataWriter.write(new SheetContext(sheet, wb, nextRow, columns, headerRowIndex));
        }
        if (summaryConfig != null) {
            nextRow = summaryConfig.toAfterDataWriter().write(new SheetContext(sheet, wb, nextRow, columns, headerRowIndex));
        }

        ExcelWriteSupport.applyColumnWidths(sheet, columns);

        parent.updateLastWrittenRow(sheetIndex, nextRow - 1);
        return parent;
    }

    private ExcelColumn<T> buildColumn(String name, ExcelRowFunction<T, @Nullable Object> function,
                                        @Nullable ColumnConfig<T> config) {
        ColumnStyleConfig<T, ?> c = config != null ? config : new ColumnConfig<>();
        if (defaultStyleConfig != null) {
            c.applyDefaults(defaultStyleConfig);
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
