package io.github.dornol.excelkit.excel;

import io.github.dornol.excelkit.core.ExcelKitException;
import io.github.dornol.excelkit.core.read.RowFunction;
import io.github.dornol.excelkit.core.Cursor;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFColor;

import org.jspecify.annotations.Nullable;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Objects;
import java.util.stream.Stream;
import java.util.stream.StreamSupport;

/**
 * Writes tabular (list) data into an existing template sheet starting at a given row.
 * <p>
 * Created by {@link ExcelTemplateWriter#list(int)} or {@link ExcelTemplateWriter#list(int, int)}.
 * The columns and the settings of how the rows are written are those of every writer ({@link DataRowsWriter}:
 * {@code column}, {@code rowColor}, {@code summary}, {@code table}, {@code writeErrorPolicy}...), written with the same
 * utilities as {@link ExcelSheetWriter}. The sheet settings of a {@link SheetWriter} (frozen panes, filter, chart,
 * rollover...) are not here: the template sheet has them already.
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
public class TemplateListWriter<T> extends DataRowsWriter<T, TemplateListWriter<T>> {

    private final ExcelTemplateWriter parent;
    private final SXSSFWorkbook wb;
    private final SXSSFSheet sheet;
    private final int startRow;
    private final Map<String, CellStyle> cellStyleCache;
    private final int sheetIndex;

    private final List<ExcelColumn<T>> columns = new ArrayList<>();
    private final Map<ExcelStyleSupporter.DerivedStyleKey, CellStyle> rowStyleCache = new HashMap<>();

    TemplateListWriter(ExcelTemplateWriter parent, SXSSFWorkbook wb, SXSSFSheet sheet,
                       int startRow, Map<String, CellStyle> cellStyleCache, int sheetIndex) {
        super(new SheetConfig<>());
        this.parent = parent;
        this.wb = wb;
        this.sheet = sheet;
        this.startRow = startRow;
        this.cellStyleCache = cellStyleCache;
        this.sheetIndex = sheetIndex;
    }












    /**
     * Writes the data stream to this sheet starting at the configured start row.
     * <p>
     * Does <b>not</b> write column headers — the template is expected to have them already.
     * Use {@link #writeWithHeaders(Stream)} if headers should be written.
     *
     * <p>
     * The stream is read but <b>not closed</b>: whoever opened it (a JPA or JDBC stream, for example)
     * closes it, typically with try-with-resources around this call.
     * <p>
     * Rows are written one at a time on the calling thread, in the stream's encounter order: a parallel
     * stream is read sequentially, so {@code parallel()} brings no speed-up here.
     *
     * @param stream the data stream
     * @return the parent {@link ExcelTemplateWriter} for further chaining
     */
    public ExcelTemplateWriter write(Stream<T> stream) {
        return writeInternal(stream, false);
    }

    /**
     * Like {@link #write(Stream)}, for an {@code Iterable}.
     *
     * @param rows the data rows
     * @return the parent {@link ExcelTemplateWriter} for further chaining
     */
    public ExcelTemplateWriter write(Iterable<T> rows) {
        Objects.requireNonNull(rows, "rows cannot be null");
        return write(StreamSupport.stream(rows.spliterator(), false));
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

    /**
     * Like {@link #writeWithHeaders(Stream)}, for an {@code Iterable}.
     *
     * @param rows the data rows
     * @return the parent {@link ExcelTemplateWriter} for further chaining
     */
    public ExcelTemplateWriter writeWithHeaders(Iterable<T> rows) {
        Objects.requireNonNull(rows, "rows cannot be null");
        return writeWithHeaders(StreamSupport.stream(rows.spliterator(), false));
    }

    private ExcelTemplateWriter writeInternal(Stream<T> stream, boolean writeHeaders) {
        Objects.requireNonNull(stream, "stream cannot be null");
        try {
            return writeRows(stream, writeHeaders);
        } catch (ExcelKitException e) {
            throw e;
        } catch (RuntimeException e) {
            // as in ExcelWriter: what goes wrong while the rows are written (a style or a part of a damaged template
            // that cannot be used, ...) becomes a library exception
            throw new ExcelWriteException("Failed to write the list into the template", e);
        }
    }

    private ExcelTemplateWriter writeRows(Stream<T> stream, boolean writeHeaders) {
        if (columns.isEmpty()) {
            throw new ExcelWriteException("columns setting required");
        }
        SheetWriteEngine.validateUniqueColumnNames(columns);
        SheetWriteEngine.validateMergeColumns(columns, null, false);

        Cursor cursor = new Cursor(startRow);
        // The header row is the one this writer writes, or else the template's own, just above the data. A summary
        // and afterData start counting from it: taking startRow for it left out the first data row of a summary.
        int headerRowIndex = writeHeaders ? startRow : startRow - 1;

        // Read once: a change made while the rows are being written does not apply to this write.
        TableOptions tableOptions = cfg.tableOptions;
        List<String> tableHeaders = tableOptions == null ? List.of() : columns.stream().map(ExcelColumn::getName).toList();
        if (tableOptions != null && !writeHeaders) {
            if (startRow == 0) throw new ExcelWriteException("Template table requires a header row before startRow");
            tableHeaders = StructuredTableWriter.validateExistingHeaders(sheet, startRow - 1, columns.size());
        }

        if (writeHeaders) {
            CellStyle headerStyle = ExcelStyleSupporter.headerStyle(wb,
                    new XSSFColor(
                            new byte[]{(byte) 255, (byte) 255, (byte) 255}));
            ExcelHeaderWriter.write(sheet, cursor, columns, headerStyle, null, null, null, 0f, false);
        }

        // What the summary formulas calculate, written as their cached results so a reader that does not recalculate sees them.
        FormulaValues values = FormulaValues.of(cfg, columns);
        stream.sequential().forEach(rowData -> {
            cursor.plusTotal();
            if (cursor.getRowOfSheet() >= ExcelLimits.SHEET_ROW_CAPACITY) {
                throw new ExcelWriteException("Sheet '" + sheet.getSheetName() + "' is full: Excel sheets hold at most "
                        + ExcelLimits.SHEET_ROW_CAPACITY + " rows, and a list written into a template stays on its sheet. "
                        + "Write fewer rows into it, or use ExcelWriter, which rolls over to new sheets.");
            }
            ExcelRowWriter.write(sheet, cursor, rowData, columns, cfg, rowStyleCache, wb, null, null, values);
            SheetWriteEngine.checkProgress(cursor, cfg.progressInterval, cfg.progressCallback);
        });

        int nextRow = SheetWriteEngine.writeAfterDataAndSummary(sheet, wb, cursor.getRowOfSheet(), columns, headerRowIndex, cfg, values);

        if (tableOptions != null) {
            int tableHeaderRow = writeHeaders ? startRow : startRow - 1;
            if (tableHeaderRow < 0) {
                throw new ExcelWriteException("Template table requires an existing header row before startRow");
            }
            int lastDataRow = cursor.getRowOfSheet() - 1;
            if (lastDataRow > tableHeaderRow) {
                StructuredTableWriter.apply(sheet, tableOptions.name(), tableHeaderRow, lastDataRow,
                        tableHeaders, tableOptions.style(), tableOptions.showRowStripes(), false);
            }
        }

        ExcelSheetPostProcessor.applyTemplateColumnWidths(sheet, columns);

        parent.updateLastWrittenRow(sheetIndex, nextRow - 1);
        return parent;
    }

    @Override
    void addColumn(ExcelColumn<T> column) {
        columns.add(column);
    }

    @Override
    ExcelColumn<T> buildColumn(String name, RowFunction<T, @Nullable Object> function,
                                       @Nullable ColumnConfig<T> config) {
        return ColumnsWriter.buildColumn(wb, cellStyleCache, cfg.defaultStyleConfig, name, function, config);
    }

}
