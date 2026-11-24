package io.github.dornol.excelkit.excel;

import io.github.dornol.excelkit.core.Cursor;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.jspecify.annotations.Nullable;

import java.util.ArrayList;
import java.util.List;
import java.util.Map;
import java.util.function.IntFunction;

/**
 * Writes the rows of one list to a sheet of a workbook and to the sheets it rolls over to: header, rows, the subtotal,
 * after-data and summary rows, and then what needs every sheet finished (column widths, validations, tables, charts).
 * It is the one engine behind {@link ExcelWriter} (a workbook of its own) and {@link ExcelSheetWriter} (a sheet of an
 * {@link ExcelWorkbook}); the two differ only in how a sheet is made and named, and in what comes around the engine
 * (the handler, the workbook's protection, {@code afterAll}, who releases the workbook after a failure).
 * <p>
 * Use: {@link #start(SXSSFSheet)}, {@link #row(Object)} for every row, {@link #finishData()}, then
 * {@link #finishSheets()}. State is per run; an engine is used once.
 *
 * @param <T> the row type
 */
final class SheetWriteEngine<T> {
    private final SXSSFWorkbook wb;
    private final List<ExcelColumn<T>> columns;
    private final SheetConfig<T> cfg;
    private final CellStyle headerStyle;
    private final Map<String, CellStyle> headerStyleCache;
    private final Map<ExcelWriteSupport.DerivedStyleKey, CellStyle> rowStyleCache;
    private final @Nullable Map<String, CellStyle> sharedStyles;
    private final IntFunction<SXSSFSheet> nextSheet;
    private final PageBreakTracker<T> pageBreaks;
    private final @Nullable MergeRunTracker<T> merges;
    private final @Nullable SubtotalTracker<T> subtotals;
    private final @Nullable FormulaValues values;

    private final List<SXSSFSheet> sheets = new ArrayList<>();
    private final List<Integer> headerRows = new ArrayList<>();
    private final List<Integer> dataEndRows = new ArrayList<>();
    private final List<Integer> nextFreeRows = new ArrayList<>();
    private SXSSFSheet sheet;
    private Cursor cursor;

    /**
     * @param sharedStyles    the styles the sheets of a workbook share (the line of an empty sheet is one), or
     *                        {@code null} when the writer owns the workbook and has no other sheet to share with
     * @param nextSheet       makes the sheet a rolled-over list continues on, from its index in the list (the first
     *                        rolled-over sheet is 1); the writer decides how it is named
     */
    SheetWriteEngine(SXSSFWorkbook wb, List<ExcelColumn<T>> columns, SheetConfig<T> cfg, CellStyle headerStyle,
                     Map<String, CellStyle> headerStyleCache, Map<ExcelWriteSupport.DerivedStyleKey, CellStyle> rowStyleCache,
                     @Nullable Map<String, CellStyle> sharedStyles, IntFunction<SXSSFSheet> nextSheet) {
        this.wb = wb;
        this.columns = columns;
        this.cfg = cfg;
        this.headerStyle = headerStyle;
        this.headerStyleCache = headerStyleCache;
        this.rowStyleCache = rowStyleCache;
        this.sharedStyles = sharedStyles;
        this.nextSheet = nextSheet;
        this.pageBreaks = new PageBreakTracker<>(cfg.pageBreakWhen);
        this.merges = MergeRunTracker.of(columns);
        this.values = FormulaValues.of(cfg, columns);
        this.subtotals = SubtotalTracker.of(cfg, columns, values);
    }

    /** Refuses what no sheet can be written from, before any sheet (and its temp file) exists. */
    static <T> void validate(List<ExcelColumn<T>> columns, SheetConfig<T> cfg) {
        if (columns.isEmpty()) {
            throw new ExcelWriteException("columns setting required");
        }
        ExcelWriteSupport.validateUniqueColumnNames(columns);
        ExcelWriteSupport.validateMergeColumns(columns, cfg.tableOptions, true);
        SubtotalTracker.validate(cfg);
        ExcelWriteSupport.validateEmptyMessage(cfg);
    }

    /** Writes the preamble, the header and the sheet options of the first sheet. */
    void start(SXSSFSheet first) {
        sheets.add(first);
        this.sheet = first;
        beginSheet(false);
    }

    /** Writes the preamble and header of the current sheet and applies what belongs at the header (frozen panes). */
    private void beginSheet(boolean rolledOver) {
        int preambleRow = ExcelWriteSupport.initSheetPreamble(sheet, wb, columns, cfg.beforeHeaderWriter);
        if (rolledOver) {
            cursor.setRowOfSheet(preambleRow);
        } else {
            cursor = new Cursor(preambleRow);
        }
        ExcelHeaderWriter.write(sheet, cursor, columns, headerStyle, wb, headerStyleCache, cfg.groupComments,
                cfg.headerRowHeightInPoints, cfg.tableOptions != null);
        // With group headers the column names sit below the group rows: that row is the header row.
        int headerRow = cursor.getRowOfSheet() - 1;
        setCurrent(headerRows, headerRow);
        ExcelWriteSupport.applySheetOptions(sheet, headerRow, cfg.freezePaneCols, cfg.freezePaneRows);
    }

    /** The entry of the current sheet in a per-sheet list (the list grows with the sheets). */
    private void setCurrent(List<Integer> list, int value) {
        int index = sheets.size() - 1;
        while (list.size() <= index) list.add(0);
        list.set(index, value);
    }

    /** Writes one row, after turning over to a new sheet when the current one is full. */
    void row(T rowData) {
        cursor.plusTotal();
        if (isOverMaxRows()) {
            StructuredTableWriter.requireRolloverAllowed(cfg.tableOptions);
            closeData();
            setCurrent(nextFreeRows, ExcelWriteSupport.writeAfterDataAndSummary(sheet, wb, cursor.getRowOfSheet(), columns,
                    headerRow(), cfg, values));
            if (values != null) values.resetSheet();
            sheet = nextSheet.apply(sheets.size());
            sheets.add(sheet);
            cursor.initRow();
            beginSheet(true);
        }
        if (subtotals != null) subtotals.beforeRow(sheet, cursor, rowData, cfg, merges, rowStyleCache, wb);
        pageBreaks.beforeRow(sheet, cursor.getRowOfSheet(), rowData);
        ExcelRowWriter.write(sheet, cursor, rowData, columns, cfg, rowStyleCache, wb, merges, subtotals, values);
        ExcelWriteSupport.checkProgress(cursor, cfg.progressInterval, cfg.progressCallback);
    }

    /**
     * Whether the sheet must turn over before the row that was just counted: it has taken {@code maxRows} data rows, or
     * (subtotal rows are no data rows) it has no room left for another row with its subtotal and the summary.
     */
    private boolean isOverMaxRows() {
        long total = cursor.getCurrentTotal();
        // The total already counts the row about to be written: it starts a new sheet when it is the first one past
        // a full sheet. (A "% limit == 1" test never holds for limit 1.)
        if (total > cfg.maxRows && (total - 1) % cfg.maxRows == 0) return true;
        return total > 1 && subtotals != null && subtotals.sheetNearlyFull(cursor);
    }

    /** Ends the open subtotal group and the open merged runs at the end of the data of the current sheet. */
    private void closeData() {
        if (subtotals != null) subtotals.close(sheet, cursor, cfg, merges, rowStyleCache, wb);
        setCurrent(dataEndRows, cursor.getRowOfSheet() - 1);
        if (merges != null) merges.closeRuns(cursor.getRowOfSheet() - 1);
    }

    /**
     * Ends the data of the last sheet: the line of an empty sheet, then the after-data and summary rows.
     *
     * @return the first row below them, where {@code afterAll} content goes
     */
    int finishData() {
        closeData();
        // With no row at all the sheet says so; the line is no data row, so a summary below it still ends at the header.
        int lastDataRow = cursor.getRowOfSheet();
        if (cursor.getCurrentTotal() == 0) {
            ExcelWriteSupport.writeEmptyMessage(sheet, wb, cursor, columns, cfg, rowStyleCache, sharedStyles);
        }
        int nextRow = ExcelWriteSupport.writeAfterDataAndSummary(sheet, wb, cursor.getRowOfSheet(), columns, headerRow(),
                cfg, values, lastDataRow);
        setCurrent(nextFreeRows, nextRow);
        return nextRow;
    }

    /**
     * What needs every sheet written: column widths, validations, outlines, hiding, protection, conditional formatting,
     * print setup and tab color of each sheet, then the tables, then one chart per sheet over that sheet's own rows
     * (a chart of only the last sheet would plot just the remainder).
     */
    void finishSheets() {
        for (int i = 0; i < sheets.size(); i++) {
            ExcelSheetPostProcessor.apply(sheets.get(i), columns, headerRows.get(i), dataEndRows.get(i), cfg, i);
        }
        TableOptions tableOptions = cfg.tableOptions;
        if (tableOptions != null) {
            List<String> headers = columns.stream().map(ExcelColumn::getName).toList();
            for (int i = 0; i < sheets.size(); i++) {
                SXSSFSheet target = sheets.get(i);
                int headerRow = headerRows.get(i);
                if (target.getLastRowNum() <= headerRow) continue;
                String name = sheets.size() == 1 ? tableOptions.name() : tableOptions.name() + "_" + (i + 1);
                StructuredTableWriter.apply(target, name, headerRow, target.getLastRowNum(), headers,
                        tableOptions.style(), tableOptions.showRowStripes(), cfg.autoFilter);
            }
        }
        if (cfg.chartConfig != null) {
            for (int i = 0; i < sheets.size(); i++) {
                ExcelWriteSupport.applyChart(sheets.get(i), cfg.chartConfig, headerRows.get(i), dataEndRows.get(i),
                        nextFreeRows.get(i), columns.size());
            }
        }
    }

    /** The sheet the rows are going to. */
    SXSSFSheet sheet() {
        return sheet;
    }

    /** The header row of the current sheet. */
    int headerRow() {
        return headerRows.get(sheets.size() - 1);
    }

    /** The position the rows are written at. */
    Cursor cursor() {
        return cursor;
    }

}
