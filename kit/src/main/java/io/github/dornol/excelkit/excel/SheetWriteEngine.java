package io.github.dornol.excelkit.excel;

import io.github.dornol.excelkit.core.Cursor;
import io.github.dornol.excelkit.core.ProgressCallback;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.streaming.SXSSFCell;
import org.apache.poi.xssf.streaming.SXSSFRow;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.jspecify.annotations.Nullable;

import java.util.ArrayList;
import java.util.HashSet;
import java.util.List;
import java.util.Map;
import java.util.Set;
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
    private static final String EMPTY_MESSAGE_STYLE = "excel-kit:empty-message";

    private final SXSSFWorkbook wb;
    private final List<ExcelColumn<T>> columns;
    private final SheetConfig<T> cfg;
    private final CellStyle headerStyle;
    private final Map<String, CellStyle> headerStyleCache;
    private final Map<ExcelStyleSupporter.DerivedStyleKey, CellStyle> rowStyleCache;
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
                     Map<String, CellStyle> headerStyleCache, Map<ExcelStyleSupporter.DerivedStyleKey, CellStyle> rowStyleCache,
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
        SheetWriteEngine.validateUniqueColumnNames(columns);
        SheetWriteEngine.validateMergeColumns(columns, cfg.tableOptions, true);
        SubtotalTracker.validate(cfg);
        SheetWriteEngine.validateEmptyMessage(cfg);
    }

    /** Writes the preamble, the header and the sheet options of the first sheet. */
    void start(SXSSFSheet first) {
        sheets.add(first);
        this.sheet = first;
        beginSheet(false);
    }

    /** Writes the preamble and header of the current sheet and applies what belongs at the header (frozen panes). */
    private void beginSheet(boolean rolledOver) {
        int preambleRow = cfg.beforeHeaderWriter == null ? 0
                : cfg.beforeHeaderWriter.write(new SheetContext(sheet, wb, 0, columns));
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
        ExcelSheetPostProcessor.applySheetOptions(sheet, headerRow, cfg.freezePaneCols, cfg.freezePaneRows);
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
            setCurrent(nextFreeRows, SheetWriteEngine.writeAfterDataAndSummary(sheet, wb, cursor.getRowOfSheet(), columns,
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
        SheetWriteEngine.checkProgress(cursor, cfg.progressInterval, cfg.progressCallback);
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
            SheetWriteEngine.writeEmptyMessage(sheet, wb, cursor, columns, cfg, rowStyleCache, sharedStyles);
        }
        int nextRow = SheetWriteEngine.writeAfterDataAndSummary(sheet, wb, cursor.getRowOfSheet(), columns, headerRow(),
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
            SXSSFSheet target = sheets.get(i);
            ExcelSheetPostProcessor.apply(target, columns, headerRows.get(i), dataEndRows.get(i), cfg, i,
                    subtotals == null ? new int[0] : subtotals.subtotalRows(target));
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
                cfg.chartConfig.apply(sheets.get(i), headerRows.get(i), dataEndRows.get(i),
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

    /** @param values what the summary formulas calculate for the sheet, or {@code null} when they are not tracked */
    static <T> int writeAfterDataAndSummary(SXSSFSheet sheet, SXSSFWorkbook wb, int startRow,
                                             List<ExcelColumn<T>> columns, int headerRowIndex,
                                             SheetConfig<T> cfg, @Nullable FormulaValues values) {
        return writeAfterDataAndSummary(sheet, wb, startRow, columns, headerRowIndex, cfg, values, startRow);
    }

    /**
     * @param dataEndRow the 1-based last data row, which is where the summary ranges end: {@code startRow} when the
     *                   rows below the data are the next ones, less when something that is not data (the empty
     *                   message) sits between
     */
    static <T> int writeAfterDataAndSummary(SXSSFSheet sheet, SXSSFWorkbook wb, int startRow,
                                             List<ExcelColumn<T>> columns, int headerRowIndex,
                                             SheetConfig<T> cfg, @Nullable FormulaValues values, int dataEndRow) {
        int row = startRow;
        if (cfg.afterDataWriter != null) {
            row = cfg.afterDataWriter.write(new SheetContext(sheet, wb, row, columns, headerRowIndex));
        }
        if (cfg.summaryConfig != null) {
            // The summary covers the data rows only, not whatever afterData wrote below them.
            row = cfg.summaryConfig.toAfterDataWriter(cfg.subtotalKey != null, values).write(
                    new SheetContext(sheet, wb, row, columns, headerRowIndex, dataEndRow));
        }
        return row;
    }

    /** Rejects what the empty message would corrupt: a merged line has no place in a table. */
    static <T> void validateEmptyMessage(SheetConfig<T> cfg) {
        if (cfg.emptyMessage != null && cfg.tableOptions != null) {
            throw new ExcelWriteException("emptyMessage(...) cannot be combined with table(...): the message is a merged "
                    + "cell, and Excel does not allow merged cells in a table");
        }
    }

    /**
     * Writes the line of an empty sheet below the header: the text centered over all the columns, in gray italics.
     * It takes a row of the sheet but is no data row.
     */
    static <T> void writeEmptyMessage(SXSSFSheet sheet, SXSSFWorkbook wb, Cursor cursor, List<ExcelColumn<T>> columns,
                                      SheetConfig<T> cfg, Map<ExcelStyleSupporter.DerivedStyleKey, CellStyle> styleCache,
                                      @Nullable Map<String, CellStyle> workbookStyles) {
        String text = cfg.emptyMessage;
        if (text == null) return;
        int rowIndex = cursor.getRowOfSheet();
        if (rowIndex >= ExcelLimits.SHEET_ROW_CAPACITY) {
            throw new ExcelWriteException("Sheet '" + sheet.getSheetName() + "' is full: no room for the empty message");
        }
        // The style does not depend on the columns, so the sheets of a workbook share one.
        CellStyle style = workbookStyles != null
                ? workbookStyles.computeIfAbsent(EMPTY_MESSAGE_STYLE, k -> ExcelStyleSupporter.emptyMessageStyle(wb, cfg))
                : styleCache.computeIfAbsent(new ExcelStyleSupporter.DerivedStyleKey(-1, -1, "empty-message"), k -> ExcelStyleSupporter.emptyMessageStyle(wb, cfg));
        SXSSFRow row = sheet.createRow(rowIndex);
        row.setHeightInPoints(cfg.rowHeightInPoints);
        cursor.plusRow();
        for (int i = 0; i < columns.size(); i++) {
            SXSSFCell cell = row.createCell(i);
            if (i == 0) ExcelCellText.write(cell, text);
            cell.setCellStyle(style);
        }
        if (columns.size() > 1) sheet.addMergedRegion(new CellRangeAddress(rowIndex, rowIndex, 0, columns.size() - 1));
    }

    /**
     * Rejects a column that merges equal values where the sheet cannot hold merged cells: in a table (Excel repairs
     * the file) or in a list written into a template.
     */
    static <T> void validateMergeColumns(List<ExcelColumn<T>> columns, @Nullable TableOptions tableOptions, boolean supported) {
        for (ExcelColumn<T> col : columns) {
            if (!col.isMergeSameValues()) continue;
            if (!supported) {
                throw new ExcelWriteException("Column '" + col.getName() + "' asks for mergeSameValues(), "
                        + "which a template list writer does not apply");
            }
            if (tableOptions != null) {
                throw new ExcelWriteException("Column '" + col.getName() + "' asks for mergeSameValues(), but "
                        + "the sheet is a table(...): Excel does not allow merged cells in a table");
            }
        }
    }

    static <T> void validateUniqueColumnNames(List<ExcelColumn<T>> columns) {
        Set<String> seen = new HashSet<>();
        for (ExcelColumn<T> col : columns) {
            if (!seen.add(col.getName())) {
                throw new ExcelWriteException("Duplicate column name: '" + col.getName() + "'");
            }
        }
    }

    static void checkProgress(Cursor cursor, int interval, @Nullable ProgressCallback callback) {
        if (callback != null && interval > 0 && cursor.getCurrentTotal() % interval == 0) {
            callback.onProgress(cursor.getCurrentTotal(), cursor);
        }
    }
}
