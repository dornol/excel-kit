package io.github.dornol.excelkit.excel;

import io.github.dornol.excelkit.core.Cursor;
import org.apache.poi.xssf.streaming.SXSSFSheet;

import org.jspecify.annotations.Nullable;

import java.util.LinkedHashMap;
import java.util.Map;

/** Mutable state scoped to one writer execution. */
final class ExcelWriteSession<T> {
    private final ExcelWriteOptions<T> options;
    private final Map<SXSSFSheet, Integer> headerRows = new LinkedHashMap<>();
    private final Map<SXSSFSheet, Integer> dataEndRows = new LinkedHashMap<>();
    private final Map<SXSSFSheet, Integer> nextFreeRows = new LinkedHashMap<>();
    private final PageBreakTracker<T> pageBreaks;
    private final @Nullable MergeRunTracker<T> merges;
    private final @Nullable SubtotalTracker<T> subtotals;
    private final @Nullable FormulaValues values;
    private SXSSFSheet sheet;
    private Cursor cursor;

    ExcelWriteSession(ExcelWriteOptions<T> options) {
        this.options = options;
        this.pageBreaks = new PageBreakTracker<>(options.sheetConfig().pageBreakWhen);
        this.merges = MergeRunTracker.of(options.columns());
        this.values = FormulaValues.of(options.sheetConfig(), options.columns());
        this.subtotals = SubtotalTracker.of(options.sheetConfig(), options.columns(), values);
    }

    ExcelWriteOptions<T> options() { return options; }
    PageBreakTracker<T> pageBreaks() { return pageBreaks; }
    @Nullable MergeRunTracker<T> merges() { return merges; }
    @Nullable SubtotalTracker<T> subtotals() { return subtotals; }
    @Nullable FormulaValues values() { return values; }
    SXSSFSheet sheet() { return sheet; }
    void sheet(SXSSFSheet sheet) { this.sheet = sheet; }
    Cursor cursor() { return cursor; }
    void cursor(Cursor cursor) { this.cursor = cursor; }
    void headerRow(int row) { headerRows.put(sheet, row); }
    int headerRow() { return headerRows.get(sheet); }
    int headerRow(SXSSFSheet target) { return headerRows.get(target); }
    void dataEndRow(int row) { dataEndRows.put(sheet, row); }
    int dataEndRow(SXSSFSheet target) { return dataEndRows.getOrDefault(target, target.getLastRowNum()); }
    /** The first row below the after-data and summary rows of the current sheet, where a chart may be placed. */
    void nextFreeRow(int row) { nextFreeRows.put(sheet, row); }
    int nextFreeRow(SXSSFSheet target) { return nextFreeRows.getOrDefault(target, dataEndRow(target) + 1); }
    Map<SXSSFSheet, Integer> headerRows() { return Map.copyOf(headerRows); }
}
