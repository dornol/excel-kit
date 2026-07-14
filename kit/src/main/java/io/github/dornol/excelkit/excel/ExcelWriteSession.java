package io.github.dornol.excelkit.excel;

import io.github.dornol.excelkit.core.Cursor;
import org.apache.poi.xssf.streaming.SXSSFSheet;

import java.util.LinkedHashMap;
import java.util.Map;

/** Mutable state scoped to one writer execution. */
final class ExcelWriteSession<T> {
    private final ExcelWriteOptions<T> options;
    private final Map<SXSSFSheet, Integer> headerRows = new LinkedHashMap<>();
    private SXSSFSheet sheet;
    private Cursor cursor;

    ExcelWriteSession(ExcelWriteOptions<T> options) {
        this.options = options;
    }

    ExcelWriteOptions<T> options() { return options; }
    SXSSFSheet sheet() { return sheet; }
    void sheet(SXSSFSheet sheet) { this.sheet = sheet; }
    Cursor cursor() { return cursor; }
    void cursor(Cursor cursor) { this.cursor = cursor; }
    void headerRow(int row) { headerRows.put(sheet, row); }
    int headerRow() { return headerRows.get(sheet); }
    int headerRow(SXSSFSheet target) { return headerRows.get(target); }
    Map<SXSSFSheet, Integer> headerRows() { return Map.copyOf(headerRows); }
}
