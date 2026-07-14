package io.github.dornol.excelkit.excel;

import io.github.dornol.excelkit.core.Cursor;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.jspecify.annotations.Nullable;
import java.util.List;
import java.util.Map;

/** Header-writing entry point shared by writer variants. */
final class ExcelHeaderWriter {
    private ExcelHeaderWriter() { }

    static <T> void write(SXSSFSheet sheet, Cursor cursor, List<ExcelColumn<T>> columns,
                          CellStyle style, @Nullable SXSSFWorkbook workbook,
                          @Nullable Map<String, CellStyle> styleCache,
                          @Nullable Map<List<String>, ExcelCellComment> groupComments, float rowHeight) {
        ExcelWriteSupport.writeColumnHeaders(sheet, cursor, columns, style, workbook,
                styleCache, groupComments, rowHeight);
    }

    static <T> void write(SXSSFSheet sheet, Cursor cursor, List<ExcelColumn<T>> columns, CellStyle style) {
        ExcelWriteSupport.writeColumnHeaders(sheet, cursor, columns, style);
    }
}
