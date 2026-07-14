package io.github.dornol.excelkit.excel;

import io.github.dornol.excelkit.core.Cursor;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import java.util.List;
import java.util.Map;

/** Row-writing entry point shared by writer variants. */
final class ExcelRowWriter {
    private ExcelRowWriter() { }

    static <T> void write(SXSSFSheet sheet, Cursor cursor, T row, List<ExcelColumn<T>> columns,
                          SheetConfig<T> config, Map<String, CellStyle> styleCache, SXSSFWorkbook workbook) {
        ExcelWriteSupport.writeRowCells(sheet, cursor, row, columns, config, styleCache, workbook);
    }
}
