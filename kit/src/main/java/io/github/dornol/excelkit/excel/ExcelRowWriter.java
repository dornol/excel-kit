package io.github.dornol.excelkit.excel;

import io.github.dornol.excelkit.core.Cursor;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.streaming.SXSSFCell;
import org.apache.poi.xssf.streaming.SXSSFRow;
import org.jspecify.annotations.Nullable;
import java.util.List;
import java.util.Map;
import java.util.function.Function;

/** Row-writing entry point shared by writer variants. */
final class ExcelRowWriter {
    private ExcelRowWriter() { }

    static <T> void write(SXSSFSheet sheet, Cursor cursor, T row, List<ExcelColumn<T>> columns,
                          SheetConfig<T> config, Map<String, CellStyle> styleCache, SXSSFWorkbook workbook) {
        SXSSFRow target = sheet.createRow(cursor.getRowOfSheet());
        target.setHeightInPoints(config.rowHeightInPoints);
        cursor.plusRow();

        ExcelColor rowColor = config.rowColorFunction == null ? null : config.rowColorFunction.apply(row);
        @Nullable RowStyleConfig rowStyle = matchingStyle(row, config.rowStyleEntries);
        for (int i = 0; i < columns.size(); i++) {
            SXSSFCell cell = target.createCell(i);
            ExcelColumn<T> column = columns.get(i);
            @Nullable Object value = column.applyFunction(row, cursor, config.writeErrorPolicy);
            column.setColumnData(cell, value, config.writeErrorPolicy);

            ExcelColor color = effectiveColor(column, value, row, rowStyle, rowColor);
            if (rowStyle != null && rowStyle.hasAnyStyle())
                cell.setCellStyle(ExcelWriteSupport.resolveRowStyle(column.getStyle(), color, rowStyle,
                        styleCache, workbook));
            else if (color != null)
                cell.setCellStyle(ExcelWriteSupport.resolveColorStyle(column.getStyle(), color, styleCache, workbook));
            else cell.setCellStyle(column.getStyle());

            if (config.autoWidthSampleRows > 0 && cursor.getRowOfSheet() < config.autoWidthSampleRows)
                column.fitColumnWidthByValue(value);
            addComment(cell, column, row, workbook);
        }
    }

    private static <T> @Nullable RowStyleConfig matchingStyle(T row,
            List<SheetConfig.RowStyleEntry<T>> entries) {
        for (SheetConfig.RowStyleEntry<T> entry : entries)
            if (entry.predicate().test(row)) return entry.style();
        return null;
    }

    private static <T> @Nullable ExcelColor effectiveColor(ExcelColumn<T> column,
            @Nullable Object value, T row, @Nullable RowStyleConfig style, @Nullable ExcelColor rowColor) {
        ExcelColor color = column.getCellColorFunction() == null ? null
                : column.getCellColorFunction().apply(value, row);
        if (color == null && style != null) color = style.backgroundColor;
        return color == null ? rowColor : color;
    }

    private static <T> void addComment(SXSSFCell cell, ExcelColumn<T> column, T row,
                                        SXSSFWorkbook workbook) {
        Function<T, @Nullable String> function = column.getCommentFunction();
        if (function == null) return;
        String text = function.apply(row);
        if (text != null) ExcelWriteSupport.addCellComment(cell, text, null,
                column.getCommentWidth(), column.getCommentHeight(), workbook);
    }
}
