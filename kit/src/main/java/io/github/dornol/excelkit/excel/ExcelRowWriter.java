package io.github.dornol.excelkit.excel;

import org.apache.poi.ss.usermodel.*;

import io.github.dornol.excelkit.excel.style.ExcelColor;
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
                          SheetConfig<T> config, Map<ExcelStyleSupporter.DerivedStyleKey, CellStyle> styleCache, SXSSFWorkbook workbook) {
        write(sheet, cursor, row, columns, config, styleCache, workbook, null, null, null);
    }

    /**
     * @param merges    merges the runs of equal values of the columns that asked for it, or {@code null}
     * @param values    the results of the formulas that cover the data cells, or {@code null}
     * @param subtotals the subtotal rows of the sheet, or {@code null}: they say which data row of the sheet this is,
     *                  because a sheet can turn over before {@code maxRows} when the subtotal rows take its room
     */
    static <T> void write(SXSSFSheet sheet, Cursor cursor, T row, List<ExcelColumn<T>> columns,
                          SheetConfig<T> config, Map<ExcelStyleSupporter.DerivedStyleKey, CellStyle> styleCache, SXSSFWorkbook workbook,
                          @Nullable MergeRunTracker<T> merges, @Nullable SubtotalTracker<T> subtotals,
                          @Nullable FormulaValues values) {
        if (cursor.getRowOfSheet() >= ExcelLimits.SHEET_ROW_CAPACITY) {
            throw new ExcelWriteException("Sheet '" + sheet.getSheetName() + "' is full: Excel sheets hold at most "
                    + ExcelLimits.SHEET_ROW_CAPACITY + " rows. Lower maxRows(...) so the writer rolls over to a "
                    + "new sheet earlier; rows written before the data (beforeHeader, header) and after it "
                    + "(afterData, summary) count towards the limit.");
        }
        SXSSFRow target = sheet.createRow(cursor.getRowOfSheet());
        target.setHeightInPoints(config.rowHeightInPoints);
        if (merges != null) merges.startRow(sheet, target.getRowNum());
        cursor.plusRow();

        ExcelColor rowColor = config.rowColorFunction == null ? null : config.rowColorFunction.apply(row);
        // every second data row of the sheet: the total counts rows over all sheets, a sheet takes maxRows of them
        long dataIndexInSheet = subtotals != null ? subtotals.dataIndex() : (cursor.getCurrentTotal() - 1) % config.maxRows;
        if (rowColor == null && config.stripeColor != null && dataIndexInSheet % 2 == 1) {
            rowColor = config.stripeColor;
        }
        @Nullable RowStyleConfig rowStyle = matchingStyle(row, config.rowStyleEntries);
        for (int i = 0; i < columns.size(); i++) {
            SXSSFCell cell = target.createCell(i);
            ExcelColumn<T> column = columns.get(i);
            @Nullable Object value = column.applyFunction(row, cursor, config.writeErrorPolicy);
            if (merges != null && merges.ordinal(i) >= 0) merges.value(merges.ordinal(i), value);
            CellStyle baseStyle = column.setColumnData(cell, value, config.writeErrorPolicy);
            if (values != null) values.cell(i, cell);

            ExcelColor color = effectiveColor(column, value, row, rowStyle, rowColor);
            if (rowStyle != null && rowStyle.hasAnyStyle())
                cell.setCellStyle(ExcelStyleSupporter.resolveRowStyle(baseStyle, color, rowStyle,
                        styleCache, workbook, config.maxCellStyles));
            else if (color != null)
                cell.setCellStyle(ExcelStyleSupporter.resolveColorStyle(baseStyle, color, styleCache, workbook,
                        config.maxCellStyles));
            else cell.setCellStyle(baseStyle);

            if (config.autoWidthSampleRows > 0 && cursor.getCurrentTotal() <= config.autoWidthSampleRows)
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
        if (text != null) ExcelRowWriter.addCellComment(cell, text, null,
                column.getCommentWidth(), column.getCommentHeight(), workbook);
    }

    static void addCellComment(SXSSFCell cell, String text, SXSSFWorkbook wb) {
        addCellComment(cell, text, null, 0, 0, wb);
    }

    static void addCellComment(SXSSFCell cell, String text, @Nullable String author,
                                int width, int height, SXSSFWorkbook wb) {
        Drawing<?> drawing = cell.getSheet().createDrawingPatriarch();
        CreationHelper factory = wb.getCreationHelper();
        ClientAnchor anchor = factory.createClientAnchor();
        int w = width > 0 ? width : 2;
        int h = height > 0 ? height : 3;
        anchor.setCol1(cell.getColumnIndex());
        anchor.setCol2(cell.getColumnIndex() + w);
        anchor.setRow1(cell.getRowIndex());
        anchor.setRow2(cell.getRowIndex() + h);
        Comment comment = drawing.createCellComment(anchor);
        comment.setString(factory.createRichTextString(text));
        if (author != null) {
            comment.setAuthor(author);
        }
        cell.setCellComment(comment);
    }
}
