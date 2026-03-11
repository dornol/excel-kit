package io.github.dornol.excelkit.excel;

import io.github.dornol.excelkit.shared.Cursor;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DataValidation;
import org.apache.poi.ss.usermodel.DataValidationConstraint;
import org.apache.poi.ss.usermodel.DataValidationHelper;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellRangeAddressList;
import org.apache.poi.xssf.streaming.SXSSFCell;
import org.apache.poi.xssf.streaming.SXSSFRow;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFColor;

import java.util.List;
import java.util.Map;
import java.util.function.Function;

/**
 * Package-private utility methods shared by {@link ExcelWriter} and {@link ExcelSheetWriter}
 * to eliminate duplicate write logic.
 *
 * @author dhkim
 */
class ExcelWriteSupport {

    static final int AUTO_WIDTH_SAMPLE_ROWS = 100;
    static final int EXCEL_MAX_ROWS = 1_048_575;

    private ExcelWriteSupport() {
    }

    static <T> void writeColumnHeaders(SXSSFSheet sheet, Cursor cursor,
                                        List<ExcelColumn<T>> columns, CellStyle headerStyle) {
        SXSSFRow headRow = sheet.createRow(cursor.getRowOfSheet());
        cursor.plusRow();
        for (int j = 0; j < columns.size(); j++) {
            SXSSFCell cell = headRow.createCell(j);
            cell.setCellValue(columns.get(j).getName());
            cell.setCellStyle(headerStyle);
        }
    }

    static void applySheetOptions(SXSSFSheet sheet, int headerRowIdx,
                                   boolean autoFilter, int freezePaneRows, int columnCount) {
        if (autoFilter) {
            sheet.setAutoFilter(new CellRangeAddress(headerRowIdx, headerRowIdx, 0, columnCount - 1));
        }
        if (freezePaneRows > 0) {
            sheet.createFreezePane(0, headerRowIdx + freezePaneRows);
        }
    }

    static <T> void writeRowCells(SXSSFSheet sheet, Cursor cursor, T rowData,
                                   List<ExcelColumn<T>> columns, float rowHeightInPoints,
                                   Function<T, ExcelColor> rowColorFunction,
                                   Map<String, CellStyle> rowStyleCache, SXSSFWorkbook wb) {
        SXSSFRow row = sheet.createRow(cursor.getRowOfSheet());
        row.setHeightInPoints(rowHeightInPoints);
        cursor.plusRow();

        ExcelColor rowColor = (rowColorFunction != null) ? rowColorFunction.apply(rowData) : null;

        for (int j = 0; j < columns.size(); j++) {
            SXSSFCell cell = row.createCell(j);
            ExcelColumn<T> column = columns.get(j);
            Object columnData = column.applyFunction(rowData, cursor);
            column.setColumnData(cell, columnData);
            if (rowColor != null) {
                cell.setCellStyle(resolveRowColorStyle(column.getStyle(), rowColor, rowStyleCache, wb));
            } else {
                cell.setCellStyle(column.getStyle());
            }
            if (cursor.getRowOfSheet() < AUTO_WIDTH_SAMPLE_ROWS) {
                column.fitColumnWidthByValue(columnData);
            }
        }
    }

    static CellStyle resolveRowColorStyle(CellStyle baseStyle, ExcelColor color,
                                           Map<String, CellStyle> cache, SXSSFWorkbook wb) {
        String key = baseStyle.getIndex() + "_" + color.getR() + "_" + color.getG() + "_" + color.getB();
        return cache.computeIfAbsent(key, k -> {
            CellStyle style = wb.createCellStyle();
            style.cloneStyleFrom(baseStyle);
            style.setFillForegroundColor(new XSSFColor(new byte[]{
                    (byte) color.getR(), (byte) color.getG(), (byte) color.getB()}));
            style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            return style;
        });
    }

    static <T> void applyDataValidations(SXSSFSheet sheet, List<ExcelColumn<T>> columns,
                                          int headerRowIndex) {
        DataValidationHelper helper = sheet.getDataValidationHelper();
        for (int j = 0; j < columns.size(); j++) {
            String[] options = columns.get(j).getDropdownOptions();
            if (options != null) {
                DataValidationConstraint constraint = helper.createExplicitListConstraint(options);
                CellRangeAddressList range = new CellRangeAddressList(
                        headerRowIndex + 1, EXCEL_MAX_ROWS, j, j);
                DataValidation validation = helper.createValidation(constraint, range);
                validation.setSuppressDropDownArrow(false);
                validation.setShowErrorBox(true);
                sheet.addValidationData(validation);
            }
        }
    }

    static <T> void applyColumnWidths(SXSSFSheet sheet, List<ExcelColumn<T>> columns) {
        for (int j = 0; j < columns.size(); j++) {
            sheet.setColumnWidth(j, columns.get(j).getColumnWidth());
        }
    }

    static <T> int initSheetPreamble(SXSSFSheet sheet, SXSSFWorkbook wb,
                                      List<ExcelColumn<T>> columns,
                                      BeforeHeaderWriter writer) {
        int currentRow = 0;
        if (writer != null) {
            currentRow = writer.write(new SheetContext(sheet, wb, currentRow, columns));
        }
        return currentRow;
    }
}
