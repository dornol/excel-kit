package io.github.dornol.excelkit.excel;

import io.github.dornol.excelkit.shared.Cursor;
import io.github.dornol.excelkit.shared.ProgressCallback;
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
import java.util.Objects;
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

    /**
     * Writes column headers, with optional group header row if any column has a groupName.
     */
    static <T> void writeColumnHeaders(SXSSFSheet sheet, Cursor cursor,
                                        List<ExcelColumn<T>> columns, CellStyle headerStyle) {
        boolean hasGroups = columns.stream().anyMatch(c -> c.getGroupName() != null);
        if (hasGroups) {
            writeGroupAndColumnHeaders(sheet, cursor, columns, headerStyle);
        } else {
            writeSingleHeaderRow(sheet, cursor, columns, headerStyle);
        }
    }

    private static <T> void writeSingleHeaderRow(SXSSFSheet sheet, Cursor cursor,
                                                  List<ExcelColumn<T>> columns, CellStyle headerStyle) {
        SXSSFRow headRow = sheet.createRow(cursor.getRowOfSheet());
        cursor.plusRow();
        for (int j = 0; j < columns.size(); j++) {
            SXSSFCell cell = headRow.createCell(j);
            cell.setCellValue(columns.get(j).getName());
            cell.setCellStyle(headerStyle);
        }
    }

    private static <T> void writeGroupAndColumnHeaders(SXSSFSheet sheet, Cursor cursor,
                                                        List<ExcelColumn<T>> columns, CellStyle headerStyle) {
        int groupRowIdx = cursor.getRowOfSheet();
        SXSSFRow groupRow = sheet.createRow(groupRowIdx);
        cursor.plusRow();
        int columnRowIdx = cursor.getRowOfSheet();
        SXSSFRow columnRow = sheet.createRow(columnRowIdx);
        cursor.plusRow();

        for (int j = 0; j < columns.size(); j++) {
            ExcelColumn<T> col = columns.get(j);
            String group = col.getGroupName();

            // Column header row (always written)
            SXSSFCell colCell = columnRow.createCell(j);
            colCell.setCellValue(col.getName());
            colCell.setCellStyle(headerStyle);

            // Group header row
            SXSSFCell grpCell = groupRow.createCell(j);
            grpCell.setCellStyle(headerStyle);

            if (group != null) {
                grpCell.setCellValue(group);
            }
        }

        // Merge adjacent group cells with the same name
        int i = 0;
        while (i < columns.size()) {
            String group = columns.get(i).getGroupName();
            if (group != null) {
                int start = i;
                while (i < columns.size() && Objects.equals(group, columns.get(i).getGroupName())) {
                    i++;
                }
                if (i - start > 1) {
                    sheet.addMergedRegion(new CellRangeAddress(groupRowIdx, groupRowIdx, start, i - 1));
                }
            } else {
                // No group: merge vertically (group row + column row)
                sheet.addMergedRegion(new CellRangeAddress(groupRowIdx, columnRowIdx, i, i));
                groupRow.getCell(i).setCellValue(columns.get(i).getName());
                i++;
            }
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
                                   Map<String, CellStyle> rowStyleCache, SXSSFWorkbook wb,
                                   int autoWidthSampleRows) {
        SXSSFRow row = sheet.createRow(cursor.getRowOfSheet());
        row.setHeightInPoints(rowHeightInPoints);
        cursor.plusRow();

        ExcelColor rowColor = (rowColorFunction != null) ? rowColorFunction.apply(rowData) : null;

        for (int j = 0; j < columns.size(); j++) {
            SXSSFCell cell = row.createCell(j);
            ExcelColumn<T> column = columns.get(j);
            Object columnData = column.applyFunction(rowData, cursor);
            column.setColumnData(cell, columnData);

            // Resolve effective color: cellColor > rowColor > column default
            ExcelColor effectiveColor = null;
            CellColorFunction<T> cellColorFn = column.getCellColorFunction();
            if (cellColorFn != null) {
                effectiveColor = cellColorFn.apply(columnData, rowData);
            }
            if (effectiveColor == null) {
                effectiveColor = rowColor;
            }

            if (effectiveColor != null) {
                cell.setCellStyle(resolveColorStyle(column.getStyle(), effectiveColor, rowStyleCache, wb));
            } else {
                cell.setCellStyle(column.getStyle());
            }

            if (autoWidthSampleRows > 0 && cursor.getRowOfSheet() < autoWidthSampleRows) {
                column.fitColumnWidthByValue(columnData);
            }
        }
    }

    static CellStyle resolveColorStyle(CellStyle baseStyle, ExcelColor color,
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

    static <T> void applyColumnOutline(SXSSFSheet sheet, List<ExcelColumn<T>> columns) {
        int i = 0;
        while (i < columns.size()) {
            int level = columns.get(i).getOutlineLevel();
            if (level > 0) {
                int start = i;
                while (i < columns.size() && columns.get(i).getOutlineLevel() == level) {
                    i++;
                }
                sheet.groupColumn(start, i - 1);
            } else {
                i++;
            }
        }
    }

    static <T> void validateUniqueColumnNames(List<ExcelColumn<T>> columns) {
        java.util.Set<String> seen = new java.util.HashSet<>();
        for (ExcelColumn<T> col : columns) {
            if (!seen.add(col.getName())) {
                throw new ExcelWriteException("Duplicate column name: '" + col.getName() + "'");
            }
        }
    }

    static void checkProgress(Cursor cursor, int interval, ProgressCallback callback) {
        if (callback != null && interval > 0 && cursor.getCurrentTotal() % interval == 0) {
            callback.onProgress(cursor.getCurrentTotal(), cursor);
        }
    }
}
