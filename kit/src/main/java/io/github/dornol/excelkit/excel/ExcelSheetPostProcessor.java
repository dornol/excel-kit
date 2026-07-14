package io.github.dornol.excelkit.excel;

import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.ss.usermodel.DataValidation;
import org.apache.poi.ss.usermodel.DataValidationConstraint;
import org.apache.poi.ss.usermodel.DataValidationHelper;
import org.apache.poi.ss.util.CellRangeAddressList;
import java.util.List;

/** Applies final sheet features after streaming rows are complete. */
final class ExcelSheetPostProcessor {
    private ExcelSheetPostProcessor() {
    }

    static <T> void apply(SXSSFSheet sheet, List<ExcelColumn<T>> columns,
                          int headerRowIndex, SheetConfig<T> config) {
        applyColumnWidths(sheet, columns);
        applyDataValidations(sheet, columns, headerRowIndex);
        applyColumnOutline(sheet, columns);
        applyColumnHidden(sheet, columns);
        if (config.sheetPassword != null) sheet.protectSheet(config.sheetPassword);
        if (config.conditionalRules != null) for (ExcelConditionalRule rule : config.conditionalRules)
            rule.apply(sheet, headerRowIndex, columns.size(), sheet.getLastRowNum());
        if (config.printSetup != null) config.printSetup.apply(sheet, headerRowIndex);
        ExcelWriteSupport.applyTabColor(sheet, config.tabColor);
        ExcelWriteSupport.applyNamedRanges(sheet, config.namedRanges, headerRowIndex);
    }

    static <T> void applyColumnWidths(SXSSFSheet sheet, List<ExcelColumn<T>> columns) {
        for (int i = 0; i < columns.size(); i++) sheet.setColumnWidth(i, columns.get(i).getColumnWidth());
    }

    private static <T> void applyDataValidations(SXSSFSheet sheet, List<ExcelColumn<T>> columns, int header) {
        DataValidationHelper helper = sheet.getDataValidationHelper();
        for (int i = 0; i < columns.size(); i++) {
            ExcelColumn<T> column = columns.get(i);
            if (column.getDropdownOptions() != null) {
                DataValidationConstraint constraint = helper.createExplicitListConstraint(column.getDropdownOptions());
                DataValidation validation = helper.createValidation(constraint,
                        new CellRangeAddressList(header + 1, ExcelWriteSupport.EXCEL_MAX_ROWS, i, i));
                validation.setSuppressDropDownArrow(false);
                validation.setShowErrorBox(true);
                sheet.addValidationData(validation);
            }
            if (column.getValidation() != null) column.getValidation().apply(helper, sheet, i, header);
        }
    }

    private static <T> void applyColumnHidden(SXSSFSheet sheet, List<ExcelColumn<T>> columns) {
        for (int i = 0; i < columns.size(); i++) if (columns.get(i).isHidden()) sheet.setColumnHidden(i, true);
    }

    private static <T> void applyColumnOutline(SXSSFSheet sheet, List<ExcelColumn<T>> columns) {
        int i = 0;
        while (i < columns.size()) {
            int level = columns.get(i).getOutlineLevel();
            if (level <= 0) { i++; continue; }
            int start = i;
            while (i < columns.size() && columns.get(i).getOutlineLevel() == level) i++;
            sheet.groupColumn(start, i - 1);
        }
    }
}
