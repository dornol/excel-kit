package io.github.dornol.excelkit.excel;

import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.ss.usermodel.DataValidation;
import org.apache.poi.ss.usermodel.DataValidationConstraint;
import org.apache.poi.ss.usermodel.DataValidationHelper;
import org.apache.poi.ss.util.CellRangeAddressList;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTCols;
import java.util.List;

/** Applies final sheet features after streaming rows are complete. */
final class ExcelSheetPostProcessor {
    private ExcelSheetPostProcessor() {
    }

    /**
     * @param headerRowIndex the row holding the column names (the last row of a multi-row header)
     * @param dataEndRow the last data row of this sheet; conditional formatting stops there so that
     *                   summary and {@code afterData} rows are not formatted (or scaled) as data
     * @param sheetIndex 0 for the first sheet of a write, 1 for its first rollover sheet, ...
     */
    static <T> void apply(SXSSFSheet sheet, List<ExcelColumn<T>> columns,
                          int headerRowIndex, int dataEndRow, SheetConfig<T> config, int sheetIndex) {
        applyColumnWidths(sheet, columns);
        applyDataValidations(sheet, columns, headerRowIndex);
        applyColumnOutline(sheet, columns);
        applyColumnHidden(sheet, columns);
        // The filter covers the header and the data rows (not the summary below them), as Excel writes it. A table brings
        // its own filter buttons: see StructuredTableWriter, where a sheet-level filter would overlap the table.
        if (config.autoFilter && config.tableOptions == null) {
            sheet.setAutoFilter(new CellRangeAddress(headerRowIndex,
                    Math.max(headerRowIndex, dataEndRow), 0, columns.size() - 1));
        }
        if (config.sheetPassword != null) sheet.protectSheet(config.sheetPassword);
        if (config.conditionalRules != null) for (ExcelConditionalRule rule : config.conditionalRules)
            rule.apply(sheet, headerRowIndex, columns.size(), dataEndRow);
        if (config.printSetup != null) config.printSetup.apply(sheet, headerRowIndex, dataEndRow);
        requirePageBreaksWithinLimit(sheet);
        ExcelWriteSupport.applyTabColor(sheet, config.tabColor);
        if (config.zoom > 0) sheet.setZoom(config.zoom);
        ExcelWriteSupport.applyVisibility(sheet, config.visibility);
        ExcelWriteSupport.applyNamedRanges(sheet, config.namedRanges, headerRowIndex, sheetIndex);
    }

    /** Excel keeps at most 1026 horizontal page breaks per sheet and repairs a file that holds more. */
    private static final int MAX_ROW_BREAKS = 1026;

    private static void requirePageBreaksWithinLimit(SXSSFSheet sheet) {
        int breaks = sheet.getRowBreaks().length;
        if (breaks > MAX_ROW_BREAKS) {
            throw new ExcelWriteException("Sheet '" + sheet.getSheetName() + "' has " + breaks
                    + " page breaks; Excel keeps at most " + MAX_ROW_BREAKS + ". Use fewer breaks (a larger rowBreakEvery, "
                    + "or a pageBreakWhen that holds less often)");
        }
    }

    /**
     * The widths of a list written into a template. A template is somebody's layout, so a column width it sets
     * stays; only a column given {@code width(n)}, or one the template left at the default width, is set here.
     */
    static <T> void applyTemplateColumnWidths(SXSSFSheet sheet, List<ExcelColumn<T>> columns) {
        var templateColumns = SXSSFSheetHelper.getXSSFSheet(sheet).getCTWorksheet().getColsArray();
        for (int i = 0; i < columns.size(); i++) {
            ExcelColumn<T> column = columns.get(i);
            if (column.isFixedWidth() || !templateSetsWidth(templateColumns, i)) {
                sheet.setColumnWidth(i, column.getColumnWidth());
            }
        }
    }

    /** Whether the template gave the column (0-based) a width of its own: a {@code <col>} with customWidth. */
    private static boolean templateSetsWidth(
            CTCols[] groups, int column) {
        for (var group : groups) {
            for (var col : group.getColArray()) {
                if (col.getCustomWidth() && col.getMin() <= column + 1 && column + 1 <= col.getMax()) return true;
            }
        }
        return false;
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
                        new CellRangeAddressList(header + 1, ExcelLimits.EXCEL_MAX_ROWS, i, i));
                validation.setSuppressDropDownArrow(false);
                validation.setShowErrorBox(true);
                sheet.addValidationData(validation);
            }
            if (column.getValidation() != null) column.getValidation().apply(helper, sheet, i, header);
            DependentDropdown dependent = column.getDependentDropdown();
            if (dependent != null) dependent.apply(sheet, i, parentIndex(columns, i, dependent.parentColumn()), header);
        }
    }

    private static <T> int parentIndex(List<ExcelColumn<T>> columns, int self, String parentColumn) {
        for (int j = 0; j < columns.size(); j++) {
            if (j != self && columns.get(j).getName().equals(parentColumn)) return j;
        }
        throw new ExcelWriteException("The dependent dropdown of column '" + columns.get(self).getName()
                + "' names the parent column '" + parentColumn + "', which is not another column of the sheet");
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
