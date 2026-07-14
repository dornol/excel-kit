package io.github.dornol.excelkit.excel;

import org.apache.poi.xssf.streaming.SXSSFSheet;
import java.util.List;

/** Applies final sheet features after streaming rows are complete. */
final class ExcelSheetPostProcessor {
    private ExcelSheetPostProcessor() {
    }

    static <T> void apply(SXSSFSheet sheet, List<ExcelColumn<T>> columns,
                          int headerRowIndex, SheetConfig<T> config) {
        ExcelWriteSupport.applyColumnWidths(sheet, columns);
        ExcelWriteSupport.applyDataValidations(sheet, columns, headerRowIndex);
        ExcelWriteSupport.applyColumnOutline(sheet, columns);
        ExcelWriteSupport.applyColumnHidden(sheet, columns);
        ExcelWriteSupport.applySheetProtection(sheet, config.sheetPassword);
        ExcelWriteSupport.applyConditionalFormatting(sheet, config.conditionalRules, headerRowIndex,
                columns.size(), sheet.getLastRowNum());
        ExcelWriteSupport.applyPrintSetup(sheet, config.printSetup, headerRowIndex);
        ExcelWriteSupport.applyTabColor(sheet, config.tabColor);
        ExcelWriteSupport.applyNamedRanges(sheet, config.namedRanges, headerRowIndex);
    }
}
