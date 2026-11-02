package io.github.dornol.excelkit.excel;

import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFSheet;

/** Resolves the backing sheet through public workbook APIs, including after rename/reorder. */
final class SXSSFSheetHelper {
    private SXSSFSheetHelper() { }

    /**
     * Resolves the registered streaming sheet by identity, not by its mutable name.
     * Removed or otherwise unregistered sheets must not resolve to a replacement
     * sheet that happens to have the same name.
     *
     * @throws ExcelWriteException if the sheet cannot be resolved
     */
    static XSSFSheet getXSSFSheet(SXSSFSheet sheet) {
        try {
            var workbook = sheet.getWorkbook();
            int index = workbook.getSheetIndex(sheet);
            if (index < 0) {
                throw new ExcelWriteException("Cannot resolve backing XSSFSheet: sheet is not registered in its streaming workbook");
            }
            return workbook.getXSSFWorkbook().getSheetAt(index);
        } catch (ExcelWriteException e) {
            throw e;
        } catch (RuntimeException e) {
            throw new ExcelWriteException("Cannot resolve backing XSSFSheet through public workbook APIs", e);
        }
    }
}
