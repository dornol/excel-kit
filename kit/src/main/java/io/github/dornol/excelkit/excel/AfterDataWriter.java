package io.github.dornol.excelkit.excel;

import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

/**
 * A functional interface for writing custom content after data rows.
 * <p>
 * Used by both {@code afterData} (called on every sheet after its data rows)
 * and {@code afterAll} (called once on the last sheet after all data).
 *
 * @author dhkim
 * @since 2025-07-19
 */
@FunctionalInterface
public interface AfterDataWriter {

    /**
     * Writes custom content to the sheet after the data rows.
     *
     * @param sheet    the current sheet
     * @param workbook the workbook (useful for creating CellStyles, etc.)
     * @param nextRow  the first row index available for writing (after the last data row)
     * @param context  column metadata (count, names) for the current sheet
     * @return the next available row index after the written content
     */
    int write(SXSSFSheet sheet, SXSSFWorkbook workbook, int nextRow, SheetContext context);
}
