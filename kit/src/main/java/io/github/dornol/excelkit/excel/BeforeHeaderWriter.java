ㅁpackage io.github.dornol.excelkit.excel;

import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

/**
 * A functional interface for writing custom content before the column header row.
 * <p>
 * Called on every sheet (including rollover sheets), so the implementation must
 * always produce the same number of rows.
 *
 * @author dhkim
 * @since 2025-07-19
 */
@FunctionalInterface
public interface BeforeHeaderWriter {

    /**
     * Writes custom content to the sheet before the column headers.
     *
     * @param sheet    the current sheet
     * @param workbook the workbook (useful for creating CellStyles, etc.)
     * @param startRow the first row index available for writing
     *                 (after the title rows if a title is set, otherwise 0)
     * @return the next available row index where the column header should start
     */
    int write(SXSSFSheet sheet, SXSSFWorkbook workbook, int startRow);
}
