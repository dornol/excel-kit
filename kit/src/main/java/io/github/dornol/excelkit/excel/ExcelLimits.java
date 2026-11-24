package io.github.dornol.excelkit.excel;

import org.apache.poi.ss.SpreadsheetVersion;

/**
 * The limits of an .xlsx file and the defaults the writers derive from them. Kept apart from
 * {@link SheetWriteEngine}, which works on {@link SheetConfig}, so the configuration and option classes can use
 * these numbers without depending on the write logic.
 *
 * @author dhkim
 */
final class ExcelLimits {

    /** Rows sampled to size a column when auto-width is on. */
    static final int AUTO_WIDTH_SAMPLE_ROWS = 100;

    /** Rows an Excel sheet can hold; row indices run from 0 to {@code SHEET_ROW_CAPACITY - 1}. */
    static final int SHEET_ROW_CAPACITY = 1_048_576;

    /** The last row index an Excel sheet can hold, which is also the most data rows below a header row. */
    static final int EXCEL_MAX_ROWS = SHEET_ROW_CAPACITY - 1;

    /** Data rows per sheet before a writer rolls over to a new sheet. */
    static final int DEFAULT_MAX_ROWS = 1_000_000;

    /**
     * Cell styles an .xlsx workbook can hold. Every distinct color a {@code cellColor}, {@code rowColor}
     * or {@code rowStyle} function returns needs a style of its own, so a function with many distinct
     * results runs into this.
     */
    static final int MAX_CELL_STYLES = SpreadsheetVersion.EXCEL2007.getMaxCellStyles();

    /** The most characters a cell of an .xlsx workbook holds (32,767). */
    static final int MAX_CELL_TEXT = SpreadsheetVersion.EXCEL2007.getMaxTextLength();

    /**
     * Validates a {@code maxRows} setting: positive, and small enough that a sheet can hold that many
     * data rows below its header row.
     */
    static int requireMaxRows(int maxRows) {
        if (maxRows <= 0) {
            throw new IllegalArgumentException("maxRows must be positive");
        }
        if (maxRows > EXCEL_MAX_ROWS) {
            throw new IllegalArgumentException("maxRows must be at most " + EXCEL_MAX_ROWS
                    + ", the most data rows an Excel sheet can hold below its header row, got " + maxRows);
        }
        return maxRows;
    }

    private ExcelLimits() {
    }
}
