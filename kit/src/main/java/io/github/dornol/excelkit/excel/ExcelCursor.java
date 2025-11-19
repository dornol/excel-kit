package io.github.dornol.excelkit.excel;

/**
 * Tracks the current writing position within an Excel export operation.
 * <p>
 * Used internally during streaming writing to track row position within the current sheet
 * and the total number of rows processed globally (across multiple sheets).
 * This allows support for sheet rollover, row-based formatting, and row-level indexing.
 *
 * @author dhkim
 * @since 2025-07-19
 */
public class ExcelCursor {
    private int rowOfSheet;
    private int currentTotal;

    /**
     * Creates a new ExcelCursor with row index and total count initialized to 0.
     */
    ExcelCursor() {
        this.rowOfSheet = 0;
        this.currentTotal = 0;
    }

    /**
     * Increments the current row index in the current sheet by 1.
     */
    void plusRow() {
        this.rowOfSheet++;
    }

    /**
     * Resets the current sheet's row index to 0.
     * Typically called when a new sheet is created.
     */
    void initRow() {
        this.rowOfSheet = 0;
    }

    /**
     * Increments the total number of processed rows (across all sheets) by 1.
     */
    void plusTotal() {
        this.currentTotal++;
    }

    /**
     * Returns the current row index within the sheet.
     *
     * @return Row index in the current sheet (0-based)
     */
    public int getRowOfSheet() {
        return rowOfSheet;
    }

    /**
     * Returns the total number of processed rows, across all sheets.
     *
     * @return Total number of rows written
     */
    public int getCurrentTotal() {
        return currentTotal;
    }
}
