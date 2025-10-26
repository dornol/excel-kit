package io.github.dornol.excelkit.csv;

/**
 * Tracks the current row position during CSV export.
 * <p>
 * This class is used to provide row-level context while writing data,
 * such as current sheet-relative row index or total number of records processed.
 *
 * @author dhkim
 * @since 2025-07-19
 */
public class CsvCursor {
    private int rowOfSheet;
    private int currentTotal;

    /**
     * Constructs a new cursor with both row and total counters set to 0.
     */
    public CsvCursor() {
        this.rowOfSheet = 0;
        this.currentTotal = 0;
    }

    /**
     * Increments the current row index (within the sheet).
     */
    void plusRow() {
        this.rowOfSheet++;
    }

    /**
     * Resets the row index to 0. Typically called when starting a new file or section.
     */
    void initRow() {
        this.rowOfSheet = 0;
    }

    /**
     * Increments the total row count processed (across the entire export).
     */
    void plusTotal() {
        this.currentTotal++;
    }

    /**
     * Returns the current row index relative to the current sheet or file.
     *
     * @return The 0-based row number
     */
    public int getRowOfSheet() {
        return rowOfSheet;
    }

    /**
     * Returns the total number of rows processed so far.
     *
     * @return The cumulative row count
     */
    public int getCurrentTotal() {
        return currentTotal;
    }
}
