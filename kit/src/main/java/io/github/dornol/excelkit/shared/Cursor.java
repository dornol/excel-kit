package io.github.dornol.excelkit.shared;

/**
 * Tracks the current writing position during an export operation.
 * <p>
 * Used internally during streaming writing to track row position within the current sheet (or file)
 * and the total number of rows processed globally (across multiple sheets).
 * This allows support for sheet rollover, row-based formatting, and row-level indexing.
 *
 * @author dhkim
 * @since 2025-07-19
 */
public class Cursor {
    private final int baseRow;
    private int rowOfSheet;
    private int currentTotal;

    /**
     * Creates a new Cursor with row index and total count initialized to 0.
     */
    public Cursor() {
        this(0);
    }

    /**
     * Creates a new Cursor starting from a specific row index.
     *
     * @param baseRow The starting row index for each sheet (e.g., if there's a title)
     */
    public Cursor(int baseRow) {
        this.baseRow = baseRow;
        this.rowOfSheet = baseRow;
        this.currentTotal = 0;
    }

    /**
     * Increments the current row index in the current sheet by 1.
     */
    public void plusRow() {
        this.rowOfSheet++;
    }

    /**
     * Resets the current sheet's row index to the base row.
     * Typically called when a new sheet or file section is created.
     */
    public void initRow() {
        this.rowOfSheet = this.baseRow;
    }

    /**
     * Increments the total number of processed rows (across all sheets) by 1.
     */
    public void plusTotal() {
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
