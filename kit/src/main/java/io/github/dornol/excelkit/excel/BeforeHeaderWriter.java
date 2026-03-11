package io.github.dornol.excelkit.excel;

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
     * @param context provides the current sheet, workbook, starting row index,
     *                and column metadata
     * @return the next available row index where the column header should start
     */
    int write(SheetContext context);
}
