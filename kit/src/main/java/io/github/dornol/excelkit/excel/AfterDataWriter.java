package io.github.dornol.excelkit.excel;

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
     * @param context provides the current sheet, workbook, next available row index,
     *                and column metadata
     * @return the next available row index after the written content
     */
    int write(SheetContext context);
}
