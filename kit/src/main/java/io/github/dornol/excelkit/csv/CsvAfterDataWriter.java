package io.github.dornol.excelkit.csv;

import java.io.PrintWriter;

/**
 * A functional interface for writing custom content after data rows in CSV output.
 * <p>
 * Unlike Excel's {@code AfterDataWriter}, this receives a {@link PrintWriter} directly
 * since CSV has no sheet/workbook concept.
 *
 * @author dhkim
 * @since 2025-07-19
 */
@FunctionalInterface
public interface CsvAfterDataWriter {

    /**
     * Writes custom content after all data rows.
     *
     * @param writer the PrintWriter used to append additional CSV lines
     */
    void write(PrintWriter writer);
}
