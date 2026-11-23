package io.github.dornol.excelkit.excel;

import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.jspecify.annotations.Nullable;

import java.util.function.BiPredicate;

/**
 * Sets a page break above a data row when the writer's {@code pageBreakWhen} condition holds between the row
 * before it and that row. One tracker lives for one write; a rolled-over sheet starts without a previous row, so
 * its first data row is never given a break right below the header.
 */
final class PageBreakTracker<T> {
    private final @Nullable BiPredicate<T, T> when;
    private @Nullable SXSSFSheet currentSheet;
    private @Nullable T previous;

    PageBreakTracker(@Nullable BiPredicate<T, T> when) {
        this.when = when;
    }

    /**
     * Called before {@code row} is written to {@code rowIndex} (0-based) of {@code sheet}.
     */
    void beforeRow(SXSSFSheet sheet, int rowIndex, T row) {
        if (when == null) return;
        if (sheet != currentSheet) {
            currentSheet = sheet;
            previous = null;
        }
        T before = previous;
        if (before != null && when.test(before, row)) sheet.setRowBreak(rowIndex - 1); // POI breaks below the row it is given: below the row before this one
        previous = row;
    }
}
