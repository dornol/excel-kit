package io.github.dornol.excelkit.excel;

import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.jspecify.annotations.Nullable;

import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;
import java.util.Objects;

/**
 * Merges the runs of equal values of the columns that asked for it ({@code mergeSameValues}), vertically, while
 * the rows are written. One tracker lives for one write.
 * <p>
 * A run ends where the value changes, where it is empty, or where the run of a merge column to its left ended, so
 * that a column merges only inside the merged regions of the merge columns before it (a team inside its
 * department). A run of one row is not merged. The cells below the first one keep their values.
 */
final class MergeRunTracker<T> {
    private final int[] ordinalOfColumn;
    private final int[] columns;
    private final int[] start;
    private final @Nullable Object[] value;
    private final boolean[] open;
    private @Nullable SXSSFSheet sheet;
    private int rowIndex;
    private boolean boundary;

    private MergeRunTracker(int columnCount, List<Integer> mergeColumns) {
        this.ordinalOfColumn = new int[columnCount];
        Arrays.fill(ordinalOfColumn, -1);
        this.columns = new int[mergeColumns.size()];
        for (int k = 0; k < columns.length; k++) {
            columns[k] = mergeColumns.get(k);
            ordinalOfColumn[columns[k]] = k;
        }
        this.start = new int[columns.length];
        this.value = new Object[columns.length];
        this.open = new boolean[columns.length];
    }

    /** The tracker for the columns, or {@code null} when none of them merges. */
    static <T> @Nullable MergeRunTracker<T> of(List<ExcelColumn<T>> columns) {
        var merging = new ArrayList<Integer>();
        for (int i = 0; i < columns.size(); i++) {
            if (columns.get(i).isMergeSameValues()) merging.add(i);
        }
        return merging.isEmpty() ? null : new MergeRunTracker<>(columns.size(), merging);
    }

    /** The position of the column among the merge columns, or -1 when it does not merge. */
    int ordinal(int column) {
        return ordinalOfColumn[column];
    }

    /** Called before the cells of a row are written to {@code row} (0-based) of {@code target}. */
    void startRow(SXSSFSheet target, int row) {
        if (target != sheet) {
            sheet = target;
            Arrays.fill(open, false);
        }
        rowIndex = row;
        boundary = false;
    }

    /** Called with the value of a merge column in the current row, the columns in ascending order. */
    void value(int ordinal, @Nullable Object cellValue) {
        boolean mergeable = isMergeable(cellValue);
        if (open[ordinal] && !boundary && mergeable && Objects.equals(value[ordinal], cellValue)) {
            return; // the run goes on
        }
        closeRun(ordinal, rowIndex - 1);
        boundary = true;
        if (mergeable) {
            open[ordinal] = true;
            start[ordinal] = rowIndex;
            value[ordinal] = cellValue;
        }
    }

    /** Ends every open run at {@code lastRow}, the last data row of the sheet being written. */
    void closeRuns(int lastRow) {
        for (int k = 0; k < columns.length; k++) closeRun(k, lastRow);
    }

    private void closeRun(int ordinal, int lastRow) {
        if (!open[ordinal]) return;
        open[ordinal] = false;
        value[ordinal] = null;
        if (lastRow > start[ordinal] && sheet != null) {
            // The regions cannot overlap by construction, so the check of addMergedRegion (which grows with every
            // region already on the sheet) is skipped.
            sheet.addMergedRegionUnsafe(new CellRangeAddress(start[ordinal], lastRow, columns[ordinal], columns[ordinal]));
        }
    }

    private static boolean isMergeable(@Nullable Object cellValue) {
        if (cellValue == null) return false;
        return !(cellValue instanceof CharSequence text) || !text.toString().isBlank();
    }
}
