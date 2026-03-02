package io.github.dornol.excelkit.excel;

import java.util.Collections;
import java.util.List;
import java.util.stream.Collectors;

/**
 * Provides contextual information about the columns configured for a sheet.
 * <p>
 * Passed to {@link BeforeHeaderWriter} and {@link AfterDataWriter} callbacks
 * so that implementations can access column metadata (count, names) without hard-coding.
 *
 * @author dhkim
 * @since 0.4.0
 */
public class SheetContext {

    private final int columnCount;
    private final List<String> columnNames;

    SheetContext(List<? extends ExcelColumn<?>> columns) {
        this.columnCount = columns.size();
        this.columnNames = Collections.unmodifiableList(
                columns.stream()
                        .map(ExcelColumn::getName)
                        .collect(Collectors.toList())
        );
    }

    /**
     * Returns the number of columns configured for this sheet.
     *
     * @return the column count
     */
    public int getColumnCount() {
        return columnCount;
    }

    /**
     * Returns an unmodifiable list of column header names, in order.
     *
     * @return the column names
     */
    public List<String> getColumnNames() {
        return columnNames;
    }
}
