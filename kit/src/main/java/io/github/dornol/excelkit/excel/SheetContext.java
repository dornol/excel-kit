package io.github.dornol.excelkit.excel;

import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

import java.util.Collections;
import java.util.List;
import java.util.stream.Collectors;

/**
 * Provides contextual information passed to {@link BeforeHeaderWriter} and
 * {@link AfterDataWriter} callbacks, including the current sheet, workbook,
 * row position, and column metadata.
 *
 * <p>A new instance is created for each callback invocation so that the
 * sheet reference is always up-to-date (e.g. after sheet rollover).
 *
 * @author dhkim
 * @since 0.3.0
 */
public class SheetContext {

    private final SXSSFSheet sheet;
    private final SXSSFWorkbook workbook;
    private final int currentRow;
    private final int columnCount;
    private final List<String> columnNames;

    SheetContext(SXSSFSheet sheet, SXSSFWorkbook workbook, int currentRow,
                 List<? extends ExcelColumn<?>> columns) {
        this.sheet = sheet;
        this.workbook = workbook;
        this.currentRow = currentRow;
        this.columnCount = columns.size();
        this.columnNames = Collections.unmodifiableList(
                columns.stream()
                        .map(ExcelColumn::getName)
                        .collect(Collectors.toList())
        );
    }

    /**
     * Returns the current sheet being written to.
     *
     * @return the current sheet
     */
    public SXSSFSheet getSheet() {
        return sheet;
    }

    /**
     * Returns the workbook (useful for creating CellStyles, etc.).
     *
     * @return the workbook
     */
    public SXSSFWorkbook getWorkbook() {
        return workbook;
    }

    /**
     * Returns the current row index available for writing.
     *
     * @return the current row index
     */
    public int getCurrentRow() {
        return currentRow;
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

    /**
     * Converts a zero-based column index to an Excel column letter.
     * <p>
     * Examples:
     * <ul>
     *     <li>0 → "A"</li>
     *     <li>1 → "B"</li>
     *     <li>25 → "Z"</li>
     *     <li>26 → "AA"</li>
     * </ul>
     *
     * @param colIndex zero-based column index
     * @return the Excel column letter(s)
     */
    public static String columnLetter(int colIndex) {
        StringBuilder sb = new StringBuilder();
        int idx = colIndex + 1;
        while (idx > 0) {
            idx--;
            sb.insert(0, (char) ('A' + idx % 26));
            idx /= 26;
        }
        return sb.toString();
    }
}