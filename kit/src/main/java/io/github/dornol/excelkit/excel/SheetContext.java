package io.github.dornol.excelkit.excel;

import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

import java.util.List;

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
    private final int headerRowIndex;

    SheetContext(SXSSFSheet sheet, SXSSFWorkbook workbook, int currentRow,
                 List<? extends ExcelColumn<?>> columns) {
        this(sheet, workbook, currentRow, columns, 0);
    }

    SheetContext(SXSSFSheet sheet, SXSSFWorkbook workbook, int currentRow,
                 List<? extends ExcelColumn<?>> columns, int headerRowIndex) {
        this.sheet = sheet;
        this.workbook = workbook;
        this.currentRow = currentRow;
        this.columnCount = columns.size();
        this.columnNames = List.copyOf(columns.stream().map(ExcelColumn::getName).toList());
        this.headerRowIndex = headerRowIndex;
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
    /**
     * Returns the header row index (zero-based) for this sheet.
     *
     * @return the header row index
     */
    public int getHeaderRowIndex() {
        return headerRowIndex;
    }

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
     * Merges a rectangular region of cells identified by zero-based row and
     * column indices.
     *
     * <p>Example – merge the first three columns of row 0:
     * <pre>{@code ctx.mergeCells(0, 0, 0, 2);}</pre>
     *
     * @param firstRow zero-based index of the first row in the region
     * @param lastRow  zero-based index of the last row in the region
     * @param firstCol zero-based index of the first column in the region
     * @param lastCol  zero-based index of the last column in the region
     * @return this {@code SheetContext} for method chaining
     */
    public SheetContext mergeCells(int firstRow, int lastRow, int firstCol, int lastCol) {
        sheet.addMergedRegion(new CellRangeAddress(firstRow, lastRow, firstCol, lastCol));
        return this;
    }

    /**
     * Merges a rectangular region of cells identified by an Excel-style range
     * string such as {@code "A1:C3"}.
     *
     * <p>Example – merge cells A1 through C1:
     * <pre>{@code ctx.mergeCells("A1:C1");}</pre>
     *
     * @param cellRange an Excel-notation range (e.g. {@code "A1:C3"})
     * @return this {@code SheetContext} for method chaining
     */
    public SheetContext mergeCells(String cellRange) {
        sheet.addMergedRegion(CellRangeAddress.valueOf(cellRange));
        return this;
    }

    /**
     * Groups (outlines) a range of rows so they can be collapsed/expanded in Excel.
     * <p>
     * Typically called from an {@link AfterDataWriter} callback.
     *
     * @param firstRow zero-based index of the first row to group
     * @param lastRow  zero-based index of the last row to group
     * @return this {@code SheetContext} for method chaining
     * @since 0.7.0
     */
    public SheetContext groupRows(int firstRow, int lastRow) {
        sheet.groupRow(firstRow, lastRow);
        return this;
    }

    /**
     * Groups (outlines) a range of rows, optionally collapsing them.
     *
     * @param firstRow  zero-based index of the first row to group
     * @param lastRow   zero-based index of the last row to group
     * @param collapsed whether the group should be initially collapsed
     * @return this {@code SheetContext} for method chaining
     * @since 0.7.0
     */
    public SheetContext groupRows(int firstRow, int lastRow, boolean collapsed) {
        sheet.groupRow(firstRow, lastRow);
        if (collapsed) {
            sheet.setRowGroupCollapsed(firstRow, true);
        }
        return this;
    }

    /**
     * Creates a workbook-scoped named range with the given reference formula.
     *
     * @param name      the name for the range (e.g., "Categories")
     * @param reference the reference formula (e.g., "Sheet1!$A$1:$A$10")
     * @return this {@code SheetContext} for method chaining
     */
    public SheetContext namedRange(String name, String reference) {
        var namedRange = workbook.createName();
        namedRange.setNameName(name);
        namedRange.setRefersToFormula(reference);
        return this;
    }

    /**
     * Creates a workbook-scoped named range for a column range on the current sheet.
     *
     * @param name     the name for the range
     * @param col      zero-based column index
     * @param firstRow zero-based first row index
     * @param lastRow  zero-based last row index
     * @return this {@code SheetContext} for method chaining
     */
    public SheetContext namedRange(String name, int col, int firstRow, int lastRow) {
        String sheetName = sheet.getSheetName();
        String colLetter = columnLetter(col);
        String ref = "'" + sheetName + "'!$" + colLetter + "$" + (firstRow + 1)
                + ":$" + colLetter + "$" + (lastRow + 1);
        return namedRange(name, ref);
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