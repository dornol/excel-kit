package io.github.dornol.excelkit.excel;

import org.apache.poi.xssf.streaming.SXSSFCell;
import org.apache.poi.xssf.streaming.SXSSFRow;
import org.jspecify.annotations.Nullable;

import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

/**
 * Fluent DSL for adding summary (footer) rows with formulas such as SUM, AVERAGE, COUNT, MIN, MAX.
 * <p>
 * Usage:
 * <pre>{@code
 * writer
 *     .summary(s -> s
 *         .label("Total")
 *         .sum("Amount")
 *         .average("Score"))
 *     .write(data);
 * }</pre>
 *
 * @author dhkim
 * @since 0.7.2
 */
public class ExcelSummary {

    /**
     * Supported summary operations.
     */
    public enum Op {
        SUM, AVERAGE, COUNT, MIN, MAX
    }

    private final Map<Op, List<String>> entries = new LinkedHashMap<>();
    private @Nullable String labelColumnName;
    private @Nullable String labelText;

    /**
     * Sets the label text for the summary row(s).
     * The label is placed in the first column by default.
     *
     * @param text the label text (e.g., "Total", "Summary")
     * @return this summary for chaining
     */
    public ExcelSummary label(String text) {
        this.labelText = text;
        this.labelColumnName = null;
        return this;
    }

    /**
     * Sets the label text for the summary row(s) in a specific column.
     *
     * @param columnName the column to place the label in
     * @param text       the label text
     * @return this summary for chaining
     */
    public ExcelSummary label(String columnName, String text) {
        this.labelColumnName = columnName;
        this.labelText = text;
        return this;
    }

    /**
     * Adds a SUM formula for the specified column.
     */
    public ExcelSummary sum(String columnName) {
        return addEntry(Op.SUM, columnName);
    }

    /**
     * Adds an AVERAGE formula for the specified column.
     */
    public ExcelSummary average(String columnName) {
        return addEntry(Op.AVERAGE, columnName);
    }

    /**
     * Adds a COUNT formula for the specified column.
     */
    public ExcelSummary count(String columnName) {
        return addEntry(Op.COUNT, columnName);
    }

    /**
     * Adds a MIN formula for the specified column.
     */
    public ExcelSummary min(String columnName) {
        return addEntry(Op.MIN, columnName);
    }

    /**
     * Adds a MAX formula for the specified column.
     */
    public ExcelSummary max(String columnName) {
        return addEntry(Op.MAX, columnName);
    }

    private ExcelSummary addEntry(Op op, String columnName) {
        entries.computeIfAbsent(op, k -> new ArrayList<>()).add(columnName);
        return this;
    }

    /**
     * Converts this summary configuration into an {@link AfterDataWriter} callback.
     */
    AfterDataWriter toAfterDataWriter() {
        return ctx -> {
            List<String> columnNames = ctx.getColumnNames();
            int headerRow = ctx.getHeaderRowIndex();
            int currentRow = ctx.getCurrentRow();

            // Data range: from row after header to last data row (1-based for Excel)
            int dataStartRow = headerRow + 2; // 1-based, skip header
            int dataEndRow = currentRow;       // 1-based (currentRow is 0-based next row, so = last data row + 1 in 0-based = last data row in 1-based)

            int row = currentRow;
            for (var entry : entries.entrySet()) {
                Op op = entry.getKey();
                List<String> cols = entry.getValue();

                SXSSFRow summaryRow = ctx.getSheet().createRow(row);

                // Label
                int labelIdx = 0;
                if (labelColumnName != null) {
                    int idx = columnNames.indexOf(labelColumnName);
                    if (idx >= 0) labelIdx = idx;
                }
                String text = entries.size() > 1
                        ? op.name().substring(0, 1) + op.name().substring(1).toLowerCase()
                        : (labelText != null ? labelText : op.name().substring(0, 1) + op.name().substring(1).toLowerCase());
                summaryRow.createCell(labelIdx).setCellValue(text);

                // Formulas
                for (String colName : cols) {
                    int colIdx = columnNames.indexOf(colName);
                    if (colIdx < 0) continue;
                    String colLetter = SheetContext.columnLetter(colIdx);
                    String formula = op.name() + "(" + colLetter + dataStartRow + ":" + colLetter + dataEndRow + ")";
                    SXSSFCell cell = summaryRow.createCell(colIdx);
                    cell.setCellFormula(formula);
                }

                row++;
            }

            return row;
        };
    }
}
