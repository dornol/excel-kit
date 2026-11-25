package io.github.dornol.excelkit.excel;

import io.github.dornol.excelkit.excel.style.ExcelColor;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.xssf.streaming.SXSSFCell;
import org.apache.poi.xssf.streaming.SXSSFRow;
import org.jspecify.annotations.Nullable;

import java.util.Locale;
import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.Objects;
import java.util.function.Consumer;

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

    /** The look of the summary rows: bold on light gray, like the subtotal rows, until {@link #style(Consumer)} changes it. */
    private final RowStyleConfig style = new RowStyleConfig().bold(true).backgroundColor(ExcelColor.LIGHT_GRAY);

    /** Creates a new summary configuration. */
    public ExcelSummary() {}

    /**
     * Styles the summary rows (bold on light gray by default, as the subtotal rows of {@code groupSubtotal(...)} are).
     * Every cell of a summary row takes the number format and alignment of its column, so a total reads as the
     * numbers above it, and is locked under {@code protectSheet(...)}. To get plain rows as before, clear what the
     * default sets: {@code style(r -> r.bold(false).backgroundColor(null))}.
     *
     * @param configurer changes the default style
     * @return this summary for chaining
     * @since 0.29.0
     */
    public ExcelSummary style(Consumer<RowStyleConfig> configurer) {
        Objects.requireNonNull(configurer, "configurer cannot be null").accept(style);
        return this;
    }

    /**
     * Supported summary operations.
     */
    public enum Op {
        /** Sum formula. */
        SUM,
        /** Average formula. */
        AVERAGE,
        /** Count formula: counts the cells that hold a number. */
        COUNT,
        /** Count formula that counts every cell that is not empty, text included ({@code COUNTA}). */
        COUNTA,
        /** Min formula. */
        MIN,
        /** Max formula. */
        MAX
    }

    private final Map<Op, List<String>> entries = new LinkedHashMap<>();
    private @Nullable String labelColumnName;
    private @Nullable String labelText;

    /**
     * Sets the label text for the summary row(s).
     * The label is placed in the first column by default.
     * <p>
     * With one operation the row carries the label as it is. With several, every row needs its own name, so each
     * label is followed by its operation: {@code "Total - Sum"}, {@code "Total - Average"}. Without a label a row is
     * named after its operation alone.
     *
     * @param text the label text (e.g., "Total", "Summary")
     * @return this summary for chaining
     */
    public ExcelSummary label(String text) {
        Objects.requireNonNull(text, "text cannot be null");
        this.labelText = text;
        this.labelColumnName = null;
        return this;
    }

    /**
     * Sets the label text for the summary row(s) in a specific column; with several operations each label is followed
     * by its operation, as for {@link #label(String)}.
     *
     * @param columnName the column to place the label in
     * @param text       the label text
     * @return this summary for chaining
     */
    public ExcelSummary label(String columnName, String text) {
        Objects.requireNonNull(columnName, "columnName cannot be null");
        Objects.requireNonNull(text, "text cannot be null");
        this.labelColumnName = columnName;
        this.labelText = text;
        return this;
    }

    /**
     * Adds a SUM formula for the specified column.
     *
     * @param columnName the target column name
     * @return this summary for chaining
     */
    public ExcelSummary sum(String columnName) {
        return addEntry(Op.SUM, columnName);
    }

    /**
     * Adds an AVERAGE formula for the specified column.
     *
     * @param columnName the target column name
     * @return this summary for chaining
     */
    public ExcelSummary average(String columnName) {
        return addEntry(Op.AVERAGE, columnName);
    }

    /**
     * Adds a COUNT formula for the specified column.
     *
     * @param columnName the target column name
     * @return this summary for chaining
     */
    public ExcelSummary count(String columnName) {
        return addEntry(Op.COUNT, columnName);
    }

    /**
     * Adds a COUNTA formula for the specified column: it counts the cells that are not empty, text included,
     * where {@link #count(String)} counts only the cells that hold a number.
     *
     * @param columnName the target column name
     * @return this summary for chaining
     * @since 0.29.0
     */
    public ExcelSummary countA(String columnName) {
        return addEntry(Op.COUNTA, columnName);
    }

    /**
     * Adds a MIN formula for the specified column.
     *
     * @param columnName the target column name
     * @return this summary for chaining
     */
    public ExcelSummary min(String columnName) {
        return addEntry(Op.MIN, columnName);
    }

    /**
     * Adds a MAX formula for the specified column.
     *
     * @param columnName the target column name
     * @return this summary for chaining
     */
    public ExcelSummary max(String columnName) {
        return addEntry(Op.MAX, columnName);
    }

    private ExcelSummary addEntry(Op op, String columnName) {
        Objects.requireNonNull(columnName, "columnName cannot be null");
        entries.computeIfAbsent(op, k -> new ArrayList<>()).add(columnName);
        return this;
    }

    /**
     * The number format a summary cell takes instead of its column's, or {@code null} for the column's own. A sum, a
     * minimum and a maximum read in the unit of the column; a count is a number of cells, whatever the column holds
     * (not "$2.00" or a date); an average of whole numbers shows the decimals it has, up to two (8.5, not 9; 3, not 3.00).
     *
     * @param calculated the result written next to the formula, or {@code null} when it is not known
     */
    static @Nullable String formatOf(Op op, String columnFormat, @Nullable Double calculated) {
        return switch (op) {
            case COUNT, COUNTA -> "#,##0";
            case AVERAGE -> withDecimals(columnFormat, calculated);
            default -> null;
        };
    }

    /**
     * A whole-number format ({@code #,##0}, {@code 0}, {@code 0%}) with the decimals the value needs, up to two, so that
     * 3 stays {@code 3}, 8.5 reads {@code 8.5} and 8.25 reads {@code 8.25}; without a known value, two decimals. The
     * format is fixed when the file is written: an average that a later edit turns fractional shows the decimals of the
     * one written. A format with decimals, text or dates is kept ({@code null}).
     */
    private static @Nullable String withDecimals(String format, @Nullable Double calculated) {
        java.util.regex.Matcher m = java.util.regex.Pattern.compile("([#0,]+)(%?)").matcher(format);
        if (!m.matches()) return null;
        int decimals = 2;
        if (calculated != null && Double.isFinite(calculated)) {
            double shown = calculated * (m.group(2).isEmpty() ? 1 : 100); // a percent format shows the value times 100
            java.math.BigDecimal rounded = java.math.BigDecimal.valueOf(shown).setScale(2, java.math.RoundingMode.HALF_UP)
                    .stripTrailingZeros();
            decimals = Math.max(0, rounded.scale());
        }
        if (decimals == 0) return null;
        return m.group(1) + "." + "0".repeat(decimals) + m.group(2);
    }

    Map<Op, List<String>> entries() {
        return entries;
    }

    /** The number of rows the summary writes: one for every operation. */
    int rowCount() {
        return entries.size();
    }

    /**
     * Converts this summary configuration into an {@link AfterDataWriter} callback.
     */
    /**
     * @param subtotals    whether the data holds subtotal rows: then {@code SUBTOTAL} counts every row once
     * @param columns      the columns of the sheet: a summary cell takes the number format and alignment of its column
     * @param styleCache   the styles derived from the columns' ones, shared with the data rows
     * @param maxStyles    the cap on cell styles of the workbook
     */
    AfterDataWriter toAfterDataWriter(boolean subtotals, @Nullable FormulaValues values, List<? extends ExcelColumn<?>> columns,
                                      Map<ExcelStyleSupporter.DerivedStyleKey, CellStyle> styleCache, int maxStyles) {
        // Locked, so that a column a person may edit does not let them overwrite a formula.
        RowStyleConfig locked = style.lockedCopy();
        return ctx -> {
            List<String> columnNames = ctx.getColumnNames();
            int headerRow = ctx.getHeaderRowIndex();
            int currentRow = ctx.getCurrentRow();

            // Data range: from row after header to last data row (1-based for Excel)
            int dataStartRow = headerRow + 2; // 1-based, skip header
            int dataEndRow = ctx.dataEndRow(); // 1-based last data row, even when afterData wrote rows below it

            int row = currentRow;
            for (var entry : entries.entrySet()) {
                Op op = entry.getKey();
                List<String> cols = entry.getValue();

                SXSSFRow summaryRow = ctx.getSheet().createRow(row);
                // every cell of the row, so that the band runs across the whole width of the sheet
                CellStyle[] styles = new CellStyle[columns.size()];
                for (int i = 0; i < styles.length; i++) {
                    styles[i] = ExcelStyleSupporter.resolveRowStyle(columns.get(i).getStyle(), style.backgroundColor, locked,
                            styleCache, ctx.getWorkbook(), maxStyles);
                    summaryRow.createCell(i).setCellStyle(styles[i]);
                }

                // Label
                int labelIdx = 0;
                if (labelColumnName != null) {
                    int idx = columnNames.indexOf(labelColumnName);
                    if (idx < 0) throw new ExcelWriteException("Summary label column not found: '"
                            + labelColumnName + "'. Columns: " + columnNames);
                    labelIdx = idx;
                }
                String operation = op == Op.COUNTA ? "CountA" : op.name().substring(0, 1) + op.name().substring(1).toLowerCase(Locale.ROOT);
                // One operation: the given label as it is. Several: every row needs its own name, so the label is
                // followed by the operation ("Total - Sum"). Without a label the operation names the row.
                String text = labelText == null ? operation
                        : entries.size() > 1 ? labelText + " - " + operation : labelText;
                ExcelCellText.write(summaryRow.getCell(labelIdx), text);

                // Formulas
                for (String colName : cols) {
                    int colIdx = columnNames.indexOf(colName);
                    if (colIdx < 0) throw new ExcelWriteException("Summary column not found: '"
                            + colName + "'. Columns: " + columnNames);
                    // Without data rows the range would be reversed (A2:A1) and include the header.
                    if (dataEndRow < dataStartRow) continue;
                    String colLetter = SheetContext.columnLetter(colIdx);
                    String range = colLetter + dataStartRow + ":" + colLetter + dataEndRow;
                    String formula = subtotals ? "SUBTOTAL(" + ExcelSubtotal.functionNumber(op) + "," + range + ")"
                            : op.name() + "(" + range + ")";
                    SXSSFCell cell = summaryRow.getCell(colIdx);
                    Double calculated = values == null ? null : values.sheetValue(colIdx, op);
                    ExcelDataType.setFormula(cell, formula, calculated);
                    String format = formatOf(op, columns.get(colIdx).getStyle().getDataFormatString(), calculated);
                    if (format != null) {
                        cell.setCellStyle(ExcelStyleSupporter.resolveRowStyle(columns.get(colIdx).getStyle(), style.backgroundColor,
                                locked, styleCache, ctx.getWorkbook(), maxStyles, format));
                    }
                }

                row++;
            }

            return row;
        };
    }
}
