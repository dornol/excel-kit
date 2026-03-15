package io.github.dornol.excelkit.excel;

import org.apache.poi.ss.usermodel.ComparisonOperator;
import org.apache.poi.ss.usermodel.ConditionalFormattingRule;
import org.apache.poi.ss.usermodel.PatternFormatting;
import org.apache.poi.ss.usermodel.SheetConditionalFormatting;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFColor;

import org.jspecify.annotations.Nullable;

import java.util.ArrayList;
import java.util.List;

/**
 * Builder for conditional formatting rules to apply to Excel sheets.
 * <p>
 * Supports cell-value-based rules with background color styling.
 *
 * <pre>{@code
 * new ExcelWriter<Product>()
 *     .addColumn("Price", Product::getPrice, c -> c.type(ExcelDataType.INTEGER))
 *     .conditionalFormatting(cf -> cf
 *         .columns(1)
 *         .greaterThan("1000", ExcelColor.LIGHT_RED)
 *         .lessThan("100", ExcelColor.LIGHT_GREEN))
 *     .write(stream)
 *     .consumeOutputStream(out);
 * }</pre>
 *
 * @author dhkim
 * @since 0.6.0
 */
public class ExcelConditionalRule {

    private final List<RuleEntry> rules = new ArrayList<>();
    private int @Nullable [] columnIndices;
    private int startRow = -1;

    /**
     * Sets the columns to apply conditional formatting to (0-based indices).
     *
     * @param indices 0-based column indices
     * @return this rule for chaining
     */
    public ExcelConditionalRule columns(int... indices) {
        this.columnIndices = indices;
        return this;
    }

    /**
     * Sets the starting data row (0-based).
     * If not set, defaults to the row after the header.
     *
     * @param startRow 0-based starting row
     * @return this rule for chaining
     */
    public ExcelConditionalRule startRow(int startRow) {
        this.startRow = startRow;
        return this;
    }

    /**
     * Adds a rule: cell value greater than the given value.
     */
    public ExcelConditionalRule greaterThan(String value, ExcelColor bgColor) {
        rules.add(new RuleEntry(ComparisonOperator.GT, value, null, bgColor));
        return this;
    }

    /**
     * Adds a rule: cell value greater than or equal to the given value.
     */
    public ExcelConditionalRule greaterThanOrEqual(String value, ExcelColor bgColor) {
        rules.add(new RuleEntry(ComparisonOperator.GE, value, null, bgColor));
        return this;
    }

    /**
     * Adds a rule: cell value less than the given value.
     */
    public ExcelConditionalRule lessThan(String value, ExcelColor bgColor) {
        rules.add(new RuleEntry(ComparisonOperator.LT, value, null, bgColor));
        return this;
    }

    /**
     * Adds a rule: cell value less than or equal to the given value.
     */
    public ExcelConditionalRule lessThanOrEqual(String value, ExcelColor bgColor) {
        rules.add(new RuleEntry(ComparisonOperator.LE, value, null, bgColor));
        return this;
    }

    /**
     * Adds a rule: cell value equal to the given value.
     */
    public ExcelConditionalRule equalTo(String value, ExcelColor bgColor) {
        rules.add(new RuleEntry(ComparisonOperator.EQUAL, value, null, bgColor));
        return this;
    }

    /**
     * Adds a rule: cell value not equal to the given value.
     */
    public ExcelConditionalRule notEqualTo(String value, ExcelColor bgColor) {
        rules.add(new RuleEntry(ComparisonOperator.NOT_EQUAL, value, null, bgColor));
        return this;
    }

    /**
     * Adds a rule: cell value between the given values (inclusive).
     */
    public ExcelConditionalRule between(String value1, String value2, ExcelColor bgColor) {
        rules.add(new RuleEntry(ComparisonOperator.BETWEEN, value1, value2, bgColor));
        return this;
    }

    /**
     * Adds a rule: cell value not between the given values.
     */
    public ExcelConditionalRule notBetween(String value1, String value2, ExcelColor bgColor) {
        rules.add(new RuleEntry(ComparisonOperator.NOT_BETWEEN, value1, value2, bgColor));
        return this;
    }

    /**
     * Applies all configured rules to the given sheet.
     * Package-private, called by the writer.
     */
    void apply(SXSSFSheet sheet, int headerRowIndex, int columnCount) {
        if (rules.isEmpty()) return;

        SheetConditionalFormatting scf = sheet.getSheetConditionalFormatting();
        int dataStartRow = (startRow >= 0) ? startRow : headerRowIndex + 1;

        int[] cols = (columnIndices != null) ? columnIndices : defaultColumnRange(columnCount);

        for (int colIdx : cols) {
            CellRangeAddress[] ranges = {
                    new CellRangeAddress(dataStartRow, ExcelWriteSupport.EXCEL_MAX_ROWS, colIdx, colIdx)
            };

            for (RuleEntry entry : rules) {
                ConditionalFormattingRule rule = scf.createConditionalFormattingRule(
                        entry.operator, entry.value1, entry.value2);
                PatternFormatting pf = rule.createPatternFormatting();
                pf.setFillBackgroundColor(new XSSFColor(new byte[]{
                        (byte) entry.bgColor.getR(), (byte) entry.bgColor.getG(), (byte) entry.bgColor.getB()}));
                scf.addConditionalFormatting(ranges, rule);
            }
        }
    }

    private int[] defaultColumnRange(int columnCount) {
        int[] result = new int[columnCount];
        for (int i = 0; i < columnCount; i++) {
            result[i] = i;
        }
        return result;
    }

    private record RuleEntry(byte operator, String value1, String value2, ExcelColor bgColor) {
    }
}
