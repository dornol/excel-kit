package io.github.dornol.excelkit.excel;

import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.xssf.streaming.SXSSFCell;
import org.jspecify.annotations.Nullable;

import java.util.List;
import java.util.Map;

/**
 * The values the {@code SUBTOTAL} and summary formulas of a sheet will calculate, worked out while the rows are
 * written, so that a formula carries its result next to it. Excel recalculates when it opens the file, but a file
 * seen in Protected View, a preview or another program that does not calculate shows an empty cell without it.
 * <p>
 * Every data cell of a column that a formula covers is read back from the cell itself, once it is written: a cell
 * counts for the number functions only when it holds a number in the file (a text cell does not, a date does), which is
 * how Excel reads it. Nothing is kept per row: one small accumulator per formula column, for the open group and for
 * the sheet. A result that cannot be known for certain (the column holds formula cells of its own, or error cells) is
 * left out, and the formula is written without a result as before.
 */
final class FormulaValues {

    /** What one column holds, over the rows of a group or of a sheet. */
    static final class Accumulator {
        private long numbers;
        private double sum;
        private double min = Double.POSITIVE_INFINITY;
        private double max = Double.NEGATIVE_INFINITY;
        private long nonEmpty;
        /** A formula or an error cell: the number functions cannot be known. */
        private boolean uncertain;
        /** An empty text cell: whether {@code COUNTA} counts it is not settled here. */
        private boolean ambiguousText;

        void add(SXSSFCell cell) {
            CellType type = cell.getCellType();
            switch (type) {
                case NUMERIC -> {
                    double v = cell.getNumericCellValue();
                    numbers++;
                    nonEmpty++;
                    sum += v;
                    if (v < min) min = v;
                    if (v > max) max = v;
                }
                case STRING -> {
                    nonEmpty++;
                    if (cell.getStringCellValue().isEmpty()) ambiguousText = true;
                }
                case BOOLEAN -> nonEmpty++;
                case FORMULA, ERROR -> {
                    nonEmpty++;
                    uncertain = true;
                }
                default -> { }
            }
        }

        /** What Excel calculates for the function over the cells added, or {@code null} when it is not known. */
        @Nullable Double value(ExcelSummary.Op op) {
            Double result = switch (op) {
                case COUNTA -> ambiguousText ? null : (double) nonEmpty;
                case SUM -> uncertain ? null : (Double) sum;
                case COUNT -> uncertain ? null : (double) numbers;
                case AVERAGE -> uncertain || numbers == 0 ? null : (Double) (sum / numbers);
                case MIN -> uncertain ? null : numbers == 0 ? 0.0 : (Double) min;
                case MAX -> uncertain ? null : numbers == 0 ? 0.0 : (Double) max;
            };
            return result == null || !Double.isFinite(result) ? null : result;
        }

        void reset() {
            numbers = 0;
            sum = 0;
            min = Double.POSITIVE_INFINITY;
            max = Double.NEGATIVE_INFINITY;
            nonEmpty = 0;
            uncertain = false;
            ambiguousText = false;
        }
    }

    private final @Nullable Accumulator[] group;
    private final @Nullable Accumulator[] sheet;

    private FormulaValues(@Nullable Accumulator[] group, @Nullable Accumulator[] sheet) {
        this.group = group;
        this.sheet = sheet;
    }

    /** The values of the formulas the configuration writes, or {@code null} when it writes none. */
    static <T> @Nullable FormulaValues of(SheetConfig<T> cfg, List<ExcelColumn<T>> columns) {
        Accumulator[] group = new Accumulator[columns.size()];
        Accumulator[] sheet = new Accumulator[columns.size()];
        boolean any = false;
        if (cfg.subtotalConfig != null) {
            for (List<String> columnNames : cfg.subtotalConfig.entries().values()) {
                for (String name : columnNames) {
                    int index = SheetWriteEngine.indexOfColumn(columns, name);
                    if (index >= 0) {
                        group[index] = new Accumulator();
                        any = true;
                    }
                }
            }
        }
        if (cfg.summaryConfig != null) {
            for (Map.Entry<ExcelSummary.Op, List<String>> entry : cfg.summaryConfig.entries().entrySet()) {
                for (String name : entry.getValue()) {
                    int index = SheetWriteEngine.indexOfColumn(columns, name);
                    if (index >= 0) {
                        sheet[index] = new Accumulator();
                        any = true;
                    }
                }
            }
        }
        return any ? new FormulaValues(group, sheet) : null;
    }

    /** Called with a data cell of the column once it is written. */
    void cell(int column, SXSSFCell cell) {
        Accumulator g = group[column];
        if (g != null) g.add(cell);
        Accumulator s = sheet[column];
        if (s != null) s.add(cell);
    }

    /** What the subtotal of the open group calculates in the column, or {@code null}. */
    @Nullable Double groupValue(int column, ExcelSummary.Op op) {
        Accumulator g = group[column];
        return g == null ? null : g.value(op);
    }

    /** What the summary of the sheet calculates in the column, or {@code null}. */
    @Nullable Double sheetValue(int column, ExcelSummary.Op op) {
        Accumulator s = sheet[column];
        return s == null ? null : s.value(op);
    }

    /** A group is closed: the next row starts a new one. */
    void resetGroup() {
        for (Accumulator a : group) if (a != null) a.reset();
    }

    /** A sheet is finished (summary written): the next row starts a new one. */
    void resetSheet() {
        for (Accumulator a : sheet) if (a != null) a.reset();
    }
}
