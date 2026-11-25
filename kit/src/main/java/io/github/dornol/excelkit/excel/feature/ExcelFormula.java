package io.github.dornol.excelkit.excel.feature;

import org.jspecify.annotations.Nullable;

import java.util.Objects;

/**
 * An Excel formula, optionally with its result, as the value of a column.
 * <p>
 * A formula written as a plain string ({@link io.github.dornol.excelkit.excel.ExcelDataType#FORMULA}) has no result in
 * the file: Excel calculates it when it opens the file, but a viewer that does not calculate (Excel's Protected View
 * for a download or an attachment, a preview pane, a program that reads the file) shows the cell empty. With a result
 * the cell holds both, and Excel still recalculates when the file is opened.
 *
 * <pre>{@code
 * writer.column("Total", order -> ExcelFormula.of("B2*C2", order.price() * order.qty()))
 *       .column("Label", order -> ExcelFormula.of("A2&\"!\"", order.name() + "!"))
 *       .column("Plain", order -> ExcelFormula.of("B2*C2"))      // no result, as a string formula
 *       .write(stream);
 * }</pre>
 * In a column function that takes the row and the cursor, {@code cursor.getRowOfSheet()} is the number of the row as Excel
 * numbers it, so a formula can point at its own row: {@code "B" + cursor.getRowOfSheet() + "*2"}.
 * <p>
 * The result is yours to give: the library does not calculate it, and a wrong one shows until the file is recalculated.
 * <p>
 * <b>Security warning:</b> do not build the formula from untrusted input, as for
 * {@link io.github.dornol.excelkit.excel.ExcelDataType#FORMULA}.
 * <p>
 * In a CSV file the value is written as its result, or as the formula when it has none.
 *
 * @since 0.29.0
 */
public final class ExcelFormula {
    /** The longest text a cell holds. */
    private static final int MAX_RESULT_TEXT = 32_767;

    private final String formula;
    private final @Nullable Object result;

    private ExcelFormula(String formula, @Nullable Object result) {
        this.formula = formula;
        this.result = result;
    }

    /**
     * A formula without a result.
     *
     * @param formula the formula, with or without a leading {@code =}
     * @return the formula
     * @throws IllegalArgumentException if the formula is blank
     */
    public static ExcelFormula of(String formula) {
        return new ExcelFormula(normalize(formula), null);
    }

    /**
     * A formula with a numeric result.
     *
     * @param formula the formula, with or without a leading {@code =}
     * @param result  the number the formula gives
     * @return the formula
     * @throws IllegalArgumentException if the formula is blank or the result is not a finite number
     */
    public static ExcelFormula of(String formula, Number result) {
        Objects.requireNonNull(result, "result cannot be null");
        double value = result.doubleValue();
        if (!Double.isFinite(value)) {
            throw new IllegalArgumentException("The result of a formula must be a finite number, got " + result);
        }
        return new ExcelFormula(normalize(formula), value);
    }

    /**
     * A formula with a text result. An empty text is no result: the file does not keep an empty cached result.
     *
     * @param formula the formula, with or without a leading {@code =}
     * @param result  the text the formula gives
     * @return the formula
     * @throws IllegalArgumentException if the formula is blank or the result is longer than a cell holds
     */
    public static ExcelFormula of(String formula, String result) {
        Objects.requireNonNull(result, "result cannot be null");
        if (result.length() > MAX_RESULT_TEXT) {
            throw new IllegalArgumentException("The result of a formula holds at most " + MAX_RESULT_TEXT
                    + " characters, got " + result.length());
        }
        return new ExcelFormula(normalize(formula), result);
    }

    /**
     * A formula with a true/false result.
     *
     * @param formula the formula, with or without a leading {@code =}
     * @param result  the value the formula gives
     * @return the formula
     * @throws IllegalArgumentException if the formula is blank
     */
    public static ExcelFormula of(String formula, boolean result) {
        return new ExcelFormula(normalize(formula), result);
    }

    private static String normalize(String formula) {
        Objects.requireNonNull(formula, "formula cannot be null");
        String text = formula.strip();
        if (text.startsWith("=")) text = text.substring(1).strip();
        if (text.isEmpty()) throw new IllegalArgumentException("formula cannot be blank");
        return text;
    }

    /** The formula, without a leading {@code =}. */
    public String formula() {
        return formula;
    }

    /** The result: a {@code Double}, a {@code String} or a {@code Boolean}; {@code null} when the formula has none. */
    public @Nullable Object result() {
        return result;
    }

    /** The result as text, or the formula when it has none: what a CSV file holds. */
    @Override
    public String toString() {
        return result == null ? formula : String.valueOf(result);
    }

    @Override
    public boolean equals(Object o) {
        return o instanceof ExcelFormula other && formula.equals(other.formula) && Objects.equals(result, other.result);
    }

    @Override
    public int hashCode() {
        return Objects.hash(formula, result);
    }
}
