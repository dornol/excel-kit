package io.github.dornol.excelkit.excel;

import org.jspecify.annotations.Nullable;

import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.Objects;
import java.util.function.Consumer;
import java.util.function.Function;

/**
 * Configures the subtotal rows {@code groupSubtotal(...)} writes below every group of rows.
 * <p>
 * A subtotal is an Excel {@code SUBTOTAL} formula over the rows of its group, so it follows the data when a person
 * edits it, and a {@code summary(...)} over the whole sheet does not count the subtotals twice (Excel's
 * {@code SUBTOTAL} skips other {@code SUBTOTAL} cells).
 *
 * @since 0.29.0
 */
public class ExcelSubtotal {
    private final Map<ExcelSummary.Op, List<String>> entries = new LinkedHashMap<>();
    private @Nullable String labelColumn;
    private @Nullable String keyColumn;
    private Function<@Nullable Object, String> labelText = key -> (key == null ? "" : key + " ") + "Subtotal";
    private final RowStyleConfig style = new RowStyleConfig().bold(true)
            .backgroundColor(io.github.dornol.excelkit.excel.style.ExcelColor.LIGHT_GRAY);

    /** Creates an empty configuration. */
    public ExcelSubtotal() {}

    /**
     * Sets the column that holds the text of a subtotal row and how the text is made from the group's key. By default
     * the first column holds {@code "<key> Subtotal"}. The column must not also be a subtotal column.
     *
     * @param columnName the column of the text
     * @param text       the text of a group, given its key (which may be {@code null})
     * @return this configuration for chaining
     */
    public ExcelSubtotal label(String columnName, Function<@Nullable Object, String> text) {
        this.labelColumn = Objects.requireNonNull(columnName, "columnName cannot be null");
        this.labelText = Objects.requireNonNull(text, "text cannot be null");
        return this;
    }

    /**
     * Writes the group's value into the given column of its subtotal row, as the first row of the group holds it. An
     * auto filter on that column then keeps the subtotal row with the rows of its group: filtering {@code Dept = A}
     * shows the rows of A and its subtotal, where the label alone ({@code "A total"}) would be filtered out. The
     * column must not also hold the label or a subtotal. Without it the subtotal row has nothing in that column.
     *
     * @param columnName the column of the group key, usually the one the rows are grouped by
     * @return this configuration for chaining
     */
    public ExcelSubtotal keyColumn(String columnName) {
        this.keyColumn = Objects.requireNonNull(columnName, "columnName cannot be null");
        return this;
    }

    /**
     * Styles the subtotal rows (bold on light gray by default).
     *
     * @param configurer changes the default style
     * @return this configuration for chaining
     */
    public ExcelSubtotal style(Consumer<RowStyleConfig> configurer) {
        Objects.requireNonNull(configurer, "configurer cannot be null").accept(style);
        return this;
    }

    /** Subtotals the column with {@code SUM}. @param columnName the column @return this configuration */
    public ExcelSubtotal sum(String columnName) {
        return add(ExcelSummary.Op.SUM, columnName);
    }

    /** Subtotals the column with {@code AVERAGE}. @param columnName the column @return this configuration */
    public ExcelSubtotal average(String columnName) {
        return add(ExcelSummary.Op.AVERAGE, columnName);
    }

    /** Subtotals the column with {@code COUNT}. @param columnName the column @return this configuration */
    public ExcelSubtotal count(String columnName) {
        return add(ExcelSummary.Op.COUNT, columnName);
    }

    /**
     * Subtotals the column with {@code COUNTA}: the cells that are not empty, text included (where {@code count}
     * counts only the cells that hold a number).
     *
     * @param columnName the column
     * @return this configuration for chaining
     */
    public ExcelSubtotal countA(String columnName) {
        return add(ExcelSummary.Op.COUNTA, columnName);
    }

    /** Subtotals the column with {@code MIN}. @param columnName the column @return this configuration */
    public ExcelSubtotal min(String columnName) {
        return add(ExcelSummary.Op.MIN, columnName);
    }

    /** Subtotals the column with {@code MAX}. @param columnName the column @return this configuration */
    public ExcelSubtotal max(String columnName) {
        return add(ExcelSummary.Op.MAX, columnName);
    }

    private ExcelSubtotal add(ExcelSummary.Op op, String columnName) {
        Objects.requireNonNull(columnName, "columnName cannot be null");
        entries.computeIfAbsent(op, k -> new ArrayList<>()).add(columnName);
        return this;
    }

    Map<ExcelSummary.Op, List<String>> entries() { return entries; }

    @Nullable String labelColumn() { return labelColumn; }

    @Nullable String keyColumn() { return keyColumn; }

    Function<@Nullable Object, String> labelText() { return labelText; }

    RowStyleConfig style() { return style; }

    /** The first argument of Excel's {@code SUBTOTAL}: the function, counting hidden rows (the 1-11 family). */
    static int functionNumber(ExcelSummary.Op op) {
        return switch (op) {
            case AVERAGE -> 1;
            case COUNT -> 2;
            case COUNTA -> 3;
            case MAX -> 4;
            case MIN -> 5;
            case SUM -> 9;
        };
    }
}
