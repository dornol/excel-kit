package io.github.dornol.excelkit.csv;

import org.jspecify.annotations.Nullable;

/**
 * The settings of one column of a {@link CsvWriter}, given to the {@code .column(name, fn, c -> ...)} overloads.
 * <p>
 * A CSV file has no cell styles, so unlike {@link io.github.dornol.excelkit.excel.ColumnConfig} this holds what a text
 * column can have: the text to write for a missing value.
 *
 * <pre>{@code
 * CsvWriter.<Item>create()
 *     .column("Status", Item::status, c -> c.nullValue("N/A"))
 *     .write(items);
 * }</pre>
 *
 * @author dhkim
 * @since 0.29.0
 */
public final class CsvColumnConfig {
    @Nullable Object nullValue;

    /** Creates the settings of a column with the defaults: a missing value is an empty field. */
    public CsvColumnConfig() {}

    /**
     * Sets what is written when the column's function returns {@code null} (or fails and the writer tolerates it, see
     * {@link CsvWriter#failOnColumnError(boolean)}). Without it a missing value is an empty field.
     * <p>
     * The value goes through the same rules as any other value: it is quoted when it has to be, and with
     * {@link CsvWriter#csvInjectionDefense(boolean) CSV injection defense} on (the default) a text that starts with
     * {@code =}, {@code +}, {@code -} or {@code @} gets a leading single quote, so a placeholder such as {@code "-"} comes
     * out as {@code '-}; use {@code "N/A"} or turn the defense off for trusted data.
     *
     * @param value what to write for a missing value (for example {@code "N/A"} or {@code 0}), or {@code null} for an
     *              empty field
     * @return this configuration for chaining
     */
    public CsvColumnConfig nullValue(@Nullable Object value) {
        this.nullValue = value;
        return this;
    }
}
