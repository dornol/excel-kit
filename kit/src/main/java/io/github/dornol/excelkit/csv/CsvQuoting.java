package io.github.dornol.excelkit.csv;

/**
 * Quoting strategies for CSV field values.
 *
 * <pre>{@code
 * new CsvWriter<Item>()
 *     .quoting(CsvQuoting.ALL)
 *     .column("Name", Item::name)
 *     .write(stream);
 * }</pre>
 *
 * @author dhkim
 * @since 0.9.2
 */
public enum CsvQuoting {

    /**
     * Quote only when necessary — when the value contains the delimiter, quotes, or newlines.
     * This is the default behavior.
     */
    MINIMAL,

    /**
     * Quote all fields unconditionally.
     */
    ALL,

    /**
     * Quote fields that are not purely numeric.
     * Numeric values (integer and decimal) are left unquoted.
     */
    NON_NUMERIC
}
