package io.github.dornol.excelkit.excel;

/**
 * Predefined Excel-compatible data format strings used for number, date, time, and currency formatting.
 * <p>
 * These formats are used to configure {@link org.apache.poi.ss.usermodel.CellStyle} in Excel export.
 * You can apply them via {@link ExcelDataType} or directly when building columns.
 *
 * @author dhkim
 * @since 2025-07-19
 */
public enum ExcelDataFormat {

    /** Number with thousand separator and no decimals (e.g., 1,000) */
    NUMBER("#,##0"),

    /** Number with 1 decimal place (e.g., 1,000.1) */
    NUMBER_1("#,##0.0"),         // 소수점 1자리

    /** Number with 2 decimal places (e.g., 1,000.12) */
    NUMBER_2("#,##0.00"),        // 소수점 2자리

    /** Number with 4 decimal places (e.g., 1,000.1234) */
    NUMBER_4("#,##0.0000"),      // 소수점 4자리

    /** Percent format with 2 decimal places (e.g., 12.34%) */
    PERCENT("0.00%"),

    /** Date-time format (e.g., 2025-07-19 14:23:00) */
    DATETIME("yyyy-mm-dd hh:mm:ss"),

    /** Date only format (e.g., 2025-07-19) */
    DATE("yyyy-mm-dd"),

    /** Time only format (e.g., 14:23:00) */
    TIME("hh:mm:ss"),

    /** Korean currency with "원" suffix (e.g., 1,000원) */
    CURRENCY_KRW("#,##0\"원\""),

    /** US currency with "$" prefix and 2 decimal places (e.g., $1,000.00) */
    CURRENCY_USD("\"$\"#,##0.00"),
    ;

    /** Raw Excel format string */
    private final String format;

    ExcelDataFormat(String format) {
        this.format = format;
    }

    /**
     * Returns the Excel-compatible format string.
     *
     * @return Excel format string (e.g., "#,##0.00", "yyyy-mm-dd", etc.)
     */
    public String getFormat() {
        return format;
    }
}
