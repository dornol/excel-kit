package io.github.dornol.excelkit.excel.style;



/**
 * Predefined Excel-compatible data format strings used for number, date, time, and currency formatting.
 * <p>
 * These formats are used to configure {@link org.apache.poi.ss.usermodel.CellStyle} in Excel export.
 * You can apply them via {@link io.github.dornol.excelkit.excel.ExcelDataType} or directly when building columns.
 *
 * @author dhkim
 * @since 0.1.2
 */
public enum ExcelDataFormat {

    /**
     * Text. The cell keeps what it holds and what a person types into it as text: Excel does not turn {@code 0012}
     * into the number 12 when the cell is edited (it does in a cell with the General format). Use it for ids, postal
     * codes and phone numbers with leading zeros, on text columns only: a number in a text cell is no longer
     * added up or sorted as a number once edited.
     */
    TEXT("@"),

    /** Whole number without a thousand separator (e.g., 1000) */
    NUMBER_PLAIN("0"),

    /** Number with thousand separator and no decimals (e.g., 1,000) */
    NUMBER("#,##0"),

    /** Number with 1 decimal place (e.g., 1,000.1) */
    NUMBER_1("#,##0.0"),         // 1 decimal place

    /** Number with 2 decimal places (e.g., 1,000.12) */
    NUMBER_2("#,##0.00"),        // 2 decimal places

    /** Number with 4 decimal places (e.g., 1,000.1234) */
    NUMBER_4("#,##0.0000"),      // 4 decimal places

    /** Percent format without decimals (e.g., 12%) */
    PERCENT_0("0%"),

    /** Percent format with 1 decimal place (e.g., 12.3%) */
    PERCENT_1("0.0%"),

    /** Percent format with 2 decimal places (e.g., 12.34%) */
    PERCENT("0.00%"),

    /** Date-time format (e.g., 2025-07-19 14:23:00) */
    DATETIME("yyyy-mm-dd hh:mm:ss"),

    /** Date only format (e.g., 2025-07-19) */
    DATE("yyyy-mm-dd"),

    /** Date and time to the minute (e.g., 2025-07-19 14:23) */
    DATETIME_HM("yyyy-mm-dd hh:mm"),

    /** Date with dots (e.g., 2025.07.19) */
    DATE_DOT("yyyy.mm.dd"),

    /** Date with slashes (e.g., 2025/07/19) */
    DATE_SLASH("yyyy/mm/dd"),

    /** Korean date (e.g., 2025년 7월 19일) */
    DATE_KR("yyyy\"년\" m\"월\" d\"일\""),

    /** Year and month (e.g., 2025-07) */
    YEAR_MONTH("yyyy-mm"),

    /** Time only format (e.g., 14:23:00) */
    TIME("hh:mm:ss"),

    /** Time to the minute (e.g., 14:23) */
    TIME_HM("hh:mm"),

    /** Korean currency with "원" suffix (e.g., 1,000원) */
    CURRENCY_KRW("#,##0\"원\""),

    /** US currency with "$" prefix and 2 decimal places (e.g., $1,000.00) */
    CURRENCY_USD("\"$\"#,##0.00"),

    /** Euro with "€" prefix and 2 decimal places (e.g., €1,000.00) */
    CURRENCY_EUR("\"€\"#,##0.00"),

    /** Japanese yen with "¥" prefix and no decimals (e.g., ¥1,000) */
    CURRENCY_JPY("\"¥\"#,##0"),
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
