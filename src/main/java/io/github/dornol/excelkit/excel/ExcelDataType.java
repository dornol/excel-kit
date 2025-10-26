package io.github.dornol.excelkit.excel;

import java.math.BigDecimal;
import java.time.LocalDate;
import java.time.LocalDateTime;


/**
 * Enum representing supported Excel cell data types.
 * <p>
 * Each type defines how to write a specific Java type into an Excel cell
 * and optionally provides a default number/date format for styling.
 * <p>
 * This enum is used to simplify type-safe column setup in {@link ExcelColumn}.
 *
 * @author dhkim
 * @since 2025-07-19
 */
public enum ExcelDataType {

    /**
     * Generic string type. Calls {@code String.valueOf(value)}.
     */
    STRING((cell, value) -> cell.setCellValue(String.valueOf(value)), null),

    /**
     * Boolean values are converted to "Y" or "N".
     */
    BOOLEAN_TO_YN((cell, value) -> cell.setCellValue(Boolean.TRUE.equals(value) ? "Y" : "N"), null),

    /**
     * Long integer value.
     */
    LONG((cell, value) -> cell.setCellValue((Long) value), ExcelDataFormat.NUMBER.getFormat()),

    /**
     * Integer value.
     */
    INTEGER((cell, value) -> cell.setCellValue((Integer) value), ExcelDataFormat.NUMBER.getFormat()),

    /**
     * Double value with 2 decimal places.
     */
    DOUBLE((cell, value) -> cell.setCellValue((Double) value), ExcelDataFormat.NUMBER_2.getFormat()),

    /**
     * Float value with 2 decimal places.
     */
    FLOAT((cell, value) -> cell.setCellValue((Float) value), ExcelDataFormat.NUMBER_2.getFormat()),

    /**
     * Double value interpreted as a percentage (e.g. 0.25 → 25%).
     */
    DOUBLE_PERCENT((cell, value) -> cell.setCellValue((Double) value), ExcelDataFormat.PERCENT.getFormat()),

    /**
     * Float value interpreted as a percentage.
     */
    FLOAT_PERCENT((cell, value) -> cell.setCellValue((Float) value), ExcelDataFormat.PERCENT.getFormat()),

    /**
     * LocalDateTime formatted as "yyyy-MM-dd HH:mm:ss".
     */
    DATETIME((cell, value) -> cell.setCellValue((LocalDateTime) value), ExcelDataFormat.DATETIME.getFormat()),

    /**
     * LocalDate formatted as "yyyy-MM-dd".
     */
    DATE((cell, value) -> cell.setCellValue((LocalDate) value), ExcelDataFormat.DATE.getFormat()),

    /**
     * LocalDateTime formatted as "HH:mm:ss".
     */
    TIME((cell, value) -> cell.setCellValue((LocalDateTime) value), ExcelDataFormat.TIME.getFormat()),

    /**
     * BigDecimal converted to double (2 decimal places).
     */
    BIG_DECIMAL_TO_DOUBLE((cell, value) -> cell.setCellValue(((BigDecimal) value).doubleValue()), ExcelDataFormat.NUMBER_2.getFormat()),

    /**
     * BigDecimal converted to long (no decimal).
     */
    BIG_DECIMAL_TO_LONG((cell, value) -> cell.setCellValue(((BigDecimal) value).longValue()), LONG.defaultFormat)
    ;

    private final ExcelColumnSetter setter;
    private final String defaultFormat;

    ExcelDataType(ExcelColumnSetter setter, String defaultFormat) {
        this.setter = setter;
        this.defaultFormat = defaultFormat;
    }

    /**
     * Returns the column setter function used to write this type into a cell.
     */
    ExcelColumnSetter getSetter() {
        return setter;
    }

    /**
     * Returns the default Excel format string for this type, or null if none.
     */
    String getDefaultFormat() {
        return defaultFormat;
    }
}
