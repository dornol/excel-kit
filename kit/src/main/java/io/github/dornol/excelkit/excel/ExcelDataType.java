package io.github.dornol.excelkit.excel;

import org.apache.poi.common.usermodel.HyperlinkType;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Hyperlink;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

import org.jspecify.annotations.Nullable;

import java.math.BigDecimal;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.LocalTime;


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
    LONG((cell, value) -> cell.setCellValue(((Number) value).longValue()), ExcelDataFormat.NUMBER.getFormat()),

    /**
     * Integer value.
     */
    INTEGER((cell, value) -> cell.setCellValue(((Number) value).intValue()), ExcelDataFormat.NUMBER.getFormat()),

    /**
     * Double value with 2 decimal places.
     */
    DOUBLE((cell, value) -> cell.setCellValue(((Number) value).doubleValue()), ExcelDataFormat.NUMBER_2.getFormat()),

    /**
     * Float value with 2 decimal places.
     */
    FLOAT((cell, value) -> cell.setCellValue(((Number) value).doubleValue()), ExcelDataFormat.NUMBER_2.getFormat()),

    /**
     * Double value interpreted as a percentage (e.g. 0.25 → 25%).
     */
    DOUBLE_PERCENT((cell, value) -> cell.setCellValue(((Number) value).doubleValue()), ExcelDataFormat.PERCENT.getFormat()),

    /**
     * Float value interpreted as a percentage.
     */
    FLOAT_PERCENT((cell, value) -> cell.setCellValue(((Number) value).doubleValue()), ExcelDataFormat.PERCENT.getFormat()),

    /**
     * LocalDateTime formatted as "yyyy-MM-dd HH:mm:ss".
     */
    DATETIME((cell, value) -> cell.setCellValue((LocalDateTime) value), ExcelDataFormat.DATETIME.getFormat()),

    /**
     * LocalDate formatted as "yyyy-MM-dd".
     */
    DATE((cell, value) -> cell.setCellValue((LocalDate) value), ExcelDataFormat.DATE.getFormat()),

    /**
     * LocalTime formatted as "HH:mm:ss".
     */
    TIME((cell, value) -> cell.setCellValue(((LocalTime) value).atDate(LocalDate.EPOCH)), ExcelDataFormat.TIME.getFormat()),

    /**
     * BigDecimal converted to double (2 decimal places).
     */
    BIG_DECIMAL_TO_DOUBLE((cell, value) -> cell.setCellValue(((BigDecimal) value).doubleValue()), ExcelDataFormat.NUMBER_2.getFormat()),

    /**
     * BigDecimal converted to long (no decimal).
     */
    BIG_DECIMAL_TO_LONG((cell, value) -> cell.setCellValue(((BigDecimal) value).longValue()), LONG.defaultFormat),

    /**
     * Formula type. The value is treated as an Excel formula string (without leading '=').
     * <p>
     * Example: {@code "SUM(A2:A100)"} or {@code "AVERAGE(B2:B50)"}
     */
    FORMULA((cell, value) -> cell.setCellFormula(String.valueOf(value)), null),

    /**
     * Hyperlink type. Creates a clickable URL link in the cell.
     * <p>
     * Accepts either a plain {@code String} (used as both URL and label)
     * or an {@link ExcelHyperlink} instance (separate URL and label).
     */
    HYPERLINK((cell, value) -> {
        String url;
        String label;
        if (value instanceof ExcelHyperlink link) {
            url = link.url();
            label = link.label();
        } else {
            url = String.valueOf(value);
            label = url;
        }
        cell.setCellValue(label);
        CreationHelper createHelper = cell.getSheet().getWorkbook().getCreationHelper();
        Hyperlink hyperlink = createHelper.createHyperlink(HyperlinkType.URL);
        hyperlink.setAddress(url);
        cell.setHyperlink(hyperlink);
    }, null),

    /**
     * Image type. Embeds an image in the cell.
     * <p>
     * Accepts an {@link ExcelImage} instance containing the image bytes and type.
     * The image is anchored to the cell and auto-sized.
     */
    IMAGE((cell, value) -> {
        if (!(value instanceof ExcelImage img)) {
            throw new ExcelWriteException(
                    "IMAGE column requires ExcelImage value, but got: " + (value == null ? "null" : value.getClass().getSimpleName()));
        }
        var wb = cell.getSheet().getWorkbook();
        int pictureIdx = wb.addPicture(img.data(), img.imageType());
        var drawing = cell.getSheet().createDrawingPatriarch();
        var anchor = wb.getCreationHelper().createClientAnchor();
        anchor.setCol1(cell.getColumnIndex());
        anchor.setRow1(cell.getRowIndex());
        anchor.setCol2(cell.getColumnIndex() + 1);
        anchor.setRow2(cell.getRowIndex() + 1);
        drawing.createPicture(anchor, pictureIdx);
    }, null),

    /**
     * Rich text type. Creates a cell with mixed formatting (partial bold, italic, colors, etc.).
     * <p>
     * Accepts an {@link ExcelRichText} instance built using its fluent API.
     * If the value is not an {@code ExcelRichText}, it falls back to {@code String.valueOf(value)}.
     */
    RICH_TEXT((cell, value) -> {
        if (value instanceof ExcelRichText rt) {
            SXSSFWorkbook wb = (SXSSFWorkbook) cell.getSheet().getWorkbook();
            cell.setCellValue(rt.toRichTextString(wb, ExcelRichText.getFontCache(wb)));
        } else {
            cell.setCellValue(String.valueOf(value));
        }
    }, null)
    ;

    private final ExcelColumnSetter setter;
    private final @Nullable String defaultFormat;

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
    @Nullable String getDefaultFormat() {
        return defaultFormat;
    }
}
