package io.github.dornol.excelkit.excel;

import org.apache.poi.common.usermodel.HyperlinkType;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Hyperlink;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import org.jspecify.annotations.Nullable;

import java.math.BigDecimal;
import java.time.Instant;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.LocalTime;
import java.time.OffsetDateTime;
import java.time.ZoneId;
import java.time.ZonedDateTime;
import java.util.Date;


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
    STRING((cell, value) -> ExcelCellText.write(cell, ExcelCellText.textOf(value)), null),

    /**
     * Boolean values are converted to "Y" or "N".
     */
    BOOLEAN_TO_YN((cell, value) -> cell.setCellValue((Boolean) value ? "Y" : "N"), null),

    /**
     * Long integer value.
     */
    LONG((cell, value) -> setLong(cell, value instanceof Long number ? number : decimal((Number) value).longValueExact()),
            ExcelDataFormat.NUMBER.getFormat()),

    /**
     * Integer value.
     */
    INTEGER((cell, value) -> cell.setCellValue(value instanceof Integer number
            ? number : decimal((Number) value).intValueExact()), ExcelDataFormat.NUMBER.getFormat()),

    /**
     * Double value with 2 decimal places.
     */
    DOUBLE((cell, value) -> setNumber(cell, (Number) value), ExcelDataFormat.NUMBER_2.getFormat()),

    /**
     * Float value with 2 decimal places.
     */
    FLOAT((cell, value) -> setNumber(cell, (Number) value), ExcelDataFormat.NUMBER_2.getFormat()),

    /**
     * Double value interpreted as a percentage (e.g. 0.25 → 25%).
     */
    DOUBLE_PERCENT((cell, value) -> setNumber(cell, (Number) value), ExcelDataFormat.PERCENT.getFormat()),

    /**
     * Float value interpreted as a percentage.
     */
    FLOAT_PERCENT((cell, value) -> setNumber(cell, (Number) value), ExcelDataFormat.PERCENT.getFormat()),

    /**
     * LocalDateTime formatted as "yyyy-MM-dd HH:mm:ss".
     */
    DATETIME((cell, value) -> {
        Object temporal = normalizeTemporal(value);
        setDateTime(cell, temporal instanceof LocalDate date ? date.atStartOfDay() : (LocalDateTime) temporal);
    }, ExcelDataFormat.DATETIME.getFormat()),

    /**
     * LocalDate formatted as "yyyy-MM-dd".
     */
    DATE((cell, value) -> setDateTime(cell, ((LocalDate) normalizeTemporal(value)).atStartOfDay()),
            ExcelDataFormat.DATE.getFormat()),

    /**
     * LocalTime formatted as "HH:mm:ss".
     */
    TIME((cell, value) -> setTime(cell, (LocalTime) normalizeTemporal(value)), ExcelDataFormat.TIME.getFormat()),

    /**
     * BigDecimal converted to double (2 decimal places).
     */
    BIG_DECIMAL_TO_DOUBLE((cell, value) -> setNumber(cell, (Number) value), ExcelDataFormat.NUMBER_2.getFormat()),

    /**
     * BigDecimal converted to long (no decimal).
     */
    BIG_DECIMAL_TO_LONG((cell, value) -> setLong(cell, decimal((Number) value).longValueExact()), LONG.defaultFormat),

    /**
     * Formula type. The value is treated as an Excel formula string (without leading '=').
     * <p>
     * Example: {@code "SUM(A2:A100)"} or {@code "AVERAGE(B2:B50)"}
     * <p>
     * <b>Security warning:</b> Do not pass untrusted user input as formula values.
     * DDE formulas (e.g., {@code cmd|'/c calc'}) can execute arbitrary commands
     * when the Excel file is opened. Always use hardcoded or validated formula strings.
     * <p>
     * Note: No runtime guard is applied here because DDE pipe syntax is used by
     * legitimate integrations (e.g., Bloomberg {@code Bloomberg|'USDKRW Curncy'!Last},
     * Reuters). Blocking pipe characters would break these use cases. Input validation
     * is the caller's responsibility.
     */
    FORMULA((cell, value) -> setUncalculatedFormula(cell, String.valueOf(value)), null),

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
        ExcelCellText.write(cell, label);
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
        anchor.setCol2(cell.getColumnIndex() + img.width());
        anchor.setRow2(cell.getRowIndex() + img.height());
        drawing.createPicture(anchor, pictureIdx);
    }, null),

    /**
     * Rich text type. Creates a cell with mixed formatting (partial bold, italic, colors, etc.).
     * <p>
     * Accepts an {@link ExcelRichText} instance built using its fluent API.
     * If the value is not an {@code ExcelRichText}, it falls back to {@code String.valueOf(value)}.
     * <p>
     * The formatting is kept only when the workbook has shared strings
     * ({@code ExcelWriter.create(o -> o.useSharedStrings(true))}); otherwise the write fails, or, with
     * {@link ExcelWriteErrorPolicy#LENIENT}, writes the plain text.
     */
    RICH_TEXT((cell, value) -> {
        if (value instanceof ExcelRichText rt) {
            writeRichText(cell, rt, cell.getSheet().getWorkbook());
        } else {
            ExcelCellText.write(cell, ExcelCellText.textOf(value));
        }
    }, null)
    ;

    private static final int EXCEL_SIGNIFICANT_DIGITS = 15;
    private static final long FIRST_SIXTEEN_DIGIT_INTEGER = 1_000_000_000_000_000L;
    private static final LocalDate START_1900 = LocalDate.of(1900, 1, 1);
    private static final LocalDate START_1904 = LocalDate.of(1904, 1, 1);
    private static final LocalDate LAST_EXCEL_DATE = LocalDate.of(9999, 12, 31);
    private static final long NANOS_PER_DAY = 86_400_000_000_000L;
    private static final int NANOS_PER_MILLISECOND = 1_000_000;
    private static final String DATETIME_WITH_MILLIS = ExcelDataFormat.DATETIME.getFormat() + ".000";
    private static final String TIME_WITH_MILLIS = ExcelDataFormat.TIME.getFormat() + ".000";

    /**
     * Writes rich text into a cell. Text that has no formatting of its own takes the font of the cell's style, so a
     * column's font applies to the cell as it would to plain text.
     */
    static void writeRichText(Cell cell, ExcelRichText rt, SXSSFWorkbook wb) {
        if (!ExcelWorkbookSupport.usesSharedStrings(wb)) {
            throw new IllegalStateException("Rich text keeps its formatting only in a workbook with shared strings; "
                    + "without them POI writes the cell as plain text. Create the writer with "
                    + "ExcelWriter.create(o -> o.useSharedStrings(true)) (or ExcelWorkbook.create(...))");
        }
        XSSFFont base = cell.getCellStyle() instanceof XSSFCellStyle xs
                ? xs.getFont() : null;
        ExcelCellText.writeStreaming(cell, rt.toRichTextString(wb, ExcelRichText.getFontCache(wb), base));
    }

    static void setDateTime(Cell cell, LocalDateTime value) {
        var workbook = cell.getSheet().getWorkbook();
        boolean date1904 = workbook instanceof SXSSFWorkbook streaming ? streaming.getXSSFWorkbook().isDate1904()
                : workbook instanceof XSSFWorkbook xssf && xssf.isDate1904();
        LocalDate date = value.toLocalDate();
        if (date.isBefore(date1904 ? START_1904 : START_1900) || date.isAfter(LAST_EXCEL_DATE)) {
            throw new IllegalArgumentException("Date is outside the Excel date system: " + value);
        }
        if (value.getNano() % NANOS_PER_MILLISECOND != 0) {
            throw new IllegalArgumentException("Date-time exceeds Excel's millisecond precision: " + value);
        }
        cell.setCellValue(value);
    }

    static void setTime(Cell cell, LocalTime value) {
        if (value.getNano() % NANOS_PER_MILLISECOND != 0) {
            throw new IllegalArgumentException("Time exceeds Excel's millisecond precision: " + value);
        }
        double fraction = (double) value.toNanoOfDay() / NANOS_PER_DAY;
        cell.setCellValue(fraction);
    }

    static void setUncalculatedFormula(Cell cell, String formula) {
        cell.setCellFormula(formula);
        // SXSSF leaves the cached value out of the sheet XML for an empty string result.
        cell.setCellValue("");
        var workbook = cell.getSheet().getWorkbook();
        if (!workbook.getForceFormulaRecalculation()) workbook.setForceFormulaRecalculation(true);
    }

    boolean isInferredNumber() {
        return this == INTEGER || this == LONG || this == FLOAT || this == DOUBLE || this == BIG_DECIMAL_TO_DOUBLE;
    }

    private static BigDecimal decimal(Number value) {
        if (value instanceof BigDecimal decimal) return decimal;
        try {
            return new BigDecimal(value.toString());
        } catch (NumberFormatException e) {
            // A custom Number may expose only its numeric conversion methods.
            return BigDecimal.valueOf(value.doubleValue());
        }
    }

    private static void setLong(Cell cell, long value) {
        if (value > -FIRST_SIXTEEN_DIGIT_INTEGER && value < FIRST_SIXTEEN_DIGIT_INTEGER) {
            cell.setCellValue(value);
        } else {
            setNumber(cell, value);
        }
    }

    static void setNumber(Cell cell, Number value) {
        // A float widened to double exposes binary noise (0.1f -> 0.10000000149011612); use its decimal form.
        double numeric = value instanceof Float f && Float.isFinite(f)
                ? Double.parseDouble(Float.toString(f)) : value.doubleValue();
        if (!Double.isFinite(numeric)) throw new ArithmeticException("Non-finite Excel number: " + value);
        if (!(value instanceof Double || value instanceof Float)) {
            BigDecimal original = decimal(value);
            if (original.stripTrailingZeros().precision() > EXCEL_SIGNIFICANT_DIGITS
                    || BigDecimal.valueOf(numeric).compareTo(original) != 0) {
                throw new ArithmeticException("Value cannot be stored exactly as an Excel number: " + value);
            }
        }
        cell.setCellValue(numeric);
    }

    /**
     * Maps the supported legacy and zoned date types onto the local types the setters use.
     * {@code ZonedDateTime} and {@code OffsetDateTime} keep their own wall-clock time; {@code java.util.Date}
     * and its {@code java.sql} subclasses use the JVM default time zone. {@code Instant} is not mapped
     * because its zone cannot be derived from the value. The {@code java.sql} types are matched by name so
     * a runtime without the {@code java.sql} module keeps working.
     */
    static Object normalizeTemporal(Object value) {
        if (value instanceof ZonedDateTime zoned) return zoned.toLocalDateTime();
        if (value instanceof OffsetDateTime offset) return offset.toLocalDateTime();
        if (value instanceof Date date) {
            var zone = ZoneId.systemDefault();
            // java.sql.Date and java.sql.Time do not support toInstant(), so use the epoch milliseconds.
            var zoned = Instant.ofEpochMilli(date.getTime()).atZone(zone);
            if (isSqlType(date, "java.sql.Date")) return zoned.toLocalDate();
            if (isSqlType(date, "java.sql.Time")) return zoned.toLocalTime().withNano(0);
            // Keep the sub-millisecond precision of java.sql.Timestamp.
            return date.toInstant().atZone(zone).toLocalDateTime();
        }
        return value;
    }

    private static boolean isSqlType(Object value, String className) {
        for (Class<?> type = value.getClass(); type != null; type = type.getSuperclass()) {
            if (type.getName().equals(className)) return true;
        }
        return false;
    }

    static @Nullable ExcelDataType infer(@Nullable Object value) {
        if (value == null) return null;
        value = normalizeTemporal(value);
        if (value instanceof String) return STRING;
        if (value instanceof Boolean) return BOOLEAN_TO_YN;
        if (value instanceof Byte || value instanceof Short || value instanceof Integer) return INTEGER;
        if (value instanceof Long) return LONG;
        if (value instanceof Float) return FLOAT;
        if (value instanceof Double) return DOUBLE;
        if (value instanceof BigDecimal) return BIG_DECIMAL_TO_DOUBLE;
        if (value instanceof LocalDateTime) return DATETIME;
        if (value instanceof LocalDate) return DATE;
        if (value instanceof LocalTime) return TIME;
        if (value instanceof ExcelHyperlink) return HYPERLINK;
        if (value instanceof ExcelImage) return IMAGE;
        if (value instanceof ExcelRichText) return RICH_TEXT;
        if (value instanceof Number) return DOUBLE;
        return STRING;
    }

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

    @Nullable String getDefaultFormat(Object value) {
        value = normalizeTemporal(value);
        if (this == DATETIME && value instanceof LocalDateTime dateTime && dateTime.getNano() != 0) {
            return DATETIME_WITH_MILLIS;
        }
        if (this == TIME && value instanceof LocalTime time && time.getNano() != 0) {
            return TIME_WITH_MILLIS;
        }
        return defaultFormat;
    }
}
