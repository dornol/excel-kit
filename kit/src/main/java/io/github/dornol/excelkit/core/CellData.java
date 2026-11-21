package io.github.dornol.excelkit.core;

import org.jspecify.annotations.Nullable;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.math.BigDecimal;
import java.text.ParseException;
import java.time.DateTimeException;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.LocalTime;
import java.time.ZoneId;
import java.time.ZonedDateTime;
import java.time.format.DateTimeFormatter;
import java.time.format.DateTimeParseException;
import java.util.List;
import java.util.Locale;
import java.util.Objects;
import java.util.OptionalDouble;
import java.util.OptionalInt;
import java.util.OptionalLong;
import java.util.function.Function;

/**
 * Represents a single cell's value read from an Excel file,
 * along with its column index and formatted string content.
 * <p>
 * Provides utility methods to convert the cell's value into various Java types,
 * including number, string, boolean, date/time, etc.
 *
 * @author dhkim
 * @since 2025-07-19
 */
public final class CellData {
    private static final Logger log = LoggerFactory.getLogger(CellData.class);
    private final int columnIndex;
    private final String formattedValue;
    private final CellConversionConfig conversionConfig;
    private final @Nullable String rawNumericValue;

    /**
     * Creates cell data using the built-in {@linkplain CellConversionConfig#defaults() conversion defaults}.
     *
     * @param columnIndex    the column index of the cell (0-based)
     * @param formattedValue the formatted string value extracted from Excel
     */
    public CellData(int columnIndex, @Nullable String formattedValue) {
        this(columnIndex, formattedValue, null);
    }

    /**
     * Creates cell data with reader-scoped conversion settings.
     *
     * @param columnIndex      the column index of the cell (0-based)
     * @param formattedValue   the formatted string value extracted from Excel
     * @param conversionConfig conversion settings, or {@code null} to use {@link CellConversionConfig#defaults()}
     * @since 0.19.0
     */
    public CellData(int columnIndex, @Nullable String formattedValue,
                    @Nullable CellConversionConfig conversionConfig) {
        this(columnIndex, formattedValue, conversionConfig, null);
    }

    /**
     * Creates cell data with the original numeric text from an XLSX cell, when available.
     *
     * @param columnIndex      the column index of the cell (0-based)
     * @param formattedValue   the formatted string value
     * @param conversionConfig conversion settings, or {@code null} to use {@link CellConversionConfig#defaults()}
     * @param rawNumericValue  the stored numeric value in locale-independent notation, or {@code null}
     */
    public CellData(int columnIndex, @Nullable String formattedValue,
                    @Nullable CellConversionConfig conversionConfig, @Nullable String rawNumericValue) {
        if (formattedValue == null) {
            formattedValue = "";
        }
        if (columnIndex < 0) {
            throw new IllegalArgumentException("columnIndex must be non-negative");
        }
        this.columnIndex = columnIndex;
        this.formattedValue = formattedValue;
        this.conversionConfig = conversionConfig != null ? conversionConfig : CellConversionConfig.defaults();
        this.rawNumericValue = rawNumericValue;
    }

    public int columnIndex() {
        return columnIndex;
    }

    public String formattedValue() {
        return formattedValue;
    }

    String sourceValue() {
        return rawNumericValue == null ? formattedValue : rawNumericValue;
    }

    /** Text used by numeric conversions: the stored XLSX number when present, otherwise the displayed text. */
    private String numericText() {
        return rawNumericValue == null || rawNumericValue.isBlank() ? formattedValue : rawNumericValue.trim();
    }

    /** Stored XLSX numbers use locale-independent notation; displayed text follows the configured locale. */
    private Locale numericLocale(Locale locale) {
        return rawNumericValue == null || rawNumericValue.isBlank() ? locale : Locale.ROOT;
    }

    /**
     * Returns the number stored in an XLSX cell without applying its display format.
     * For formulas this is the cached numeric result. Dates and times return their Excel
     * serial values, and percentages return the stored fraction.
     *
     * @return the stored number, or {@code null} when no XLSX numeric value is available
     * @throws NumberFormatException if the stored numeric text is invalid
     */
    public @Nullable BigDecimal rawNumber() {
        return rawNumericValue == null ? null : new BigDecimal(rawNumericValue.trim());
    }

    private Locale effectiveLocale() {
        return conversionConfig.locale();
    }

    private List<DateTimeFormatter> effectiveDateFormats() {
        return conversionConfig.dateFormats();
    }

    private List<DateTimeFormatter> effectiveDateTimeFormats() {
        return conversionConfig.dateTimeFormats();
    }

    /**
     * Parses the value as a {@link Number} using the given locale.
     * Recognizes locale-specific grouping, a leading or trailing currency marker, and a trailing percent sign.
     * Returns {@code null} if the value is empty or blank.
     *
     * @param locale the locale to use for number formatting (e.g. {@code Locale.KOREA})
     * @return parsed number, or {@code null} if empty or blank
     * @throws IllegalArgumentException if parsing fails
     */
    public @Nullable Number asNumber(Locale locale) {
        String text = numericText();
        if (text.isBlank()) {
            return null;
        }

        try {
            return CellConversionSupport.number(text, numericLocale(locale));
        } catch (ParseException | NumberFormatException e) {
            log.warn("Failed to parse number (col {}): '{}'", columnIndex, formattedValue);
            throw new IllegalArgumentException("Failed to parse number: " + formattedValue, e);
        }
    }

    /**
     * Parses the value as a {@link Number} using the locale of this cell's {@link CellConversionConfig}.
     * Returns {@code null} if the value is empty or blank.
     *
     * @return parsed number, or {@code null} if blank
     * @see CellConversionConfig#locale()
     */
    public @Nullable Number asNumber() {
        return asNumber(effectiveLocale());
    }

    /**
     * Converts the value to {@link Long}.
     * Returns {@code null} if the value is empty or blank.
     * Fractions, non-finite values, and values outside the long range are rejected.
     *
     * @return the long value, or {@code null} if blank
     */
    public @Nullable Long asLong() {
        try {
            return CellConversionSupport.longNumber(numericText(), numericLocale(effectiveLocale()));
        } catch (ParseException | NumberFormatException | ArithmeticException e) {
            throw new IllegalArgumentException("Failed to convert to long: " + numericText(), e);
        }
    }

    /** Returns the parsed long value as an empty optional for blank cells. */
    public OptionalLong asOptionalLong() {
        Long value = asLong();
        return value == null ? OptionalLong.empty() : OptionalLong.of(value);
    }

    /**
     * Converts the value to {@link Integer}.
     * Returns {@code null} if the value is empty or blank.
     * Throws if the long value is out of int range.
     *
     * @return the integer value, or {@code null} if blank
     */
    public @Nullable Integer asInt() {
        Long longValue = asLong();
        if (longValue == null) {
            return null;
        }
        if (longValue > Integer.MAX_VALUE || longValue < Integer.MIN_VALUE) {
            throw new IllegalArgumentException("Value out of range: " + longValue);
        }
        return longValue.intValue();
    }

    /** Returns the parsed integer value as an empty optional for blank cells. */
    public OptionalInt asOptionalInt() {
        Integer value = asInt();
        return value == null ? OptionalInt.empty() : OptionalInt.of(value);
    }

    /**
     * Returns the raw string as-is.
     *
     * @return the formatted string value
     */
    public String asString() {
        return formattedValue;
    }

    /**
     * Converts the value to a boolean.
     * <p>
     * Accepts "true", "1", "y", "yes" (case-insensitive) as {@code true}.
     * Returns {@code false} if the value is empty, blank, or does not match any recognized true value.
     *
     * @return {@code true} if the value represents a true-like string, otherwise {@code false}
     */
    public boolean asBoolean() {
        if (formattedValue.isBlank()) {
            return false;
        }
        return CellConversionSupport.booleanValue(formattedValue);
    }

    /**
     * Converts the value to a nullable {@link Boolean}.
     * <p>
     * Returns {@code null} if the value is empty or blank, allowing callers to
     * distinguish between "no value" and an explicit {@code false}.
     * Accepts "true", "1", "y", "yes" (case-insensitive) as {@code true};
     * all other non-blank values are treated as {@code false}.
     *
     * @return {@code Boolean.TRUE}, {@code Boolean.FALSE}, or {@code null} if blank
     */
    public @Nullable Boolean asBooleanOrNull() {
        if (formattedValue.isBlank()) {
            return null;
        }
        return CellConversionSupport.booleanValue(formattedValue);
    }

    /**
     * Converts the formatted value to {@link LocalDateTime} using multiple supported date-time formats.
     * Returns {@code null} if the value is empty or blank.
     *
     * Supported formats include:
     * - yyyy-MM-dd[ HH:mm[:ss]]
     * - yyyy/MM/dd[ HH:mm[:ss]]
     * - MM/dd/yy[ HH:mm[:ss]]
     * - M/d/yy[ HH:mm[:ss]]
     * - ISO_LOCAL_DATE_TIME
     *
     * The patterns support optional sections for time (hours, minutes, seconds).
     * If all patterns fail, a {@link DateTimeParseException} will be thrown.
     *
     * @return the parsed date-time, or {@code null} if blank
     */
    public @Nullable LocalDateTime asLocalDateTime() {
        return CellConversionSupport.dateTime(formattedValue, effectiveDateTimeFormats());
    }

    /**
     * Converts the value to {@link LocalDateTime} using the specified format.
     * Returns {@code null} if the value is empty or blank.
     *
     * @param format the date-time pattern (e.g., "yyyy-MM-dd HH:mm:ss")
     * @return the parsed date-time, or {@code null} if blank
     */
    public @Nullable LocalDateTime asLocalDateTime(String format) {
        if (formattedValue.isBlank()) {
            return null;
        }
        return LocalDateTime.parse(formattedValue, CellConversionSupport.strictDatePattern(format));
    }

    /**
     * Converts the formatted value to {@link LocalDate} using multiple supported date formats.
     * Returns {@code null} if the value is empty or blank.
     *
     * Supported formats include:
     * - yyyy-MM-dd
     * - yyyy/MM/dd
     * - MM/dd/yy
     * - M/d/yy
     * - ISO_LOCAL_DATE
     *
     * If all patterns fail, a {@link DateTimeParseException} will be thrown.
     *
     * @return the parsed date, or {@code null} if blank
     */
    public @Nullable LocalDate asLocalDate() {
        return CellConversionSupport.date(formattedValue, effectiveDateFormats());
    }

    /**
     * Converts the value to {@link LocalDate} using the specified format.
     * Returns {@code null} if the value is empty or blank.
     *
     * @param format the date pattern (e.g., "yyyy/MM/dd")
     * @return the parsed date, or {@code null} if blank
     */
    public @Nullable LocalDate asLocalDate(String format) {
        if (formattedValue.isBlank()) {
            return null;
        }
        return LocalDate.parse(formattedValue, CellConversionSupport.strictDatePattern(format));
    }

    /**
     * Converts the value to {@link LocalTime} using ISO format (HH:mm:ss).
     * Returns {@code null} if the value is empty or blank.
     *
     * @return the parsed time, or {@code null} if blank
     */
    public @Nullable LocalTime asLocalTime() {
        return CellConversionSupport.time(formattedValue, null);
    }

    /**
     * Converts the value to {@link LocalTime} using the specified format.
     * Returns {@code null} if the value is empty or blank.
     *
     * @param format the time pattern (e.g., "HH:mm")
     * @return the parsed time, or {@code null} if blank
     */
    public @Nullable LocalTime asLocalTime(String format) {
        return CellConversionSupport.time(formattedValue, format);
    }

    /**
     * Converts the value to {@link ZonedDateTime} by parsing as {@link LocalDateTime}
     * and attaching the given time zone.
     * Uses {@link LocalDateTime#atZone(ZoneId)} to resolve gaps and overlaps.
     * Returns {@code null} if the value is empty or blank.
     *
     * @param zone the time zone to apply
     * @return the zoned date-time, or {@code null} if blank
     * @since 0.9.2
     */
    public @Nullable ZonedDateTime asZonedDateTime(ZoneId zone) {
        LocalDateTime ldt = asLocalDateTime();
        return ldt != null ? ldt.atZone(zone) : null;
    }

    /**
     * Converts the value to {@link ZonedDateTime} using the specified format and time zone.
     * Uses {@link LocalDateTime#atZone(ZoneId)} to resolve gaps and overlaps.
     * Returns {@code null} if the value is empty or blank.
     *
     * @param format the date-time pattern
     * @param zone   the time zone to apply
     * @return the zoned date-time, or {@code null} if blank
     * @since 0.9.2
     */
    public @Nullable ZonedDateTime asZonedDateTime(String format, ZoneId zone) {
        LocalDateTime ldt = asLocalDateTime(format);
        return ldt != null ? ldt.atZone(zone) : null;
    }

    /**
     * Attaches a time zone only when the local date-time has exactly one valid offset.
     *
     * @param zone the time zone to apply
     * @return the zoned date-time, or {@code null} if blank
     * @throws DateTimeException if the local date-time is in a gap or overlap
     */
    public @Nullable ZonedDateTime asZonedDateTimeStrict(ZoneId zone) {
        return strictZone(asLocalDateTime(), zone);
    }

    /**
     * Parses with the specified format and attaches a time zone only when there is one valid offset.
     *
     * @param format the date-time pattern
     * @param zone the time zone to apply
     * @return the zoned date-time, or {@code null} if blank
     * @throws DateTimeException if the local date-time is in a gap or overlap
     */
    public @Nullable ZonedDateTime asZonedDateTimeStrict(String format, ZoneId zone) {
        return strictZone(asLocalDateTime(format), zone);
    }

    private static @Nullable ZonedDateTime strictZone(@Nullable LocalDateTime local, ZoneId zone) {
        if (local == null) return null;
        var offsets = zone.getRules().getValidOffsets(local);
        if (offsets.size() != 1) {
            String reason = offsets.isEmpty() ? "does not exist" : "is ambiguous";
            throw new DateTimeException("Local date-time " + local + " " + reason + " in " + zone);
        }
        return ZonedDateTime.ofStrict(local, offsets.get(0), zone);
    }

    /**
     * Converts the value to {@link Double}.
     * Returns {@code null} if the value is empty or blank.
     * Rejects non-finite values and nonzero values that underflow to zero.
     *
     * @return the double value, or {@code null} if blank
     */
    public @Nullable Double asDouble() {
        return asFloatingPoint(false);
    }

    /** Returns the parsed double value as an empty optional for blank cells. */
    public OptionalDouble asOptionalDouble() {
        Double value = asDouble();
        return value == null ? OptionalDouble.empty() : OptionalDouble.of(value);
    }

    /**
     * Converts the value to {@link Float}.
     * Returns {@code null} if the value is empty or blank.
     * Rejects non-finite values and nonzero values that underflow to zero.
     *
     * @return the float value, or {@code null} if blank
     */
    public @Nullable Float asFloat() {
        Double value = asFloatingPoint(true);
        return value != null ? value.floatValue() : null;
    }

    private @Nullable Double asFloatingPoint(boolean singlePrecision) {
        try {
            return CellConversionSupport.floatingNumber(numericText(), numericLocale(effectiveLocale()), singlePrecision);
        } catch (ParseException | NumberFormatException e) {
            throw new IllegalArgumentException("Failed to convert to floating point: " + numericText(), e);
        }
    }

    /**
     * Parses the formatted value as a {@link BigDecimal}.
     * <p>
     * Unlike {@link #asNumber()}, this method parses the cleaned string directly
     * as a {@link BigDecimal}, avoiding intermediate {@link Double} conversion
     * that can lose precision for large or high-precision values.
     * For XLSX numeric cells the stored number is parsed rather than the display text.
     * Returns {@code null} if the value is empty or blank.
     *
     * @return the BigDecimal value, or {@code null} if blank
     * @throws IllegalArgumentException if the value cannot be parsed as a number
     */
    public @Nullable BigDecimal asBigDecimal() {
        String text = numericText();
        if (text.isBlank()) {
            return null;
        }
        try {
            String cleaned = CellConversionSupport.decimalText(text, numericLocale(effectiveLocale()));
            return new BigDecimal(cleaned);
        } catch (NumberFormatException e) {
            log.warn("Failed to parse BigDecimal (col {}): '{}'", columnIndex, text);
            throw new IllegalArgumentException("Failed to parse BigDecimal: " + text, e);
        }
    }

    /**
     * Checks if the value is empty or blank.
     *
     * @return {@code true} if the formatted value is empty or consists only of whitespace
     */
    public boolean isEmpty() {
        return formattedValue.isBlank();
    }

    /**
     * Converts the value to an enum constant by name (case-insensitive).
     * Returns {@code null} if the value is empty or blank.
     *
     * @param enumType the enum class
     * @param <E>      the enum type
     * @return the matching enum constant, or {@code null} if blank
     * @throws IllegalArgumentException if no matching constant is found
     */
    public <E extends Enum<E>> @Nullable E asEnum(Class<E> enumType) {
        if (formattedValue.isBlank()) {
            return null;
        }
        String trimmed = formattedValue.trim();
        for (E constant : enumType.getEnumConstants()) {
            if (constant.name().equalsIgnoreCase(trimmed)) {
                return constant;
            }
        }
        throw new IllegalArgumentException("No enum constant " + enumType.getSimpleName() + " for value: '" + trimmed + "'");
    }

    /**
     * Converts the cell value using a custom conversion function.
     * <p>
     * The function receives the raw string value and returns the converted result.
     * Returns {@code null} if the value is empty or blank.
     *
     * <pre>{@code
     * UUID id = cell.as(UUID::fromString);
     * MyType obj = cell.as(MyType::parse);
     * }</pre>
     *
     * @param converter the conversion function
     * @param <R>       the return type
     * @return the converted value, or {@code null} if blank
     */
    public <R> @Nullable R as(Function<String, R> converter) {
        if (formattedValue.isBlank()) {
            return null;
        }
        return converter.apply(formattedValue);
    }

    /**
     * Converts the cell value to {@link Integer}, returning the given default if blank.
     *
     * @param defaultValue the value to return if the cell is blank
     * @return the parsed integer or the default value
     */
    public int asInt(int defaultValue) {
        Integer value = asInt();
        return value != null ? value : defaultValue;
    }

    /**
     * Converts the cell value to {@link Long}, returning the given default if blank.
     *
     * @param defaultValue the value to return if the cell is blank
     * @return the parsed long or the default value
     */
    public long asLong(long defaultValue) {
        Long value = asLong();
        return value != null ? value : defaultValue;
    }

    /**
     * Converts the cell value to {@link Double}, returning the given default if blank.
     *
     * @param defaultValue the value to return if the cell is blank
     * @return the parsed double or the default value
     */
    public double asDouble(double defaultValue) {
        Double value = asDouble();
        return value != null ? value : defaultValue;
    }

    /**
     * Returns the string value, or the given default if blank.
     *
     * @param defaultValue the value to return if the cell is blank
     * @return the string value or the default value
     */
    public String asString(String defaultValue) {
        return formattedValue.isBlank() ? defaultValue : formattedValue;
    }

    /**
     * Converts the cell value using a custom function, returning the given default if blank.
     *
     * <pre>{@code
     * UUID id = cell.as(UUID::fromString, DEFAULT_UUID);
     * }</pre>
     *
     * @param converter    the conversion function
     * @param defaultValue the value to return if the cell is blank
     * @param <R>          the return type
     * @return the converted value or the default value
     */
    public <R> R as(Function<String, R> converter, R defaultValue) {
        if (formattedValue.isBlank()) {
            return defaultValue;
        }
        return converter.apply(formattedValue);
    }

    @Override
    public boolean equals(Object o) {
        if (this == o) {
            return true;
        }
        if (!(o instanceof CellData cellData)) {
            return false;
        }
        return columnIndex == cellData.columnIndex
                && Objects.equals(formattedValue, cellData.formattedValue)
                && Objects.equals(comparableRawValue(), cellData.comparableRawValue());
    }

    @Override
    public int hashCode() {
        return Objects.hash(columnIndex, formattedValue, comparableRawValue());
    }

    private @Nullable Object comparableRawValue() {
        if (rawNumericValue == null) return null;
        try {
            return new BigDecimal(rawNumericValue.trim()).stripTrailingZeros();
        } catch (NumberFormatException | ArithmeticException e) {
            // Equality remains defined even for a malformed numeric cell.
            return rawNumericValue;
        }
    }

    @Override
    public String toString() {
        return "CellData[columnIndex=" + columnIndex + ", formattedValue=" + formattedValue + "]";
    }

}
