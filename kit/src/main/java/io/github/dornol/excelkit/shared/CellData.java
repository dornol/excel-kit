package io.github.dornol.excelkit.shared;

import org.jspecify.annotations.Nullable;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.math.BigDecimal;
import java.text.NumberFormat;
import java.text.ParseException;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.LocalTime;
import java.time.format.DateTimeFormatter;
import java.time.format.DateTimeParseException;
import java.util.Collections;
import java.util.List;
import java.util.Locale;
import java.util.concurrent.CopyOnWriteArrayList;
import java.util.function.Function;

/**
 * Represents a single cell's value read from an Excel file,
 * along with its column index and formatted string content.
 * <p>
 * Provides utility methods to convert the cell's value into various Java types,
 * including number, string, boolean, date/time, etc.
 *
 * @param columnIndex    the column index of the cell (0-based)
 * @param formattedValue the formatted string value extracted from Excel
 *
 * @author dhkim
 * @since 2025-07-19
 */
public record CellData(int columnIndex, @Nullable String formattedValue) {
    private static final Logger log = LoggerFactory.getLogger(CellData.class);
    private static volatile Locale defaultLocale = Locale.KOREA;

    /**
     * Returns the default locale used by no-arg number parsing methods.
     *
     * @return The current default locale
     */
    public static Locale getDefaultLocale() {
        return defaultLocale;
    }

    /**
     * Sets the default locale for number parsing.
     * Affects all subsequent calls to {@link #asNumber()}, {@link #asLong()},
     * {@link #asInt()}, {@link #asDouble()}, {@link #asFloat()}, and {@link #asBigDecimal()}.
     * <p>
     * The default value is {@link Locale#KOREA}.
     *
     * @param locale The locale to use as default (must not be null)
     */
    public static void setDefaultLocale(Locale locale) {
        if (locale == null) {
            throw new IllegalArgumentException("locale must not be null");
        }
        defaultLocale = locale;
    }

    private static final List<DateTimeFormatter> DEFAULT_DATE_FORMATS = List.of(
            DateTimeFormatter.ofPattern("yyyy-MM-dd"),
            DateTimeFormatter.ofPattern("yyyy/MM/dd"),
            DateTimeFormatter.ofPattern("MM/dd/yy"),
            DateTimeFormatter.ofPattern("M/d/yy"),
            DateTimeFormatter.ISO_LOCAL_DATE
    );
    private static final List<DateTimeFormatter> DEFAULT_DATETIME_FORMATS = List.of(
            DateTimeFormatter.ofPattern("yyyy-MM-dd[ HH:mm[:ss]]"),
            DateTimeFormatter.ofPattern("yyyy/MM/dd[ HH:mm[:ss]]"),
            DateTimeFormatter.ofPattern("MM/dd/yy[ HH:mm[:ss]]"),
            DateTimeFormatter.ofPattern("M/d/yy[ HH:mm[:ss]]"),
            DateTimeFormatter.ISO_LOCAL_DATE_TIME
    );
    private static volatile List<DateTimeFormatter> dateFormatPatterns = new CopyOnWriteArrayList<>(DEFAULT_DATE_FORMATS);
    private static volatile List<DateTimeFormatter> dateTimeFormatPatterns = new CopyOnWriteArrayList<>(DEFAULT_DATETIME_FORMATS);

    /**
     * Adds a custom date format pattern for {@link #asLocalDate()}.
     * The pattern is inserted at the beginning of the list so it takes priority over built-in patterns.
     *
     * @param pattern the date pattern (e.g., "dd.MM.yyyy")
     */
    public static void addDateFormat(String pattern) {
        dateFormatPatterns.add(0, DateTimeFormatter.ofPattern(pattern));
    }

    /**
     * Adds a custom date-time format pattern for {@link #asLocalDateTime()}.
     * The pattern is inserted at the beginning of the list so it takes priority over built-in patterns.
     *
     * @param pattern the date-time pattern (e.g., "dd.MM.yyyy HH:mm:ss")
     */
    public static void addDateTimeFormat(String pattern) {
        dateTimeFormatPatterns.add(0, DateTimeFormatter.ofPattern(pattern));
    }

    /**
     * Returns an unmodifiable view of the currently registered date format patterns.
     *
     * @return the list of date format patterns
     */
    public static List<DateTimeFormatter> getDateFormats() {
        return Collections.unmodifiableList(dateFormatPatterns);
    }

    /**
     * Returns an unmodifiable view of the currently registered date-time format patterns.
     *
     * @return the list of date-time format patterns
     */
    public static List<DateTimeFormatter> getDateTimeFormats() {
        return Collections.unmodifiableList(dateTimeFormatPatterns);
    }

    /**
     * Resets the date format patterns to the built-in defaults.
     * Removes any custom patterns previously added via {@link #addDateFormat(String)}.
     */
    public static void resetDateFormats() {
        dateFormatPatterns = new CopyOnWriteArrayList<>(DEFAULT_DATE_FORMATS);
    }

    /**
     * Resets the date-time format patterns to the built-in defaults.
     * Removes any custom patterns previously added via {@link #addDateTimeFormat(String)}.
     */
    public static void resetDateTimeFormats() {
        dateTimeFormatPatterns = new CopyOnWriteArrayList<>(DEFAULT_DATETIME_FORMATS);
    }

    public CellData {
        if (formattedValue == null) {
            formattedValue = "";
        }
        if (columnIndex < 0) {
            throw new IllegalArgumentException("columnIndex must be non-negative");
        }
    }

    /**
     * Parses the value as a {@link Number} using the given locale.
     * This method removes formatting characters such as commas, currency symbols, and percent signs.
     * Returns {@code null} if the value is empty or blank.
     *
     * @param locale the locale to use for number formatting (e.g. {@code Locale.KOREA})
     * @return parsed number, or {@code null} if empty or blank
     * @throws IllegalArgumentException if parsing fails
     */
    public @Nullable Number asNumber(Locale locale) {
        if (formattedValue.isBlank()) {
            return null;
        }

        try {
            // Remove NBSP, currency symbols, and whitespace
            String cleaned = formattedValue
                    .replace("\u00A0", " ")
                    .replaceAll("[$,₩€%원]", "")
                    .replace(" ", "")
                    .trim();

            return NumberFormat.getNumberInstance(locale).parse(cleaned);
        } catch (ParseException e) {
            log.warn("Failed to parse number (col {}): '{}'", columnIndex, formattedValue);
            throw new IllegalArgumentException("Failed to parse number: " + formattedValue);
        }
    }

    /**
     * Parses the value as a {@link Number} using the configured default locale.
     * Returns {@code null} if the value is empty or blank.
     *
     * @see #setDefaultLocale(Locale)
     */
    public @Nullable Number asNumber() {
        return asNumber(defaultLocale);
    }

    /**
     * Converts the value to {@link Long}.
     * Returns {@code null} if the value is empty or blank.
     */
    public @Nullable Long asLong() {
        Number number = asNumber();
        return number != null ? number.longValue() : null;
    }

    /**
     * Converts the value to {@link Integer}.
     * Returns {@code null} if the value is empty or blank.
     * Throws if the long value is out of int range.
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

    /**
     * Returns the raw string as-is.
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
        return isTrueValue(formattedValue);
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
        return isTrueValue(formattedValue);
    }

    private static boolean isTrueValue(String value) {
        String val = value.trim().toLowerCase();
        return val.equals("true") || val.equals("1") || val.equals("y") || val.equals("yes");
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
     */
    public @Nullable LocalDateTime asLocalDateTime() {
        if (formattedValue.isBlank()) {
            return null;
        }

        for (var formatter : dateTimeFormatPatterns) {
            try {
                return LocalDateTime.parse(formattedValue, formatter);
            } catch (Exception e) {
                /* skip */
            }
        }

        throw new DateTimeParseException("Cannot parse LocalDateTime: " + formattedValue, formattedValue, 0);
    }

    /**
     * Converts the value to {@link LocalDateTime} using the specified format.
     * Returns {@code null} if the value is empty or blank.
     *
     * @param format the date-time pattern (e.g., "yyyy-MM-dd HH:mm:ss")
     */
    public @Nullable LocalDateTime asLocalDateTime(String format) {
        if (formattedValue.isBlank()) {
            return null;
        }
        return LocalDateTime.parse(formattedValue, DateTimeFormatter.ofPattern(format));
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
     */
    public @Nullable LocalDate asLocalDate() {
        if (formattedValue.isBlank()) {
            return null;
        }
        for (var format : dateFormatPatterns) {
            try {
                return LocalDate.parse(formattedValue, format);
            } catch (Exception e) {
                /* skip */
            }
        }
        return LocalDate.parse(formattedValue);
    }

    /**
     * Converts the value to {@link LocalDate} using the specified format.
     * Returns {@code null} if the value is empty or blank.
     *
     * @param format the date pattern (e.g., "yyyy/MM/dd")
     */
    public @Nullable LocalDate asLocalDate(String format) {
        if (formattedValue.isBlank()) {
            return null;
        }
        return LocalDate.parse(formattedValue, DateTimeFormatter.ofPattern(format));
    }

    /**
     * Converts the value to {@link LocalTime} using ISO format (HH:mm:ss).
     * Returns {@code null} if the value is empty or blank.
     */
    public @Nullable LocalTime asLocalTime() {
        if (formattedValue.isBlank()) {
            return null;
        }
        return LocalTime.parse(formattedValue);
    }

    /**
     * Converts the value to {@link LocalTime} using the specified format.
     * Returns {@code null} if the value is empty or blank.
     *
     * @param format the time pattern (e.g., "HH:mm")
     */
    public @Nullable LocalTime asLocalTime(String format) {
        if (formattedValue.isBlank()) {
            return null;
        }
        return LocalTime.parse(formattedValue, DateTimeFormatter.ofPattern(format));
    }

    /**
     * Converts the value to {@link Double}.
     * Returns {@code null} if the value is empty or blank.
     */
    public @Nullable Double asDouble() {
        Number number = asNumber();
        return number != null ? number.doubleValue() : null;
    }

    /**
     * Converts the value to {@link Float}.
     * Returns {@code null} if the value is empty or blank.
     */
    public @Nullable Float asFloat() {
        Number number = asNumber();
        return number != null ? number.floatValue() : null;
    }

    /**
     * Converts the value to {@link BigDecimal}.
     * Uses the string representation of the parsed number to avoid precision loss.
     * Returns {@code null} if the value is empty or blank.
     */
    public @Nullable BigDecimal asBigDecimal() {
        Number number = asNumber();
        return number != null ? new BigDecimal(number.toString()) : null;
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

}
