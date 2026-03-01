package io.github.dornol.excelkit.shared;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import org.jspecify.annotations.NonNull;

import java.math.BigDecimal;
import java.text.NumberFormat;
import java.text.ParseException;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.LocalTime;
import java.time.format.DateTimeFormatter;
import java.time.format.DateTimeParseException;
import java.util.ArrayList;
import java.util.Collections;
import java.util.List;
import java.util.Locale;

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
public record CellData(int columnIndex, String formattedValue) {
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
    public static void setDefaultLocale(@NonNull Locale locale) {
        if (locale == null) {
            throw new IllegalArgumentException("locale must not be null");
        }
        defaultLocale = locale;
    }

    private static final List<DateTimeFormatter> DATE_FORMAT_PATTERNS = new ArrayList<>(List.of(
            DateTimeFormatter.ofPattern("yyyy-MM-dd"),
            DateTimeFormatter.ofPattern("yyyy/MM/dd"),
            DateTimeFormatter.ofPattern("MM/dd/yy"),
            DateTimeFormatter.ofPattern("M/d/yy"),
            DateTimeFormatter.ISO_LOCAL_DATE
    ));
    private static final List<DateTimeFormatter> DATETIME_FORMAT_PATTERNS = new ArrayList<>(List.of(
            DateTimeFormatter.ofPattern("yyyy-MM-dd[ HH:mm[:ss]]"),
            DateTimeFormatter.ofPattern("yyyy/MM/dd[ HH:mm[:ss]]"),
            DateTimeFormatter.ofPattern("MM/dd/yy[ HH:mm[:ss]]"),
            DateTimeFormatter.ofPattern("M/d/yy[ HH:mm[:ss]]"),
            DateTimeFormatter.ISO_LOCAL_DATE_TIME
    ));

    /**
     * Adds a custom date format pattern for {@link #asLocalDate()}.
     * The pattern is inserted at the beginning of the list so it takes priority over built-in patterns.
     *
     * @param pattern the date pattern (e.g., "dd.MM.yyyy")
     */
    public static void addDateFormat(@NonNull String pattern) {
        DATE_FORMAT_PATTERNS.addFirst(DateTimeFormatter.ofPattern(pattern));
    }

    /**
     * Adds a custom date-time format pattern for {@link #asLocalDateTime()}.
     * The pattern is inserted at the beginning of the list so it takes priority over built-in patterns.
     *
     * @param pattern the date-time pattern (e.g., "dd.MM.yyyy HH:mm:ss")
     */
    public static void addDateTimeFormat(@NonNull String pattern) {
        DATETIME_FORMAT_PATTERNS.addFirst(DateTimeFormatter.ofPattern(pattern));
    }

    /**
     * Returns an unmodifiable view of the currently registered date format patterns.
     *
     * @return the list of date format patterns
     */
    public static List<DateTimeFormatter> getDateFormats() {
        return Collections.unmodifiableList(DATE_FORMAT_PATTERNS);
    }

    /**
     * Returns an unmodifiable view of the currently registered date-time format patterns.
     *
     * @return the list of date-time format patterns
     */
    public static List<DateTimeFormatter> getDateTimeFormats() {
        return Collections.unmodifiableList(DATETIME_FORMAT_PATTERNS);
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
    public Number asNumber(@NonNull Locale locale) {
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
    public Number asNumber() {
        return asNumber(defaultLocale);
    }

    /**
     * Converts the value to {@link Long}.
     * Returns {@code null} if the value is empty or blank.
     */
    public Long asLong() {
        Number number = asNumber();
        return number != null ? number.longValue() : null;
    }

    /**
     * Converts the value to {@link Integer}.
     * Returns {@code null} if the value is empty or blank.
     * Throws if the long value is out of int range.
     */
    public Integer asInt() {
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
    public Boolean asBooleanOrNull() {
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
    public LocalDateTime asLocalDateTime() {
        if (formattedValue.isBlank()) {
            return null;
        }

        for (var formatter : DATETIME_FORMAT_PATTERNS) {
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
    public LocalDateTime asLocalDateTime(String format) {
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
    public LocalDate asLocalDate() {
        if (formattedValue.isBlank()) {
            return null;
        }
        for (var format : DATE_FORMAT_PATTERNS) {
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
    public LocalDate asLocalDate(String format) {
        if (formattedValue.isBlank()) {
            return null;
        }
        return LocalDate.parse(formattedValue, DateTimeFormatter.ofPattern(format));
    }

    /**
     * Converts the value to {@link LocalTime} using ISO format (HH:mm:ss).
     * Returns {@code null} if the value is empty or blank.
     */
    public LocalTime asLocalTime() {
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
    public LocalTime asLocalTime(String format) {
        if (formattedValue.isBlank()) {
            return null;
        }
        return LocalTime.parse(formattedValue, DateTimeFormatter.ofPattern(format));
    }

    /**
     * Converts the value to {@link Double}.
     * Returns {@code null} if the value is empty or blank.
     */
    public Double asDouble() {
        Number number = asNumber();
        return number != null ? number.doubleValue() : null;
    }

    /**
     * Converts the value to {@link Float}.
     * Returns {@code null} if the value is empty or blank.
     */
    public Float asFloat() {
        Number number = asNumber();
        return number != null ? number.floatValue() : null;
    }

    /**
     * Converts the value to {@link BigDecimal}.
     * Uses the string representation of the parsed number to avoid precision loss.
     * Returns {@code null} if the value is empty or blank.
     */
    public BigDecimal asBigDecimal() {
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

}
