package io.github.dornol.excelkit.core;

import org.jspecify.annotations.Nullable;

import java.text.DecimalFormatSymbols;
import java.text.NumberFormat;
import java.text.ParseException;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.LocalTime;
import java.time.format.DateTimeFormatter;
import java.time.format.DateTimeParseException;
import java.util.List;
import java.util.Locale;
import java.util.regex.Pattern;

/** Stateless conversion engine used by the CellData facade. */
final class CellConversionSupport {
    private static final Pattern CURRENCY_SYMBOLS = Pattern.compile("[$₩€%원]");

    private CellConversionSupport() {
    }

    static @Nullable Number number(String value, Locale locale) throws ParseException {
        return value.isBlank() ? null : NumberFormat.getNumberInstance(locale).parse(cleanNumber(value, locale, false));
    }

    static String decimalText(String value, Locale locale) {
        return cleanNumber(value, locale, true);
    }

    static boolean booleanValue(String value) {
        String normalized = value.trim().toLowerCase(Locale.ROOT);
        return normalized.equals("true") || normalized.equals("1")
                || normalized.equals("y") || normalized.equals("yes");
    }

    static @Nullable LocalDateTime dateTime(String value, List<DateTimeFormatter> formats) {
        if (value.isBlank()) return null;
        for (DateTimeFormatter formatter : formats) {
            try { return LocalDateTime.parse(value, formatter); }
            catch (DateTimeParseException ignored) { }
        }
        throw new DateTimeParseException("Cannot parse LocalDateTime: " + value, value, 0);
    }

    static @Nullable LocalDate date(String value, List<DateTimeFormatter> formats) {
        if (value.isBlank()) return null;
        for (DateTimeFormatter formatter : formats) {
            try { return LocalDate.parse(value, formatter); }
            catch (DateTimeParseException ignored) { }
        }
        return LocalDate.parse(value);
    }

    static @Nullable LocalTime time(String value, @Nullable String pattern) {
        if (value.isBlank()) return null;
        return pattern == null ? LocalTime.parse(value)
                : LocalTime.parse(value, DateTimeFormatter.ofPattern(pattern));
    }

    private static String cleanNumber(String value, Locale locale, boolean normalizeDecimalSeparator) {
        String cleaned = CURRENCY_SYMBOLS.matcher(value.replace("\u00A0", " "))
                .replaceAll("").replace(" ", "").trim();
        DecimalFormatSymbols symbols = DecimalFormatSymbols.getInstance(locale);
        char groupingSeparator = symbols.getGroupingSeparator();
        char decimalSeparator = symbols.getDecimalSeparator();
        if (groupingSeparator != decimalSeparator) cleaned = cleaned.replace(String.valueOf(groupingSeparator), "");
        if (normalizeDecimalSeparator && decimalSeparator != '.') cleaned = cleaned.replace(decimalSeparator, '.');
        return cleaned;
    }
}
