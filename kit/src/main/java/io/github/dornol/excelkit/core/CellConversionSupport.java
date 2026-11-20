package io.github.dornol.excelkit.core;

import org.jspecify.annotations.Nullable;

import java.text.DecimalFormatSymbols;
import java.text.NumberFormat;
import java.text.ParseException;
import java.text.ParsePosition;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.LocalTime;
import java.time.format.DateTimeFormatter;
import java.time.format.DateTimeParseException;
import java.util.List;
import java.util.Locale;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/** Stateless conversion engine used by the CellData facade. */
final class CellConversionSupport {
    private static final Pattern CURRENCY_SYMBOLS = Pattern.compile("[$₩€%원]");
    private static final Pattern EXPONENT_MARKER = Pattern.compile("[Ee](?:\\+(?=\\p{Nd})|(?!\\+))");

    private CellConversionSupport() {
    }

    static @Nullable Number number(String value, Locale locale) throws ParseException {
        if (value.isBlank()) return null;

        String cleaned = cleanNumber(value, locale, false);
        NumberFormat format = NumberFormat.getNumberInstance(locale);
        ParsePosition position = new ParsePosition(0);
        Number parsed = format.parse(cleaned, position);
        if (parsed != null && position.getIndex() == cleaned.length()) return parsed;

        // NumberFormat expects the localized exponent separator and omits a positive exponent's plus sign.
        String exponentSeparator = DecimalFormatSymbols.getInstance(locale).getExponentSeparator();
        String scientific = EXPONENT_MARKER.matcher(cleaned.replace(exponentSeparator, "E"))
                .replaceAll(Matcher.quoteReplacement(exponentSeparator));
        if (!scientific.equals(cleaned)) {
            position.setIndex(0);
            parsed = format.parse(scientific, position);
            if (parsed != null && position.getIndex() == scientific.length()) return parsed;
        }
        throw new ParseException("Invalid number: " + value, position.getIndex());
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
