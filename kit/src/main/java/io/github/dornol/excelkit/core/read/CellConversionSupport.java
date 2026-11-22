package io.github.dornol.excelkit.core.read;

import org.jspecify.annotations.Nullable;

import java.math.BigDecimal;
import java.text.DecimalFormat;
import java.text.DecimalFormatSymbols;
import java.text.NumberFormat;
import java.text.ParseException;
import java.text.ParsePosition;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.LocalTime;
import java.time.chrono.IsoEra;
import java.time.format.DateTimeFormatter;
import java.time.format.DateTimeFormatterBuilder;
import java.time.format.DateTimeParseException;
import java.time.format.ResolverStyle;
import java.time.temporal.ChronoField;
import java.util.List;
import java.util.Locale;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/** Stateless conversion engine used by the CellData facade. */
final class CellConversionSupport {
    private static final Pattern LEADING_CURRENCY = Pattern.compile("^([+\\p{Pd}\\p{Cf}]*) *[$₩€] *");
    private static final Pattern TRAILING_NUMBER_AFFIX = Pattern.compile(" *[$₩€%원]$");
    private static final Pattern EXPONENT_MARKER = Pattern.compile("[Ee](?:\\+(?=\\p{Nd})|(?!\\+))");
    private static final Pattern LEADING_NUMBER_SIGN = Pattern.compile("^[+\\p{Pd}\\p{Cf}]*");

    private CellConversionSupport() {
    }

    static @Nullable Number number(String value, Locale locale) throws ParseException {
        Number parsed = number(value, locale, false);
        if (parsed instanceof Double floating) validateFloatingPoint(floating, value, locale);
        return parsed;
    }

    static @Nullable Long longNumber(String value, Locale locale) throws ParseException {
        Number parsed = number(value, locale, true);
        if (parsed == null) return null;
        if (parsed instanceof BigDecimal decimal) return decimal.longValueExact();
        return new BigDecimal(parsed.toString()).longValueExact();
    }

    static @Nullable Double floatingNumber(String value, Locale locale, boolean singlePrecision)
            throws ParseException {
        Number parsed = number(value, locale, singlePrecision);
        if (parsed == null) return null;
        // DecimalFormat's exact mode represents both signed zeros as BigDecimal.ZERO.
        if (singlePrecision && parsed instanceof BigDecimal exact && exact.signum() == 0) {
            parsed = number(value, locale, false);
        }
        double converted = singlePrecision ? parsed.floatValue() : parsed.doubleValue();
        validateFloatingPoint(converted, value, locale);
        return converted;
    }

    private static void validateFloatingPoint(double converted, String value, Locale locale) throws ParseException {
        if (!Double.isFinite(converted)) {
            throw new NumberFormatException("Non-finite floating-point value: " + value);
        }
        if (converted == 0.0 && number(value, locale, true) instanceof BigDecimal exact && exact.signum() != 0) {
            throw new NumberFormatException("Floating-point value underflow: " + value);
        }
    }

    private static @Nullable Number number(String value, Locale locale, boolean exact) throws ParseException {
        if (value.isBlank()) return null;
        Number plain = plainInteger(value, exact);
        if (plain != null) return plain;
        NumberFormat format = numberFormat(locale);
        if (exact && format instanceof DecimalFormat decimalFormat) decimalFormat.setParseBigDecimal(true);
        DecimalFormatSymbols symbols = DecimalFormatSymbols.getInstance(locale);
        String cleaned = cleanNumber(value, symbols, format);
        ParsePosition position = new ParsePosition(0);
        Number parsed = format.parse(cleaned, position);
        if (parsed != null && position.getIndex() == cleaned.length()) return parsed;

        // NumberFormat expects the localized exponent separator and omits a positive exponent's plus sign.
        String exponentSeparator = symbols.getExponentSeparator();
        String scientific = EXPONENT_MARKER.matcher(cleaned.replace(exponentSeparator, "E"))
                .replaceAll(Matcher.quoteReplacement(exponentSeparator));
        if (!scientific.equals(cleaned)) {
            position.setIndex(0);
            parsed = format.parse(scientific, position);
            if (parsed != null && position.getIndex() == scientific.length()) return parsed;
        }
        throw new ParseException("Invalid number: " + value, position.getIndex());
    }

    private static final java.util.concurrent.ConcurrentHashMap<Locale, NumberFormat> NUMBER_FORMATS =
            new java.util.concurrent.ConcurrentHashMap<>();

    /**
     * A number written as plain ASCII digits only means the same in every locale (no sign, separator or
     * affix), so it skips the locale machinery. Returns what the number format would: a {@code Long}, or a
     * {@code BigDecimal} when exact parsing is asked for. Anything else, and anything too long for a
     * {@code long}, returns {@code null} and takes the general path.
     */
    private static @Nullable Number plainInteger(String value, boolean exact) {
        int length = value.length();
        if (length > 18) return null;
        long result = 0;
        for (int i = 0; i < length; i++) {
            char c = value.charAt(i);
            if (c < '0' || c > '9') return null;
            result = result * 10 + (c - '0');
        }
        return exact ? BigDecimal.valueOf(result) : Long.valueOf(result);
    }

    /** A private copy of the locale's number format: building one from its pattern is the costly part. */
    private static NumberFormat numberFormat(Locale locale) {
        return (NumberFormat) NUMBER_FORMATS.computeIfAbsent(locale, NumberFormat::getNumberInstance).clone();
    }

    static String decimalText(String value, Locale locale) {
        DecimalFormatSymbols symbols = DecimalFormatSymbols.getInstance(locale);
        return cleanNumber(value, symbols, NumberFormat.getNumberInstance(locale))
                .replace(symbols.getDecimalSeparator(), '.');
    }

    static DateTimeFormatter strictDatePattern(String pattern) {
        DateTimeFormatterBuilder formatter = new DateTimeFormatterBuilder().appendPattern(pattern);
        if (pattern.indexOf('y') >= 0) formatter.parseDefaulting(ChronoField.ERA, IsoEra.CE.getValue());
        return formatter.toFormatter().withResolverStyle(ResolverStyle.STRICT);
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
                : LocalTime.parse(value, DateTimeFormatter.ofPattern(pattern).withResolverStyle(ResolverStyle.STRICT));
    }

    /** Only a currency sign or a percent sign can be a number's prefix or suffix; most numbers have none. */
    private static boolean hasAffixCharacter(String text) {
        for (int i = 0; i < text.length(); i++) {
            switch (text.charAt(i)) {
                case '$', '\u20A9', '\u20AC', '%', '\uC6D0' -> { return true; }
                default -> { }
            }
        }
        return false;
    }

    private static String cleanNumber(String value, DecimalFormatSymbols symbols, NumberFormat format) {
        String cleaned = value.replace('\u00A0', ' ').replace('\u202F', ' ').strip();
        if (hasAffixCharacter(cleaned)) {
            Matcher leadingCurrency = LEADING_CURRENCY.matcher(cleaned);
            cleaned = leadingCurrency.find() ? leadingCurrency.replaceFirst("$1")
                    : TRAILING_NUMBER_AFFIX.matcher(cleaned).replaceFirst("");
        }
        char groupingSeparator = symbols.getGroupingSeparator();
        if (Character.isSpaceChar(groupingSeparator)) cleaned = cleaned.replace(' ', groupingSeparator);
        int firstSeparator = cleaned.indexOf(groupingSeparator);
        if (groupingSeparator != symbols.getDecimalSeparator() && firstSeparator >= 0) {
            int groupingSize = format instanceof DecimalFormat decimalFormat ? decimalFormat.getGroupingSize() : 3;
            validateGrouping(cleaned, groupingSeparator, firstSeparator, groupingSize);
            cleaned = cleaned.replace(String.valueOf(groupingSeparator), "");
        }
        return cleaned;
    }

    private static void validateGrouping(String value, char separator, int firstSeparator, int groupSize) {
        String leading = LEADING_NUMBER_SIGN.matcher(value.substring(0, firstSeparator)).replaceFirst("");
        if (leading.isEmpty() || leading.length() > groupSize || !leading.chars().allMatch(Character::isDigit)) {
            throw new NumberFormatException("Invalid digit grouping: " + value);
        }
        int lastSeparator = value.lastIndexOf(separator);
        int index = firstSeparator;
        do {
            int start = ++index;
            while (index < value.length() && Character.isDigit(value.charAt(index))) index++;
            if (index - start != groupSize || (index < lastSeparator && value.charAt(index) != separator)) {
                throw new NumberFormatException("Invalid digit grouping: " + value);
            }
        } while (index <= lastSeparator);
    }
}
