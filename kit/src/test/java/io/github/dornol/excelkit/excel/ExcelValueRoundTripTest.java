package io.github.dornol.excelkit.excel;

import io.github.dornol.excelkit.core.read.CellData;
import org.junit.jupiter.api.Test;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.math.BigDecimal;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.LocalTime;
import java.util.ArrayList;
import java.util.List;
import java.util.Objects;
import java.util.function.Function;

import static org.junit.jupiter.api.Assertions.assertTrue;

/**
 * A value written to a cell comes back as the same value when it is read with the accessor of its type. Each
 * case writes one value, as an inferred and as an explicitly typed column where that makes sense, and reads
 * it back. A case that cannot round trip says why in {@code KNOWN_LIMITS}, so the limits of Excel (numbers are
 * doubles, a cell holds 32,767 characters, ...) are written down instead of discovered.
 */
class ExcelValueRoundTripTest {

    /**
     * @param rejection when not {@code null}, Excel cannot hold the value and the write must fail with an
     *                  {@link ExcelWriteException} whose message contains this text
     */
    private record Case(String name, Object value, Function<CellData, Object> read, Object expected, String rejection) {
        static Case of(String name, Object value, Function<CellData, Object> read) {
            return new Case(name, value, read, value, null);
        }

        static Case of(String name, Object value, Function<CellData, Object> read, Object expected) {
            return new Case(name, value, read, expected, null);
        }

        /** A value Excel cannot hold exactly: it is refused when written, not silently rounded. */
        static Case rejected(String name, Object value, String rejection) {
            return new Case(name, value, CellData::asString, null, rejection);
        }
    }

    private static List<Case> cases() {
        List<Case> cases = new ArrayList<>();
        // text
        for (String text : new String[]{" ", " padded ", "line1\nline2", "crlf\r\nx", "tab\there", "=SUM(A1)", "+1", "-1",
                "@x", "'apostrophe", "\"quoted\"", "<tag>&amp;", "\uD55C\uAE00", "emoji \uD83D\uDE00", "a".repeat(32_767),
                "0123", "1e5", "true", "2024-01-01", "#REF!", "\\", "mixed \u00e9\u00fc\u00df"}) {
            cases.add(Case.of("text <" + abbreviate(text) + ">", text, CellData::asString));
        }
        // integers
        for (int number : new int[]{0, 1, -1, Integer.MAX_VALUE, Integer.MIN_VALUE}) {
            cases.add(Case.of("int " + number, number, CellData::asInt));
            cases.add(Case.of("int " + number + " as long", number, CellData::asLong, (long) number));
        }
        for (long number : new long[]{0L, Integer.MAX_VALUE + 1L, 999_999_999_999_999L, -999_999_999_999_999L}) {
            cases.add(Case.of("long " + number, number, CellData::asLong));
        }
        // Excel keeps 15 significant digits (it truncates what is typed beyond them), so a long with more
        // significant digits is refused instead of being rounded, although a double could hold some exactly.
        // (1_000_000_000_000_000 has a single significant digit and is stored.)
        cases.add(Case.of("long 1000000000000000", 1_000_000_000_000_000L, CellData::asLong));
        for (long number : new long[]{1_234_567_890_123_456L, 9_007_199_254_740_991L, 9_007_199_254_740_993L,
                Long.MAX_VALUE, Long.MIN_VALUE}) {
            cases.add(Case.rejected("long " + number, number, "cannot be stored exactly"));
        }
        // fractions
        for (double number : new double[]{0.5, -0.5, 0.1 + 0.2, 1e-7, 1.7976931348623157E308, Double.MIN_VALUE, 123456789.123456789}) {
            cases.add(Case.of("double " + number, number, CellData::asDouble));
        }
        // The value is kept, not its scale: 0.10 comes back as 0.1 and 1E+3 as 1000.
        for (String number : new String[]{"0.10", "-0.000001", "1E+3", "123456.789"}) {
            cases.add(Case.of("BigDecimal " + number, new BigDecimal(number), CellData::asBigDecimal));
        }
        // More than 15 significant digits do not fit a double.
        for (String number : new String[]{"12345678901234567890.123456789", "123456789012345.67"}) {
            cases.add(Case.rejected("BigDecimal " + number, new BigDecimal(number), "cannot be stored exactly"));
        }
        // dates and times
        for (LocalDate date : new LocalDate[]{LocalDate.of(2024, 2, 29), LocalDate.of(1970, 1, 1), LocalDate.of(1900, 3, 1),
                LocalDate.of(1900, 1, 1), LocalDate.of(9999, 12, 31), LocalDate.of(2000, 1, 1)}) {
            cases.add(Case.of("date " + date, date, CellData::asLocalDate));
        }
        // The Excel date system starts on 1900-01-01.
        for (LocalDate date : new LocalDate[]{LocalDate.of(1899, 12, 31), LocalDate.of(1850, 6, 15)}) {
            cases.add(Case.rejected("date " + date, date, "outside the Excel date system"));
        }
        for (LocalDateTime time : new LocalDateTime[]{LocalDateTime.of(2024, 3, 5, 13, 45, 59), LocalDateTime.of(2024, 3, 5, 0, 0, 0),
                LocalDateTime.of(2024, 3, 5, 23, 59, 59), LocalDateTime.of(1999, 12, 31, 23, 59, 59)}) {
            cases.add(Case.of("datetime " + time, time, CellData::asLocalDateTime));
        }
        cases.add(Case.of("datetime with millis", LocalDateTime.of(2024, 3, 5, 13, 45, 59, 123_000_000), CellData::asLocalDateTime));
        cases.add(Case.of("time 08:30", LocalTime.of(8, 30), CellData::asLocalTime));
        cases.add(Case.of("time 23:59:59", LocalTime.of(23, 59, 59), CellData::asLocalTime));
        // what XML cannot carry, what Excel limits, and values Excel has no type for
        cases.add(Case.of("control character U+0001", "a\u0001b", CellData::asString));
        cases.add(Case.of("NUL", "a\u0000b", CellData::asString));
        cases.add(Case.of("vertical tab", "a\u000Bb", CellData::asString));
        cases.add(Case.rejected("32,768 characters", "a".repeat(32_768), "maximum length"));
        cases.add(Case.rejected("NaN", Double.NaN, "Non-finite"));
        cases.add(Case.rejected("infinity", Double.POSITIVE_INFINITY, "Non-finite"));
        cases.add(Case.of("negative zero", -0.0, CellData::asDouble));
        cases.add(Case.of("float", 1.5f, CellData::asFloat));
        cases.add(Case.of("short", (short) 7, CellData::asInt, 7));
        cases.add(Case.of("byte", (byte) 7, CellData::asInt, 7));
        cases.add(Case.of("character", 'x', CellData::asString, "x"));
        cases.add(Case.of("enum", java.time.DayOfWeek.MONDAY, CellData::asString, "MONDAY"));
        cases.add(Case.of("UUID", java.util.UUID.fromString("123e4567-e89b-12d3-a456-426614174000"), CellData::asString,
                "123e4567-e89b-12d3-a456-426614174000"));
        cases.add(Case.of("small BigInteger", java.math.BigInteger.valueOf(12345), CellData::asLong, 12345L));
        // booleans
        cases.add(Case.of("true", true, CellData::asBoolean));
        cases.add(Case.of("false", false, CellData::asBoolean));
        return cases;
    }

    @Test
    void everyValueReadsBackAsWrittenOrIsRefusedWithAReason() {
        List<String> problems = new ArrayList<>();
        for (Case test : cases()) {
            try {
                if (test.rejection() != null) {
                    try {
                        write(test);
                        problems.add(test.name() + ": written although Excel cannot hold it exactly");
                    } catch (ExcelWriteException e) {
                        if (!String.valueOf(e.getMessage()).contains(test.rejection())) {
                            problems.add(test.name() + ": refused with <" + e.getMessage() + ">, expected to mention <"
                                    + test.rejection() + ">");
                        }
                    }
                    continue;
                }
                Object actual = roundTrip(test);
                if (!same(test.expected(), actual)) {
                    problems.add(test.name() + ": wrote <" + show(test.value()) + ">, read <" + show(actual) + ">");
                }
            } catch (Exception | AssertionError e) {
                problems.add(test.name() + ": " + e.getClass().getSimpleName() + ": " + e.getMessage());
            }
        }
        assertTrue(problems.isEmpty(), problems.size() + " value(s) did not behave:\n  " + String.join("\n  ", problems));
    }

    /** Equal, but a decimal by its value: {@code 0.10} and {@code 0.1} are the same number. */
    private static boolean same(Object expected, Object actual) {
        if (expected instanceof BigDecimal left && actual instanceof BigDecimal right) return left.compareTo(right) == 0;
        return Objects.equals(expected, actual);
    }

    private static byte[] write(Case test) {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        ExcelWriter.<Object>create().column("v", v -> v).write(List.of(test.value())).writeTo(out);
        return out.toByteArray();
    }

    private static Object roundTrip(Case test) {
        byte[] file = write(test);
        List<Object> read = new ArrayList<>();
        ExcelReader.<Object>mapping(row -> test.read().apply(row.get("v")))
                .read(new ByteArrayInputStream(file), result -> {
                    assertTrue(result.success(), "row did not read: " + result.messages());
                    read.add(result.data());
                });
        assertTrue(read.size() == 1, "one row expected, got " + read.size());
        return read.get(0);
    }

    private static String abbreviate(String text) {
        String shown = text.replace("\n", "\\n").replace("\r", "\\r").replace("\t", "\\t");
        return shown.length() > 24 ? shown.substring(0, 24) + "... (" + text.length() + " chars)" : shown;
    }

    private static String show(Object value) {
        if (value == null) return "null";
        String text = value instanceof String string ? abbreviate(string) : String.valueOf(value);
        return text + " (" + value.getClass().getSimpleName() + ")";
    }
}
