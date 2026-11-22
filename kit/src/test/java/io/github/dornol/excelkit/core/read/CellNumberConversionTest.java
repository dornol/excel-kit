package io.github.dornol.excelkit.core.read;

import org.junit.jupiter.api.Test;

import java.math.BigDecimal;
import java.util.Locale;

import static org.junit.jupiter.api.Assertions.*;

/**
 * Number conversion has a shortcut for plain digits and reuses the locale's number format; neither may change
 * what a value means in a locale.
 */
class CellNumberConversionTest {

    private static CellData cell(String text, Locale locale) {
        return new CellData(0, text, CellConversionConfig.builder().locale(locale).build());
    }

    private static final Locale[] LOCALES = {Locale.ROOT, Locale.US, Locale.GERMANY, Locale.FRANCE, Locale.KOREA,
            Locale.forLanguageTag("ar-EG"), Locale.forLanguageTag("hi-IN"), Locale.forLanguageTag("de-CH")};

    @Test
    void plainDigitsMeanTheSameInEveryLocale() {
        for (Locale locale : LOCALES) {
            assertEquals(12345L, cell("12345", locale).asLong(), locale.toString());
            assertEquals(12345, cell("12345", locale).asInt(), locale.toString());
            assertEquals(12345.0, cell("12345", locale).asDouble(), locale.toString());
            assertEquals(7L, cell("007", locale).asLong(), locale.toString());
            assertEquals(0L, cell("0", locale).asLong(), locale.toString());
            assertEquals(new BigDecimal("12345"), cell("12345", locale).asBigDecimal(), locale.toString());
            assertEquals(Long.valueOf(12345), cell("12345", locale).asNumber(), locale.toString());
        }
    }

    @Test
    void theLongestPlainNumbersAndTheOnesBeyondTheShortcutStillConvert() {
        for (Locale locale : LOCALES) {
            assertEquals(999_999_999_999_999_999L, cell("999999999999999999", locale).asLong());
            assertEquals(Long.MAX_VALUE, cell("9223372036854775807", locale).asLong());
            assertThrows(IllegalArgumentException.class, () -> cell("9223372036854775808", locale).asLong());
            assertEquals(new BigDecimal("123456789012345678901234567890"),
                    cell("123456789012345678901234567890", locale).asBigDecimal());
        }
    }

    @Test
    void separatorsStillFollowTheLocale() {
        assertEquals(1234L, cell("1,234", Locale.US).asLong());
        assertEquals(1234L, cell("1.234", Locale.GERMANY).asLong());
        assertEquals(1.234, cell("1,234", Locale.GERMANY).asDouble(), 0.0);
        assertEquals(123.45, cell("123.45", Locale.US).asDouble(), 0.0);
        assertEquals(123.45, cell("123,45", Locale.FRANCE).asDouble(), 0.0);
        assertThrows(IllegalArgumentException.class, () -> cell("1.234", Locale.US).asLong(), "a fraction is not a long");
    }

    @Test
    void signsAffixesAndGarbageAreStillHandledByTheGeneralPath() {
        assertEquals(-98765L, cell("-98765", Locale.US).asLong());
        assertEquals(12L, cell("$12", Locale.US).asLong());
        assertEquals(1000L, cell("₩ 1,000", Locale.KOREA).asLong());
        assertEquals(12L, cell("12원", Locale.KOREA).asLong());
        assertEquals(1000.0, cell("1e3", Locale.US).asDouble(), 0.0);
        assertEquals(12345L, cell("  12345  ", Locale.US).asLong());
        assertNull(cell("  ", Locale.US).asLong());
        for (String bad : new String[]{"abc", "1a", "a1", "--1", "1_000", "0x10"}) {
            assertThrows(IllegalArgumentException.class, () -> cell(bad, Locale.US).asLong(), bad);
        }
    }

    @Test
    void conversionsAreSafeToRunOnSeveralThreadsAtOnce() throws Exception {
        var failures = new java.util.concurrent.ConcurrentLinkedQueue<String>();
        var threads = new Thread[8];
        for (int t = 0; t < threads.length; t++) {
            threads[t] = new Thread(() -> {
                for (int i = 0; i < 20_000; i++) {
                    if (cell("1,234,567", Locale.US).asLong() != 1_234_567L) failures.add("grouped");
                    if (cell("123.45", Locale.US).asDouble() != 123.45) failures.add("decimal");
                    if (cell("1.234.567", Locale.GERMANY).asLong() != 1_234_567L) failures.add("german");
                }
            });
            threads[t].start();
        }
        for (Thread thread : threads) thread.join();
        assertTrue(failures.isEmpty(), failures.toString());
    }
}
