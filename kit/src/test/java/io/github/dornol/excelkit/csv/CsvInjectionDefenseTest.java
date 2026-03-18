package io.github.dornol.excelkit.csv;

import org.junit.jupiter.api.Nested;
import org.junit.jupiter.api.Test;

import java.io.ByteArrayOutputStream;
import java.nio.charset.StandardCharsets;
import java.util.stream.Stream;

import static org.junit.jupiter.api.Assertions.*;

/**
 * Tests for CsvWriter CSV injection defense toggle.
 */
class CsvInjectionDefenseTest {

    private String writeCsv(CsvWriter<String> writer, String... values) throws Exception {
        var handler = writer.write(Stream.of(values));
        var baos = new ByteArrayOutputStream();
        handler.consumeOutputStream(baos);
        return baos.toString(StandardCharsets.UTF_8);
    }

    private CsvWriter<String> defaultWriter() {
        return new CsvWriter<String>()
                .bom(false)
                .column("Value", s -> s);
    }

    private CsvWriter<String> defenseDisabledWriter() {
        return new CsvWriter<String>()
                .bom(false)
                .csvInjectionDefense(false)
                .column("Value", s -> s);
    }

    // ============================================================
    // Defense enabled (default behavior)
    // ============================================================
    @Nested
    class DefenseEnabled {

        @Test
        void prefixes_equalsSign() throws Exception {
            String csv = writeCsv(defaultWriter(), "=SUM(A1)");
            assertTrue(csv.contains("'=SUM(A1)"), "Should prefix = with quote");
        }

        @Test
        void prefixes_plusSign() throws Exception {
            String csv = writeCsv(defaultWriter(), "+cmd");
            assertTrue(csv.contains("'+cmd"), "Should prefix + with quote");
        }

        @Test
        void prefixes_minusSign() throws Exception {
            String csv = writeCsv(defaultWriter(), "-calc");
            assertTrue(csv.contains("'-calc"), "Should prefix - with quote");
        }

        @Test
        void prefixes_atSign() throws Exception {
            String csv = writeCsv(defaultWriter(), "@evil");
            assertTrue(csv.contains("'@evil"), "Should prefix @ with quote");
        }

        @Test
        void prefixes_tabCharacter() throws Exception {
            String csv = writeCsv(defaultWriter(), "\tcmd");
            assertTrue(csv.contains("'\tcmd"), "Should prefix tab with quote");
        }

        @Test
        void prefixes_carriageReturn() throws Exception {
            String csv = writeCsv(defaultWriter(), "\rcmd");
            assertTrue(csv.contains("'\rcmd"), "Should prefix carriage return with quote");
        }

        @Test
        void doesNotPrefix_normalValues() throws Exception {
            String csv = writeCsv(defaultWriter(), "hello", "world", "123");
            assertTrue(csv.contains("hello"));
            assertTrue(csv.contains("world"));
            assertTrue(csv.contains("123"));
            assertFalse(csv.contains("'hello"), "Normal values should not be prefixed");
            assertFalse(csv.contains("'world"), "Normal values should not be prefixed");
            assertFalse(csv.contains("'123"), "Normal values should not be prefixed");
        }

        @Test
        void handlesNullValues_gracefully() throws Exception {
            // null extractor result should produce empty field, not NPE
            var writer = new CsvWriter<String>()
                    .bom(false)
                    .column("Value", s -> null);
            String csv = writeCsv(writer, "anything");
            // Should not throw; null becomes empty string
            assertNotNull(csv);
        }

        @Test
        void handlesEmptyString() throws Exception {
            String csv = writeCsv(defaultWriter(), "");
            // Empty string should not be prefixed
            assertNotNull(csv);
            assertFalse(csv.contains("'"), "Empty string should not get a quote prefix");
        }

        @Test
        void formulaCharNotAtStart_noPrefix() throws Exception {
            String csv = writeCsv(defaultWriter(), "a=b", "x+y", "hello@world", "a-b");
            assertTrue(csv.contains("a=b"));
            assertTrue(csv.contains("x+y"));
            assertTrue(csv.contains("hello@world"));
            assertTrue(csv.contains("a-b"));
            assertFalse(csv.contains("'a=b"), "= not at start should not be prefixed");
            assertFalse(csv.contains("'x+y"), "+ not at start should not be prefixed");
            assertFalse(csv.contains("'hello@world"), "@ not at start should not be prefixed");
            assertFalse(csv.contains("'a-b"), "- not at start should not be prefixed");
        }

        @Test
        void multiColumn_defenseAppliesToAllColumns() throws Exception {
            var writer = new CsvWriter<String[]>()
                    .bom(false)
                    .column("Col1", arr -> arr[0])
                    .column("Col2", arr -> arr[1]);
            var handler = writer.write(Stream.<String[]>of(new String[]{"=formula", "@inject"}));
            var baos = new ByteArrayOutputStream();
            handler.consumeOutputStream(baos);
            String csv = baos.toString(StandardCharsets.UTF_8);

            assertTrue(csv.contains("'=formula"), "Defense should apply to Col1");
            assertTrue(csv.contains("'@inject"), "Defense should apply to Col2");
        }
    }

    // ============================================================
    // Defense disabled
    // ============================================================
    @Nested
    class DefenseDisabled {

        @Test
        void noPrefix_forEqualsSign() throws Exception {
            String csv = writeCsv(defenseDisabledWriter(), "=SUM(A1)");
            assertTrue(csv.contains("=SUM(A1)"));
            assertFalse(csv.contains("'=SUM(A1)"));
        }

        @Test
        void noPrefix_forPlusSign() throws Exception {
            String csv = writeCsv(defenseDisabledWriter(), "+cmd");
            assertTrue(csv.contains("+cmd"));
            assertFalse(csv.contains("'+cmd"));
        }

        @Test
        void noPrefix_forMinusSign() throws Exception {
            String csv = writeCsv(defenseDisabledWriter(), "-calc");
            assertTrue(csv.contains("-calc"));
            assertFalse(csv.contains("'-calc"));
        }

        @Test
        void noPrefix_forAtSign() throws Exception {
            String csv = writeCsv(defenseDisabledWriter(), "@evil");
            assertTrue(csv.contains("@evil"));
            assertFalse(csv.contains("'@evil"));
        }

        @Test
        void normalValues_unchanged() throws Exception {
            String csv = writeCsv(defenseDisabledWriter(), "hello", "123");
            assertTrue(csv.contains("hello"));
            assertTrue(csv.contains("123"));
        }

        @Test
        void canBeToggledBackOn() throws Exception {
            var writer = new CsvWriter<String>()
                    .bom(false)
                    .csvInjectionDefense(false)
                    .csvInjectionDefense(true)  // re-enable
                    .column("Value", s -> s);

            String csv = writeCsv(writer, "=formula");
            assertTrue(csv.contains("'=formula"), "Should prefix after re-enabling");
        }
    }

    // ============================================================
    // Edge cases
    // ============================================================
    @Nested
    class EdgeCases {

        @Test
        void tabCharacterAtStart_isPrefixed() throws Exception {
            String csv = writeCsv(defaultWriter(), "\tdata");
            assertTrue(csv.contains("'\tdata"), "Tab at start should be prefixed");
        }

        @Test
        void carriageReturnAtStart_isPrefixed() throws Exception {
            String csv = writeCsv(defaultWriter(), "\rdata");
            assertTrue(csv.contains("'\rdata"), "CR at start should be prefixed");
        }

        @Test
        void valueIsJustEquals_isPrefixed() throws Exception {
            String csv = writeCsv(defaultWriter(), "=");
            assertTrue(csv.contains("'="), "Single = should be prefixed");
        }

        @Test
        void valueIsJustMinus_isPrefixed() throws Exception {
            String csv = writeCsv(defaultWriter(), "-");
            assertTrue(csv.contains("'-"), "Single - should be prefixed");
        }

        @Test
        void valueIsJustPlus_isPrefixed() throws Exception {
            String csv = writeCsv(defaultWriter(), "+");
            assertTrue(csv.contains("'+"), "Single + should be prefixed");
        }

        @Test
        void valueIsJustAt_isPrefixed() throws Exception {
            String csv = writeCsv(defaultWriter(), "@");
            assertTrue(csv.contains("'@"), "Single @ should be prefixed");
        }

        @Test
        void longValueWithFormulaCharsInMiddle_noPrefix() throws Exception {
            String longValue = "This is a long value with = and + and - and @ in the middle";
            String csv = writeCsv(defaultWriter(), longValue);
            assertTrue(csv.contains(longValue), "Long value with formula chars in middle should appear unmodified");
            assertFalse(csv.contains("'" + longValue), "Should not be prefixed when formula char is not at start");
        }

        @Test
        void valueStartingWithDigit_noPrefix() throws Exception {
            String csv = writeCsv(defaultWriter(), "0", "1", "99");
            assertFalse(csv.contains("'0"));
            assertFalse(csv.contains("'1"));
            assertFalse(csv.contains("'99"));
        }

        @Test
        void valueStartingWithQuote_noPrefix() throws Exception {
            // Single quote itself is not a formula character
            String csv = writeCsv(defaultWriter(), "'already quoted");
            // Should not double-prefix
            assertFalse(csv.contains("''already quoted"), "Should not double-prefix single quote");
        }
    }
}
