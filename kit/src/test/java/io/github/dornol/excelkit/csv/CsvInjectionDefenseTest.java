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

    private String dataLine(String csv) {
        String[] lines = csv.split("\r?\n");
        return lines.length > 1 ? lines[1].trim() : "";
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
            String dataLine = dataLine(csv);
            assertEquals("'=SUM(A1)", dataLine, "Should prefix = with quote");
        }

        @Test
        void prefixes_plusSign() throws Exception {
            String csv = writeCsv(defaultWriter(), "+cmd");
            assertEquals("'+cmd", dataLine(csv), "Should prefix + with quote");
        }

        @Test
        void prefixes_minusSign() throws Exception {
            String csv = writeCsv(defaultWriter(), "-calc");
            assertEquals("'-calc", dataLine(csv), "Should prefix - with quote");
        }

        @Test
        void prefixes_atSign() throws Exception {
            String csv = writeCsv(defaultWriter(), "@evil");
            assertEquals("'@evil", dataLine(csv), "Should prefix @ with quote");
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
            String[] lines = csv.split("\r?\n");
            assertEquals("hello", lines[1].trim());
            assertEquals("world", lines[2].trim());
            assertEquals("123", lines[3].trim());
        }

        @Test
        void handlesNullValues_gracefully() throws Exception {
            var writer = new CsvWriter<String>()
                    .bom(false)
                    .column("Value", s -> null);
            String csv = writeCsv(writer, "anything");
            String dataLine = dataLine(csv);
            assertEquals("", dataLine, "null should produce empty field");
        }

        @Test
        void handlesEmptyString() throws Exception {
            String csv = writeCsv(defaultWriter(), "");
            String dataLine = dataLine(csv);
            assertEquals("", dataLine, "Empty string should produce empty field without prefix");
        }

        @Test
        void formulaCharNotAtStart_noPrefix() throws Exception {
            String csv = writeCsv(defaultWriter(), "a=b", "x+y", "hello@world", "a-b");
            String[] lines = csv.split("\r?\n");
            assertEquals("a=b", lines[1].trim());
            assertEquals("x+y", lines[2].trim());
            assertEquals("hello@world", lines[3].trim());
            assertEquals("a-b", lines[4].trim());
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
            assertEquals("=SUM(A1)", dataLine(csv));
        }

        @Test
        void noPrefix_forPlusSign() throws Exception {
            String csv = writeCsv(defenseDisabledWriter(), "+cmd");
            assertEquals("+cmd", dataLine(csv));
        }

        @Test
        void noPrefix_forMinusSign() throws Exception {
            String csv = writeCsv(defenseDisabledWriter(), "-calc");
            assertEquals("-calc", dataLine(csv));
        }

        @Test
        void noPrefix_forAtSign() throws Exception {
            String csv = writeCsv(defenseDisabledWriter(), "@evil");
            assertEquals("@evil", dataLine(csv));
        }

        @Test
        void normalValues_unchanged() throws Exception {
            String csv = writeCsv(defenseDisabledWriter(), "hello", "123");
            String[] lines = csv.split("\r?\n");
            assertEquals("hello", lines[1].trim());
            assertEquals("123", lines[2].trim());
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
