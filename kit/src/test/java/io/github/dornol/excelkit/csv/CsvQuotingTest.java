package io.github.dornol.excelkit.csv;

import org.junit.jupiter.api.Test;

import java.io.ByteArrayOutputStream;
import java.nio.charset.StandardCharsets;
import java.util.stream.Stream;

import static org.junit.jupiter.api.Assertions.*;

/**
 * Tests for {@link CsvQuoting} strategies in {@link CsvWriter}.
 */
class CsvQuotingTest {

    record Item(String name, int value, double price) {}

    private String writeAndGet(CsvQuoting quoting, Item... items) {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        new CsvWriter<Item>()
                .quoting(quoting)
                .bom(false)
                .csvInjectionDefense(false)
                .column("Name", Item::name)
                .column("Value", i -> i.value)
                .column("Price", i -> i.price)
                .write(Stream.of(items))
                .consumeOutputStream(out);
        return out.toString(StandardCharsets.UTF_8);
    }

    // ============================================================
    // MINIMAL (default)
    // ============================================================
    @Test
    void minimal_plainValues_noQuotes() {
        String csv = writeAndGet(CsvQuoting.MINIMAL, new Item("Alice", 10, 9.99));

        assertTrue(csv.contains("Alice"));
        assertFalse(csv.contains("\"Alice\""), "Plain values should not be quoted");
    }

    @Test
    void minimal_valueWithDelimiter_quoted() {
        String csv = writeAndGet(CsvQuoting.MINIMAL, new Item("A,B", 10, 9.99));

        assertTrue(csv.contains("\"A,B\""), "Values with delimiter should be quoted");
    }

    @Test
    void minimal_valueWithQuote_escaped() {
        String csv = writeAndGet(CsvQuoting.MINIMAL, new Item("He said \"hi\"", 10, 9.99));

        assertTrue(csv.contains("\"He said \"\"hi\"\"\""), "Quotes should be escaped");
    }

    // ============================================================
    // ALL
    // ============================================================
    @Test
    void all_allFieldsQuoted() {
        String csv = writeAndGet(CsvQuoting.ALL, new Item("Alice", 10, 9.99));

        // Header
        assertTrue(csv.contains("\"Name\""), "Header should be quoted");
        assertTrue(csv.contains("\"Value\""), "Header should be quoted");

        // Data - all fields quoted
        assertTrue(csv.contains("\"Alice\""), "String should be quoted");
        assertTrue(csv.contains("\"10\""), "Integer should be quoted");
        assertTrue(csv.contains("\"9.99\""), "Double should be quoted");
    }

    @Test
    void all_nullValue_emptyQuoted() {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        new CsvWriter<String>()
                .quoting(CsvQuoting.ALL)
                .bom(false)
                .column("Val", s -> null)
                .write(Stream.of("test"))
                .consumeOutputStream(out);
        String csv = out.toString(StandardCharsets.UTF_8);

        // null → ""  (quoted empty)
        String dataLine = csv.split("\n")[1].trim();
        assertEquals("\"\"", dataLine);
    }

    @Test
    void all_valueWithQuote_properlyEscaped() {
        String csv = writeAndGet(CsvQuoting.ALL, new Item("A\"B", 10, 1.0));

        assertTrue(csv.contains("\"A\"\"B\""), "Quotes inside ALL mode should be escaped");
    }

    // ============================================================
    // NON_NUMERIC
    // ============================================================
    @Test
    void nonNumeric_stringsQuoted_numbersNot() {
        String csv = writeAndGet(CsvQuoting.NON_NUMERIC, new Item("Alice", 10, 9.99));

        // String fields quoted
        assertTrue(csv.contains("\"Alice\""), "Non-numeric should be quoted");
        assertTrue(csv.contains("\"Name\""), "Header should be quoted");

        // Numeric fields NOT quoted
        String[] lines = csv.split("\n");
        String dataLine = lines[1].trim();
        // Should contain unquoted 10 and 9.99
        assertTrue(dataLine.contains(",10,") || dataLine.endsWith(",10"),
                "Integer should not be quoted in NON_NUMERIC mode");
    }

    @Test
    void nonNumeric_negativeNumber_notQuoted() {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        new CsvWriter<String>()
                .quoting(CsvQuoting.NON_NUMERIC)
                .bom(false)
                .csvInjectionDefense(false)
                .column("Val", s -> s)
                .write(Stream.of("-42.5"))
                .consumeOutputStream(out);
        String csv = out.toString(StandardCharsets.UTF_8);

        String dataLine = csv.split("\n")[1].trim();
        assertEquals("-42.5", dataLine, "Negative number should not be quoted");
    }

    @Test
    void nonNumeric_emptyString_quoted() {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        new CsvWriter<String>()
                .quoting(CsvQuoting.NON_NUMERIC)
                .bom(false)
                .column("Val", s -> "")
                .write(Stream.of("x"))
                .consumeOutputStream(out);
        String csv = out.toString(StandardCharsets.UTF_8);

        String dataLine = csv.split("\n")[1].trim();
        assertEquals("\"\"", dataLine, "Empty string should be quoted");
    }

    // ============================================================
    // Enum coverage
    // ============================================================
    @Test
    void allQuotingValues() {
        assertEquals(3, CsvQuoting.values().length);
        for (var q : CsvQuoting.values()) {
            assertEquals(q, CsvQuoting.valueOf(q.name()));
        }
    }

    // ============================================================
    // Fluent chaining
    // ============================================================
    @Test
    void quoting_returnsSameInstance() {
        CsvWriter<String> w = new CsvWriter<>();
        assertSame(w, w.quoting(CsvQuoting.ALL));
    }
}
