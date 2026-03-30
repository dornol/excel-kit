package io.github.dornol.excelkit.csv;

import io.github.dornol.excelkit.shared.ReadResult;
import org.junit.jupiter.api.Test;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.nio.charset.StandardCharsets;
import java.util.ArrayList;
import java.util.List;
import java.util.stream.Stream;

import static org.junit.jupiter.api.Assertions.*;

/**
 * Tests for {@link CsvQuoting} strategies in {@link CsvWriter}.
 */
class CsvQuotingTest {

    record Item(String name, int value, double price) {}

    private String writeAndGet(CsvQuoting quoting, boolean injectionDefense, Item... items) {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        new CsvWriter<Item>()
                .quoting(quoting)
                .bom(false)
                .csvInjectionDefense(injectionDefense)
                .column("Name", Item::name)
                .column("Value", i -> i.value)
                .column("Price", i -> i.price)
                .write(Stream.of(items))
                .consumeOutputStream(out);
        return out.toString(StandardCharsets.UTF_8);
    }

    private String[] dataLines(String csv) {
        String[] lines = csv.split("\n");
        String[] data = new String[lines.length - 1];
        for (int i = 1; i < lines.length; i++) {
            data[i - 1] = lines[i].trim();
        }
        return data;
    }

    // ============================================================
    // MINIMAL (default)
    // ============================================================
    @Test
    void minimal_isDefault_whenNotExplicitlySet() {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        new CsvWriter<Item>()
                .bom(false)
                .csvInjectionDefense(false)
                .column("Name", Item::name)
                .column("Value", i -> i.value)
                .column("Price", i -> i.price)
                .write(Stream.of(new Item("Alice", 10, 9.99)))
                .consumeOutputStream(out);
        String[] data = dataLines(out.toString(StandardCharsets.UTF_8));

        // Default = MINIMAL: plain values not quoted
        assertEquals("Alice,10,9.99", data[0]);
    }

    @Test
    void minimal_plainValues_noQuotes() {
        String csv = writeAndGet(CsvQuoting.MINIMAL, false, new Item("Alice", 10, 9.99));
        String[] data = dataLines(csv);

        assertEquals("Alice,10,9.99", data[0]);
    }

    @Test
    void minimal_valueWithDelimiter_quoted() {
        String csv = writeAndGet(CsvQuoting.MINIMAL, false, new Item("A,B", 10, 9.99));
        String[] data = dataLines(csv);

        assertTrue(data[0].startsWith("\"A,B\""), "Value with comma should be quoted");
    }

    @Test
    void minimal_valueWithQuote_escaped() {
        String csv = writeAndGet(CsvQuoting.MINIMAL, false, new Item("say\"hi", 10, 9.99));
        String[] data = dataLines(csv);

        // say"hi → "say""hi" (wrapped in quotes, internal " doubled)
        assertTrue(data[0].startsWith("\"say\"\"hi\""), "Quotes should be doubled and wrapped");
    }

    @Test
    void minimal_valueWithNewline_quoted() {
        String csv = writeAndGet(CsvQuoting.MINIMAL, false, new Item("line1\nline2", 10, 1.0));
        // The value with newline should be enclosed in quotes
        assertTrue(csv.contains("\"line1\nline2\""), "Newline should trigger quoting");
    }

    // ============================================================
    // ALL
    // ============================================================
    @Test
    void all_headerAndDataAllQuoted() {
        String csv = writeAndGet(CsvQuoting.ALL, false, new Item("Alice", 10, 9.99));
        String[] lines = csv.split("\n");

        // Header: every field quoted
        assertEquals("\"Name\",\"Value\",\"Price\"", lines[0].trim());
        // Data: every field quoted
        assertEquals("\"Alice\",\"10\",\"9.99\"", lines[1].trim());
    }

    @Test
    void all_nullValue_quotedEmpty() {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        new CsvWriter<String>()
                .quoting(CsvQuoting.ALL)
                .bom(false)
                .column("A", s -> null)
                .column("B", s -> "x")
                .write(Stream.of("test"))
                .consumeOutputStream(out);
        String[] data = dataLines(out.toString(StandardCharsets.UTF_8));

        assertEquals("\"\",\"x\"", data[0]);
    }

    @Test
    void all_valueWithQuote_properlyEscaped() {
        String csv = writeAndGet(CsvQuoting.ALL, false, new Item("A\"B", 10, 1.0));
        String[] data = dataLines(csv);

        assertTrue(data[0].startsWith("\"A\"\"B\""), "Internal quotes should be doubled in ALL mode");
    }

    // ============================================================
    // NON_NUMERIC
    // ============================================================
    @Test
    void nonNumeric_stringsQuoted_numbersNot() {
        String csv = writeAndGet(CsvQuoting.NON_NUMERIC, false, new Item("Alice", 10, 9.99));
        String[] lines = csv.split("\n");

        // Header: non-numeric → all quoted
        assertEquals("\"Name\",\"Value\",\"Price\"", lines[0].trim());
        // Data: "Alice" quoted, 10 and 9.99 unquoted
        assertEquals("\"Alice\",10,9.99", lines[1].trim());
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

        String[] data = dataLines(out.toString(StandardCharsets.UTF_8));
        assertEquals("-42.5", data[0], "Negative number should not be quoted");
    }

    @Test
    void nonNumeric_positiveSignNumber_notQuoted() {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        new CsvWriter<String>()
                .quoting(CsvQuoting.NON_NUMERIC)
                .bom(false)
                .csvInjectionDefense(false)
                .column("Val", s -> s)
                .write(Stream.of("+3.14"))
                .consumeOutputStream(out);

        String[] data = dataLines(out.toString(StandardCharsets.UTF_8));
        assertEquals("+3.14", data[0], "Positive sign number should not be quoted");
    }

    @Test
    void nonNumeric_textThatLooksNumeric_quoted() {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        new CsvWriter<String>()
                .quoting(CsvQuoting.NON_NUMERIC)
                .bom(false)
                .column("Val", s -> s)
                .write(Stream.of("10abc"))
                .consumeOutputStream(out);

        String[] data = dataLines(out.toString(StandardCharsets.UTF_8));
        assertEquals("\"10abc\"", data[0], "Non-numeric text should be quoted");
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

        String[] data = dataLines(out.toString(StandardCharsets.UTF_8));
        assertEquals("\"\"", data[0], "Empty string is non-numeric → should be quoted");
    }

    // ============================================================
    // Interaction with csvInjectionDefense
    // ============================================================
    @Test
    void nonNumeric_withInjectionDefense_negativeNumber_prefixedAndQuoted() {
        // '-' is a formula character, so with defense ON, it gets prefixed with '
        // After prefixing, the value is no longer numeric → NON_NUMERIC will quote it
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        new CsvWriter<String>()
                .quoting(CsvQuoting.NON_NUMERIC)
                .bom(false)
                .csvInjectionDefense(true)
                .column("Val", s -> s)
                .write(Stream.of("-42"))
                .consumeOutputStream(out);

        String[] data = dataLines(out.toString(StandardCharsets.UTF_8));
        // '-42' → injection defense adds ' prefix → "'-42" which is non-numeric → quoted
        assertEquals("\"'-42\"", data[0]);
    }

    // ============================================================
    // Round-trip: write with quoting, read back
    // ============================================================
    @Test
    void roundTrip_all_shouldReadBackCorrectly() {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        new CsvWriter<Item>()
                .quoting(CsvQuoting.ALL)
                .bom(false)
                .csvInjectionDefense(false)
                .column("Name", Item::name)
                .column("Value", i -> String.valueOf(i.value))
                .write(Stream.of(new Item("Alice", 10, 9.99), new Item("Bob", 20, 1.5)))
                .consumeOutputStream(out);

        List<ReadResult<Item>> results = new ArrayList<>();
        CsvReader.<Item>mapping(row ->
                new Item(row.get("Name").asString(), row.get("Value").asInt(), 0)
        ).build(new ByteArrayInputStream(out.toByteArray()))
                .read(results::add);

        assertEquals(2, results.size());
        assertTrue(results.get(0).success());
        assertEquals("Alice", results.get(0).data().name());
        assertEquals(10, results.get(0).data().value());
        assertEquals("Bob", results.get(1).data().name());
    }

    @Test
    void roundTrip_nonNumeric_shouldReadBackCorrectly() {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        new CsvWriter<Item>()
                .quoting(CsvQuoting.NON_NUMERIC)
                .bom(false)
                .csvInjectionDefense(false)
                .column("Name", Item::name)
                .column("Value", i -> String.valueOf(i.value))
                .write(Stream.of(new Item("Alice", 10, 9.99)))
                .consumeOutputStream(out);

        List<ReadResult<Item>> results = new ArrayList<>();
        CsvReader.<Item>mapping(row ->
                new Item(row.get("Name").asString(), row.get("Value").asInt(), 0)
        ).build(new ByteArrayInputStream(out.toByteArray()))
                .read(results::add);

        assertEquals(1, results.size());
        assertEquals("Alice", results.get(0).data().name());
        assertEquals(10, results.get(0).data().value());
    }

    // ============================================================
    // Enum & fluent
    // ============================================================
    @Test
    void allQuotingValues() {
        assertEquals(3, CsvQuoting.values().length);
        for (var q : CsvQuoting.values()) {
            assertEquals(q, CsvQuoting.valueOf(q.name()));
        }
    }

    @Test
    void quoting_returnsSameInstance() {
        CsvWriter<String> w = new CsvWriter<>();
        assertSame(w, w.quoting(CsvQuoting.ALL));
    }
}
