package io.github.dornol.excelkit.csv;

import io.github.dornol.excelkit.shared.ReadResult;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.params.ParameterizedTest;
import org.junit.jupiter.params.provider.EnumSource;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.nio.charset.StandardCharsets;
import java.util.ArrayList;
import java.util.List;
import java.util.stream.Stream;

import static org.junit.jupiter.api.Assertions.*;

/**
 * Tests for {@link CsvDialect} presets and their integration with CsvWriter/CsvReader.
 */
class CsvDialectTest {

    record Item(String name, int value) {}

    // ============================================================
    // Enum coverage
    // ============================================================
    @Test
    void allDialectsExist() {
        assertEquals(4, CsvDialect.values().length);
    }

    @ParameterizedTest
    @EnumSource(CsvDialect.class)
    void dialect_hasValidProperties(CsvDialect dialect) {
        assertNotNull(dialect.getCharset());
        assertTrue(dialect.getDelimiter() != 0);
    }

    @Test
    void rfc4180_properties() {
        assertEquals(',', CsvDialect.RFC4180.getDelimiter());
        assertEquals(StandardCharsets.UTF_8, CsvDialect.RFC4180.getCharset());
        assertFalse(CsvDialect.RFC4180.isBom());
    }

    @Test
    void excel_properties() {
        assertEquals(',', CsvDialect.EXCEL.getDelimiter());
        assertTrue(CsvDialect.EXCEL.isBom());
    }

    @Test
    void tsv_properties() {
        assertEquals('\t', CsvDialect.TSV.getDelimiter());
        assertFalse(CsvDialect.TSV.isBom());
    }

    @Test
    void pipe_properties() {
        assertEquals('|', CsvDialect.PIPE.getDelimiter());
        assertFalse(CsvDialect.PIPE.isBom());
    }

    // ============================================================
    // CsvWriter integration
    // ============================================================
    @Test
    void csvWriter_tsvDialect_shouldUseTabDelimiter() {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        new CsvWriter<Item>()
                .dialect(CsvDialect.TSV)
                .column("Name", Item::name)
                .column("Value", i -> String.valueOf(i.value))
                .write(Stream.of(new Item("Alice", 10)))
                .consumeOutputStream(out);

        String content = out.toString(StandardCharsets.UTF_8);
        assertTrue(content.contains("Name\tValue"), "Should use tab delimiter");
        assertTrue(content.contains("Alice\t10"));
    }

    @Test
    void csvWriter_rfc4180_noBom() {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        new CsvWriter<Item>()
                .dialect(CsvDialect.RFC4180)
                .column("Name", Item::name)
                .write(Stream.of(new Item("A", 1)))
                .consumeOutputStream(out);

        byte[] bytes = out.toByteArray();
        // Should NOT start with BOM
        assertFalse(bytes.length >= 3
                && bytes[0] == (byte) 0xEF
                && bytes[1] == (byte) 0xBB
                && bytes[2] == (byte) 0xBF);
    }

    @Test
    void csvWriter_excelDialect_withBom() {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        new CsvWriter<Item>()
                .dialect(CsvDialect.EXCEL)
                .column("Name", Item::name)
                .write(Stream.of(new Item("A", 1)))
                .consumeOutputStream(out);

        byte[] bytes = out.toByteArray();
        // Should start with BOM
        assertTrue(bytes.length >= 3
                && bytes[0] == (byte) 0xEF
                && bytes[1] == (byte) 0xBB
                && bytes[2] == (byte) 0xBF);
    }

    @Test
    void csvWriter_pipeDialect() {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        new CsvWriter<Item>()
                .dialect(CsvDialect.PIPE)
                .column("Name", Item::name)
                .column("Value", i -> String.valueOf(i.value))
                .write(Stream.of(new Item("Alice", 10)))
                .consumeOutputStream(out);

        String content = out.toString(StandardCharsets.UTF_8);
        assertTrue(content.contains("Name|Value"));
    }

    // ============================================================
    // CsvReader integration
    // ============================================================
    @Test
    void csvReader_tsvDialect_shouldReadTabSeparated() {
        String tsv = "Name\tValue\nAlice\t10\nBob\t20";

        List<Item> results = new ArrayList<>();
        CsvReader.<Item>mapping(row ->
                new Item(row.get("Name").asString(), row.get("Value").asInt())
        ).dialect(CsvDialect.TSV)
                .build(new ByteArrayInputStream(tsv.getBytes(StandardCharsets.UTF_8)))
                .read(r -> {
                    assertTrue(r.success());
                    results.add(r.data());
                });

        assertEquals(2, results.size());
        assertEquals("Alice", results.get(0).name());
        assertEquals(20, results.get(1).value());
    }

    @Test
    void csvReader_pipeDialect_shouldReadPipeSeparated() {
        String csv = "Name|Value\nAlice|10";

        List<Item> results = new ArrayList<>();
        CsvReader.<Item>mapping(row ->
                new Item(row.get("Name").asString(), row.get("Value").asInt())
        ).dialect(CsvDialect.PIPE)
                .build(new ByteArrayInputStream(csv.getBytes(StandardCharsets.UTF_8)))
                .read(r -> results.add(r.data()));

        assertEquals(1, results.size());
        assertEquals("Alice", results.get(0).name());
    }

    // ============================================================
    // Round-trip: write with dialect, read back
    // ============================================================
    @Test
    void roundTrip_tsvDialect() {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        new CsvWriter<Item>()
                .dialect(CsvDialect.TSV)
                .column("Name", Item::name)
                .column("Value", i -> String.valueOf(i.value))
                .write(Stream.of(new Item("Alice", 10), new Item("Bob", 20)))
                .consumeOutputStream(out);

        List<Item> results = new ArrayList<>();
        CsvReader.<Item>mapping(row ->
                new Item(row.get("Name").asString(), row.get("Value").asInt())
        ).dialect(CsvDialect.TSV)
                .build(new ByteArrayInputStream(out.toByteArray()))
                .read(r -> results.add(r.data()));

        assertEquals(2, results.size());
        assertEquals("Alice", results.get(0).name());
        assertEquals(20, results.get(1).value());
    }

    // ============================================================
    // Dialect can be overridden
    // ============================================================
    @Test
    void dialect_canBeOverriddenByIndividualSettings() {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        new CsvWriter<Item>()
                .dialect(CsvDialect.TSV)       // tab
                .delimiter(';')                 // override to semicolon
                .column("Name", Item::name)
                .write(Stream.of(new Item("A", 1)))
                .consumeOutputStream(out);

        String content = out.toString(StandardCharsets.UTF_8);
        assertTrue(content.contains("Name"));
        assertFalse(content.contains("\t"), "Tab should be overridden by semicolon");
    }
}
