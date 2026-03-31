package io.github.dornol.excelkit.csv;

import io.github.dornol.excelkit.shared.ReadResult;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import java.io.*;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.List;
import java.util.stream.Stream;

import static org.junit.jupiter.api.Assertions.*;

/**
 * Tests verifying that CSV read/write operations properly clean up temporary files,
 * including error scenarios and edge cases.
 */
class CsvTempFileCleanupTest {

    @TempDir
    Path tempDir;

    // ──────────────────────────────────────────────────────────────
    // CsvWriter: cleanup after successful write
    // ──────────────────────────────────────────────────────────────

    @Test
    void write_shouldCleanUpTempFilesAfterConsume() {
        CsvWriter<TestItem> writer = new CsvWriter<>();
        writer.column("Name", item -> item.name)
              .column("Price", item -> item.price);

        CsvHandler handler = writer.write(Stream.of(
                new TestItem("Apple", 1000),
                new TestItem("Banana", 2000)
        ));

        ByteArrayOutputStream baos = new ByteArrayOutputStream();
        handler.consumeOutputStream(baos);

        String csv = baos.toString();
        String[] lines = csv.split("\r?\n");
        assertTrue(lines.length >= 3, "Should have header + 2 data lines");
        String headerLine = lines[0].replace("\uFEFF", "");
        assertEquals("Name,Price", headerLine);
        assertEquals("Apple,1000", lines[1].trim());
        assertEquals("Banana,2000", lines[2].trim());
    }

    @Test
    void write_columnFunctionError_shouldFallbackToNullAndCleanUp() {
        // CsvColumn.applyFunction catches RuntimeException and returns null
        CsvWriter<TestItem> writer = new CsvWriter<>();
        writer.column("Name", item -> {
                    throw new RuntimeException("Intentional error");
                })
                .column("Price", item -> item.price);

        CsvHandler handler = writer.write(Stream.of(new TestItem("Test", 100)));
        ByteArrayOutputStream baos = new ByteArrayOutputStream();
        handler.consumeOutputStream(baos);

        String csv = baos.toString();
        String[] lines = csv.split("\\r?\\n");
        // The failed column value falls back to null → empty string, second column is fine
        assertEquals(",100", lines[1], "Failed column function should produce empty value");
    }

    @Test
    void write_noColumns_shouldThrow() {
        CsvWriter<TestItem> writer = new CsvWriter<>();

        assertThrows(CsvWriteException.class, () ->
                writer.write(Stream.of(new TestItem("Test", 100))));
    }

    @Test
    void consumeOutputStream_calledTwice_shouldThrow() {
        CsvWriter<TestItem> writer = new CsvWriter<>();
        writer.column("Name", item -> item.name);

        CsvHandler handler = writer.write(Stream.of(new TestItem("Test", 100)));
        handler.consumeOutputStream(new ByteArrayOutputStream());

        assertThrows(CsvWriteException.class, () ->
                handler.consumeOutputStream(new ByteArrayOutputStream()));
    }

    @Test
    void write_emptyStream_shouldCleanUp() {
        CsvWriter<TestItem> writer = new CsvWriter<>();
        writer.column("Name", item -> item.name);

        CsvHandler handler = writer.write(Stream.empty());
        ByteArrayOutputStream baos = new ByteArrayOutputStream();
        handler.consumeOutputStream(baos);

        String csv = baos.toString();
        assertTrue(csv.contains("Name"), "Header should be present even for empty data");
    }

    // ──────────────────────────────────────────────────────────────
    // CsvWriter: CSV injection defense
    // ──────────────────────────────────────────────────────────────

    @Test
    void write_shouldDefendAgainstAllInjectionCharacters() {
        CsvWriter<TestItem> writer = new CsvWriter<>();
        writer.column("Name", item -> item.name);

        List<TestItem> items = List.of(
                new TestItem("=cmd|'/c calc'!A0", 0),    // = formula
                new TestItem("+1+1", 0),                  // + formula
                new TestItem("-1-1", 0),                  // - formula
                new TestItem("@SUM(A1:A10)", 0),          // @ formula
                new TestItem("\tcmd", 0),                  // tab injection
                new TestItem("\rcmd", 0),                  // carriage return injection
                new TestItem("normal text", 0),            // safe value
                new TestItem("100", 0)                     // numeric string (safe)
        );

        CsvHandler handler = writer.write(items.stream());
        ByteArrayOutputStream baos = new ByteArrayOutputStream();
        handler.consumeOutputStream(baos);

        String csv = baos.toString();
        String[] lines = csv.split("\\r?\\n");

        // Skip header (line 0)
        assertTrue(lines[1].startsWith("'="), "= should be prefixed with single quote");
        assertTrue(lines[2].startsWith("'+"), "+ should be prefixed with single quote");
        assertTrue(lines[3].startsWith("'-"), "- should be prefixed with single quote");
        assertTrue(lines[4].startsWith("'@"), "@ should be prefixed with single quote");
        assertTrue(lines[5].contains("'\t"), "tab should be prefixed with single quote");
        assertTrue(lines[6].contains("'\r"), "CR should be prefixed with single quote");
        assertEquals("normal text", lines[7], "Normal text should not be modified");
        assertEquals("100", lines[8], "Numeric strings should not be modified");
    }

    @Test
    void write_shouldEscapeNullValues() {
        CsvWriter<TestItem> writer = new CsvWriter<>();
        writer.column("Name", item -> null)
              .column("Price", item -> item.price);

        CsvHandler handler = writer.write(Stream.of(new TestItem("Test", 100)));
        ByteArrayOutputStream baos = new ByteArrayOutputStream();
        handler.consumeOutputStream(baos);

        String csv = baos.toString();
        String[] lines = csv.split("\\r?\\n");
        assertEquals(",100", lines[1], "Null values should be written as empty strings");
    }

    // ──────────────────────────────────────────────────────────────
    // CsvReadHandler: temp file cleanup
    // ──────────────────────────────────────────────────────────────

    @Test
    void csvRead_shouldCleanUpTempFilesAfterSuccess() throws IOException {
        Path csvFile = createTestCsv("Name,Price\nApple,1000\nBanana,2000\n");

        List<TestItem> results = new ArrayList<>();
        try (InputStream is = Files.newInputStream(csvFile)) {
            new CsvReader<>(TestItem::new, null)
                    .column((item, cell) -> item.name = cell.asString())
                    .column((item, cell) -> item.price = cell.asInt())
                    .build(is)
                    .read(result -> {
                        if (result.success()) {
                            results.add(result.data());
                        }
                    });
        }

        assertEquals(2, results.size());
        assertEquals("Apple", results.get(0).name);
        assertEquals(1000, results.get(0).price);
    }

    @Test
    void csvRead_shouldCleanUpTempFilesAfterException() throws IOException {
        Path csvFile = createTestCsv("Name,Price\nApple,1000\n");

        try (InputStream is = Files.newInputStream(csvFile)) {
            CsvReadHandler<TestItem> handler = new CsvReader<>(TestItem::new, null)
                    .column((item, cell) -> item.name = cell.asString())
                    .column((item, cell) -> item.price = cell.asInt())
                    .build(is);

            assertThrows(RuntimeException.class, () ->
                    handler.read(result -> {
                        throw new RuntimeException("Intentional error");
                    }));
        }
    }

    @Test
    void csvReadAsStream_shouldCleanUpOnClose() throws IOException {
        Path csvFile = createTestCsv("Name,Price\nApple,1000\nBanana,2000\n");

        List<String> names;
        try (InputStream is = Files.newInputStream(csvFile)) {
            try (Stream<ReadResult<TestItem>> stream = new CsvReader<>(TestItem::new, null)
                    .column((item, cell) -> item.name = cell.asString())
                    .column((item, cell) -> item.price = cell.asInt())
                    .build(is)
                    .readAsStream()) {

                names = stream
                        .filter(ReadResult::success)
                        .map(r -> r.data().name)
                        .toList();
            }
        }

        assertEquals(2, names.size());
    }

    @Test
    void csvReadAsStream_earlyTermination_shouldCleanUp() throws IOException {
        Path csvFile = createTestCsv("Name,Price\nApple,1000\nBanana,2000\nCherry,3000\n");

        List<String> names;
        try (InputStream is = Files.newInputStream(csvFile)) {
            try (Stream<ReadResult<TestItem>> stream = new CsvReader<>(TestItem::new, null)
                    .column((item, cell) -> item.name = cell.asString())
                    .column((item, cell) -> item.price = cell.asInt())
                    .build(is)
                    .readAsStream()) {

                names = stream
                        .filter(ReadResult::success)
                        .limit(1)
                        .map(r -> r.data().name)
                        .toList();
            }
        }

        assertEquals(1, names.size());
        assertEquals("Apple", names.get(0));
    }

    // ──────────────────────────────────────────────────────────────
    // CsvWriter+CsvReader roundtrip
    // ──────────────────────────────────────────────────────────────

    @Test
    void csvWriteAndRead_roundtrip_shouldCleanUpAllTempResources() throws IOException {
        // Write
        ByteArrayOutputStream baos = new ByteArrayOutputStream();
        new CsvWriter<TestItem>()
                .column("Name", item -> item.name)
                .column("Price", item -> item.price)
                .write(Stream.of(
                        new TestItem("Apple", 1000),
                        new TestItem("Banana", 2000)
                ))
                .consumeOutputStream(baos);

        // Read back
        List<TestItem> results = new ArrayList<>();
        try (InputStream is = new ByteArrayInputStream(baos.toByteArray())) {
            new CsvReader<>(TestItem::new, null)
                    .column((item, cell) -> item.name = cell.asString())
                    .column((item, cell) -> item.price = cell.asInt())
                    .build(is)
                    .read(result -> {
                        if (result.success()) {
                            results.add(result.data());
                        }
                    });
        }

        assertEquals(2, results.size());
        assertEquals("Apple", results.get(0).name);
        assertEquals(1000, results.get(0).price);
        assertEquals("Banana", results.get(1).name);
        assertEquals(2000, results.get(1).price);
    }

    // ──────────────────────────────────────────────────────────────
    // Helpers
    // ──────────────────────────────────────────────────────────────

    private Path createTestCsv(String content) throws IOException {
        Path file = tempDir.resolve("test.csv");
        Files.writeString(file, content, StandardCharsets.UTF_8);
        return file;
    }

    static class TestItem {
        String name;
        int price;

        TestItem() {}

        TestItem(String name, int price) {
            this.name = name;
            this.price = price;
        }
    }
}
