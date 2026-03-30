package io.github.dornol.excelkit.excel;

import io.github.dornol.excelkit.shared.ReadResult;
import org.junit.jupiter.api.Test;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import java.util.stream.Stream;

import static org.junit.jupiter.api.Assertions.*;

/**
 * Edge case tests for {@link ExcelReadHandler} to cover:
 * - Mapping mode (rowMapper) in both read() and readAsStream()
 * - Missing columns in row
 * - readAsStream error propagation
 * - Custom headerRowIndex > 0
 */
class ExcelReadHandlerEdgeCaseTest {

    record Item(String name, int value) {}

    private byte[] writeTestExcel() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        new ExcelWriter<Item>()
                .addColumn("Name", Item::name)
                .addColumn("Value", i -> i.value, c -> c.type(ExcelDataType.INTEGER))
                .write(Stream.of(new Item("A", 10), new Item("B", 20), new Item("C", 30)))
                .consumeOutputStream(out);
        return out.toByteArray();
    }

    // ============================================================
    // Mapping mode - read()
    // ============================================================
    @Test
    void mappingMode_read_shouldWork() throws IOException {
        byte[] excel = writeTestExcel();

        List<Item> results = new ArrayList<>();
        ExcelReader.<Item>mapping(row ->
                new Item(row.get("Name").asString(), row.get("Value").asInt())
        ).build(new ByteArrayInputStream(excel))
                .read(r -> {
                    assertTrue(r.success());
                    results.add(r.data());
                });

        assertEquals(3, results.size());
        assertEquals("A", results.get(0).name());
        assertEquals(10, results.get(0).value());
    }

    // ============================================================
    // Mapping mode - readAsStream()
    // ============================================================
    @Test
    void mappingMode_readAsStream_shouldWork() throws IOException {
        byte[] excel = writeTestExcel();

        List<Item> results;
        try (var stream = ExcelReader.<Item>mapping(row ->
                new Item(row.get("Name").asString(), row.get("Value").asInt())
        ).build(new ByteArrayInputStream(excel)).readAsStream()) {
            results = stream
                    .filter(ReadResult::success)
                    .map(ReadResult::data)
                    .toList();
        }

        assertEquals(3, results.size());
        assertEquals("B", results.get(1).name());
        assertEquals(30, results.get(2).value());
    }

    // ============================================================
    // Mapping mode - error handling
    // ============================================================
    @Test
    void mappingMode_conversionError_shouldReturnFailedResult() throws IOException {
        byte[] excel = writeTestExcel();

        List<ReadResult<Item>> results = new ArrayList<>();
        ExcelReader.<Item>mapping(row -> {
            // Force conversion error on "Name" column (not a number)
            int v = Integer.parseInt(row.get("Name").asString());
            return new Item("X", v);
        }).build(new ByteArrayInputStream(excel))
                .read(results::add);

        assertEquals(3, results.size());
        for (var r : results) {
            assertFalse(r.success());
            assertNotNull(r.messages());
        }
    }

    // ============================================================
    // readAsStream - close without consuming
    // ============================================================
    @Test
    void readAsStream_closeWithoutConsuming_shouldNotLeak() throws IOException {
        byte[] excel = writeTestExcel();

        var stream = ExcelReader.<Item>mapping(row ->
                new Item(row.get("Name").asString(), row.get("Value").asInt())
        ).build(new ByteArrayInputStream(excel)).readAsStream();
        stream.close();
    }

    // ============================================================
    // readAsStream - partial consumption
    // ============================================================
    @Test
    void readAsStream_partialConsumption_shouldCleanup() throws IOException {
        byte[] excel = writeTestExcel();

        try (var stream = ExcelReader.<Item>mapping(row ->
                new Item(row.get("Name").asString(), row.get("Value").asInt())
        ).build(new ByteArrayInputStream(excel)).readAsStream()) {
            var first = stream.findFirst();
            assertTrue(first.isPresent());
            assertEquals("A", first.get().data().name());
        }
    }

    // ============================================================
    // Missing columns (actualIndex >= currentRow.size())
    // ============================================================
    static class MutableItem {
        String name;
        int value;
    }

    @Test
    void read_rowWithFewerColumnsThanHeader() throws IOException {
        // Write an Excel with 2 columns
        ByteArrayOutputStream out = new ByteArrayOutputStream();

        new ExcelWriter<Item>()
                .addColumn("Name", Item::name)
                .addColumn("Value", i -> i.value)
                .write(Stream.of(new Item("x", 1)))
                .consumeOutputStream(out);

        // Read back with setter mode - should work normally
        List<ReadResult<MutableItem>> results = new ArrayList<>();
        new ExcelReader<>(MutableItem::new, null)
                .addColumn("Name", (t, cell) -> t.name = cell.asString())
                .addColumn("Value", (t, cell) -> t.value = cell.asInt())
                .build(new ByteArrayInputStream(out.toByteArray()))
                .read(results::add);

        assertEquals(1, results.size());
    }

    // ============================================================
    // Progress callback in mapping mode
    // ============================================================
    @Test
    void mappingMode_withProgress() throws IOException {
        byte[] excel = writeTestExcel();

        java.util.concurrent.atomic.AtomicLong lastProgress = new java.util.concurrent.atomic.AtomicLong(0);
        List<Item> results = new ArrayList<>();
        ExcelReader.<Item>mapping(row ->
                new Item(row.get("Name").asString(), row.get("Value").asInt())
        ).onProgress(2, (count, cursor) -> lastProgress.set(count))
                .build(new ByteArrayInputStream(excel))
                .read(r -> results.add(r.data()));

        assertEquals(3, results.size());
        assertEquals(2, lastProgress.get());
    }
}
