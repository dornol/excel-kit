package io.github.dornol.excelkit.excel;

import io.github.dornol.excelkit.core.ReadResult;
import org.junit.jupiter.api.Test;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.time.Duration;
import java.util.ArrayList;
import java.util.List;
import java.util.stream.IntStream;
import java.util.stream.Stream;

import static org.junit.jupiter.api.Assertions.*;

/**
 * Edge case tests for {@link ExcelReadHandler} to cover:
 * - Mapping mode (rowMapper)
 * - Missing columns in row
 * - Custom headerRowIndex > 0
 */
class ExcelReadHandlerEdgeCaseTest {

    record Item(String name, int value) {}

    private byte[] writeTestExcel() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        ExcelWriter.<Item>create()
                .column("Name", Item::name)
                .column("Value", i -> i.value, c -> c.type(ExcelDataType.INTEGER))
                .write(Stream.of(new Item("A", 10), new Item("B", 20), new Item("C", 30)))
                .writeTo(out);
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
        ).read(new ByteArrayInputStream(excel), r -> {
                    assertTrue(r.success());
                    results.add(r.data());
                });

        assertEquals(3, results.size());
        assertEquals("A", results.get(0).name());
        assertEquals(10, results.get(0).value());
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
        }).read(new ByteArrayInputStream(excel), results::add);

        assertEquals(3, results.size());
        for (var r : results) {
            assertFalse(r.success());
            assertNotNull(r.messages());
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

        ExcelWriter.<Item>create()
                .column("Name", Item::name)
                .column("Value", i -> i.value)
                .write(Stream.of(new Item("x", 1)))
                .writeTo(out);

        // Read back with setter mode - should work normally
        List<ReadResult<MutableItem>> results = new ArrayList<>();
        new ExcelReader<>(MutableItem::new, null)
                .column("Name", (t, cell) -> t.name = cell.asString())
                .column("Value", (t, cell) -> t.value = cell.asInt())
                .read(new ByteArrayInputStream(out.toByteArray()), results::add);

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
                .read(new ByteArrayInputStream(excel), r -> results.add(r.data()));

        assertEquals(3, results.size());
        assertEquals(2, lastProgress.get());
    }

    // ============================================================
    // countRows — pre-scan total row count
    // ============================================================
    @Test
    void countRows_shouldProvideTotalRowsInCursor() throws IOException {
        byte[] excel = writeTestExcel(); // 3 data rows

        List<Long> totals = new ArrayList<>();
        List<Long> processed = new ArrayList<>();
        ExcelReader.setter(ItemHolder::new)
                .column((h, c) -> h.name = c.asString())
                .column((h, c) -> h.value = c.asInt())
                .countRows()
                .onProgress(1, (count, cursor) -> {
                    processed.add(count);
                    totals.add(cursor.getTotalRows());
                })
                .read(new ByteArrayInputStream(excel), r -> {});

        assertEquals(List.of(1L, 2L, 3L), processed);
        // All callbacks should report the same totalRows = 3
        assertEquals(List.of(3L, 3L, 3L), totals);
    }

    @Test
    void countRows_mappingMode_shouldWork() throws IOException {
        byte[] excel = writeTestExcel(); // 3 data rows

        java.util.concurrent.atomic.AtomicLong totalRef = new java.util.concurrent.atomic.AtomicLong(-1);
        ExcelReader.<Item>mapping(row ->
                new Item(row.get("Name").asString(), row.get("Value").asInt())
        ).countRows()
                .onProgress(2, (count, cursor) -> totalRef.set(cursor.getTotalRows()))
                .read(new ByteArrayInputStream(excel), r -> {});

        assertEquals(3, totalRef.get());
    }

    @Test
    void withoutCountRows_cursorTotalRowsShouldBeNegative() throws IOException {
        byte[] excel = writeTestExcel();

        java.util.concurrent.atomic.AtomicLong totalRef = new java.util.concurrent.atomic.AtomicLong(999);
        ExcelReader.setter(ItemHolder::new)
                .column((h, c) -> h.name = c.asString())
                .column((h, c) -> h.value = c.asInt())
                .onProgress(1, (count, cursor) -> totalRef.set(cursor.getTotalRows()))
                .read(new ByteArrayInputStream(excel), r -> {});

        assertEquals(-1, totalRef.get());
    }

    public static class ItemHolder {
        String name;
        int value;
    }
}
