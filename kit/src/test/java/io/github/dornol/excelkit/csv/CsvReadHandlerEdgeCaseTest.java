package io.github.dornol.excelkit.csv;

import io.github.dornol.excelkit.shared.ReadResult;
import org.junit.jupiter.api.Test;

import java.io.ByteArrayInputStream;
import java.io.InputStream;
import java.nio.charset.StandardCharsets;
import java.util.ArrayList;
import java.util.List;
import java.util.concurrent.atomic.AtomicLong;
import java.util.stream.Stream;

import static org.junit.jupiter.api.Assertions.*;

/**
 * Edge case tests for {@link CsvReadHandler} to cover:
 * - Mapping mode in readAsStream()
 * - BOM handling
 * - Missing columns in rows
 * - Progress callback in stream mode
 * - readAsStream with progress null
 */
class CsvReadHandlerEdgeCaseTest {

    record Person(String name, int age) {}

    // ============================================================
    // readAsStream in mapping mode
    // ============================================================
    @Test
    void readAsStream_mappingMode_shouldWork() {
        String csv = "Name,Age\nAlice,30\nBob,25";

        List<Person> results;
        try (Stream<ReadResult<Person>> stream = CsvReader.<Person>mapping(row ->
                new Person(row.get("Name").asString(), row.get("Age").asInt())
        ).build(toInputStream(csv)).readAsStream()) {
            results = stream
                    .filter(ReadResult::success)
                    .map(ReadResult::data)
                    .toList();
        }

        assertEquals(2, results.size());
        assertEquals("Alice", results.get(0).name());
        assertEquals(25, results.get(1).age());
    }

    @Test
    void readAsStream_mappingMode_withProgress() {
        String csv = "Name,Age\nA,1\nB,2\nC,3\nD,4";

        AtomicLong lastProgress = new AtomicLong(0);
        try (var stream = CsvReader.<Person>mapping(row ->
                new Person(row.get("Name").asString(), row.get("Age").asInt())
        ).onProgress(2, (count, cursor) -> lastProgress.set(count))
                .build(toInputStream(csv)).readAsStream()) {
            stream.forEach(r -> {});
        }

        assertEquals(4, lastProgress.get());
    }

    // ============================================================
    // BOM handling
    // ============================================================
    static class MutablePerson {
        String name;
        int age;
    }

    @Test
    void readAsStream_withBOM_shouldStripBOM() {
        // UTF-8 BOM + csv content
        byte[] bom = new byte[]{(byte) 0xEF, (byte) 0xBB, (byte) 0xBF};
        String csvContent = "Name,Age\nAlice,30";
        byte[] csvBytes = csvContent.getBytes(StandardCharsets.UTF_8);
        byte[] combined = new byte[bom.length + csvBytes.length];
        System.arraycopy(bom, 0, combined, 0, bom.length);
        System.arraycopy(csvBytes, 0, combined, bom.length, csvBytes.length);

        List<ReadResult<MutablePerson>> results = new ArrayList<>();
        new CsvReader<>(MutablePerson::new, null)
                .addColumn("Name", (p, cell) -> p.name = cell.asString())
                .addColumn("Age", (p, cell) -> p.age = cell.asInt())
                .build(new ByteArrayInputStream(combined))
                .read(results::add);

        assertEquals(1, results.size());
        assertTrue(results.get(0).success());
    }

    @Test
    void read_withBOM_mappingMode() {
        byte[] bom = new byte[]{(byte) 0xEF, (byte) 0xBB, (byte) 0xBF};
        String csvContent = "Name,Age\nAlice,30";
        byte[] csvBytes = csvContent.getBytes(StandardCharsets.UTF_8);
        byte[] combined = new byte[bom.length + csvBytes.length];
        System.arraycopy(bom, 0, combined, 0, bom.length);
        System.arraycopy(csvBytes, 0, combined, bom.length, csvBytes.length);

        List<Person> results = new ArrayList<>();
        CsvReader.<Person>mapping(row ->
                new Person(row.get("Name").asString(), row.get("Age").asInt())
        ).build(new ByteArrayInputStream(combined))
                .read(r -> {
                    assertTrue(r.success());
                    results.add(r.data());
                });

        assertEquals(1, results.size());
        assertEquals("Alice", results.get(0).name());
    }

    // ============================================================
    // Missing columns in row (actualIndex >= line.length)
    // ============================================================
    @Test
    void read_missingColumnsInRow_shouldHandleGracefully() {
        // Header has 3 columns, data row has only 1
        String csv = "Name,Age,City\nAlice";

        List<ReadResult<String[]>> results = new ArrayList<>();
        new CsvReader<>(() -> new String[3], null)
                .addColumn("Name", (arr, cell) -> arr[0] = cell.asString())
                .addColumn("Age", (arr, cell) -> arr[1] = cell.asString())
                .addColumn("City", (arr, cell) -> arr[2] = cell.asString())
                .build(new ByteArrayInputStream(csv.getBytes(StandardCharsets.UTF_8)))
                .read(results::add);

        assertEquals(1, results.size());
        assertEquals("Alice", results.get(0).data()[0]);
        // Row succeeded regardless of missing columns
        assertTrue(results.get(0).success());
    }

    // ============================================================
    // readAsStream without progress (progressCallback null)
    // ============================================================
    @Test
    void readAsStream_withoutProgress_shouldWork() {
        String csv = "Name,Age\nAlice,30\nBob,25";

        List<Person> results;
        try (var stream = CsvReader.<Person>mapping(row ->
                new Person(row.get("Name").asString(), row.get("Age").asInt())
        ).build(toInputStream(csv)).readAsStream()) {
            results = stream.map(ReadResult::data).toList();
        }

        assertEquals(2, results.size());
    }

    // ============================================================
    // readAsStream stream close without consuming
    // ============================================================
    @Test
    void readAsStream_closeWithoutConsuming_shouldNotLeak() {
        String csv = "Name,Age\nAlice,30\nBob,25";

        var stream = CsvReader.<Person>mapping(row ->
                new Person(row.get("Name").asString(), row.get("Age").asInt())
        ).build(toInputStream(csv)).readAsStream();
        stream.close();
        // No exception means cleanup succeeded
    }

    // ============================================================
    // readAsStream with empty file (headerLine == null)
    // ============================================================
    @Test
    void readAsStream_emptyFile_throwsCsvReadException() {
        String csv = ""; // completely empty

        var handler = CsvReader.<Person>mapping(row ->
                new Person(row.get("Name").asString(), row.get("Age").asInt())
        ).build(toInputStream(csv));

        assertThrows(io.github.dornol.excelkit.csv.CsvReadException.class,
                handler::readAsStream);
    }

    // ============================================================
    // setter mode read (not mapping mode) — covers rowMapper == null branch
    // ============================================================
    @Test
    void read_setterMode_shouldWork() {
        String csv = "Name,Age\nAlice,30\nBob,25";

        List<ReadResult<MutablePerson>> results = new ArrayList<>();
        new CsvReader<>(MutablePerson::new, null)
                .addColumn("Name", (p, cell) -> p.name = cell.asString())
                .addColumn("Age", (p, cell) -> p.age = cell.asInt())
                .build(toInputStream(csv))
                .read(results::add);

        assertEquals(2, results.size());
        assertTrue(results.get(0).success());
        assertEquals("Alice", results.get(0).data().name);
    }

    @Test
    void readAsStream_setterMode_shouldWork() {
        String csv = "Name,Age\nAlice,30";

        List<ReadResult<MutablePerson>> results;
        try (var stream = new CsvReader<>(MutablePerson::new, null)
                .addColumn("Name", (p, cell) -> p.name = cell.asString())
                .addColumn("Age", (p, cell) -> p.age = cell.asInt())
                .build(toInputStream(csv))
                .readAsStream()) {
            results = stream.toList();
        }

        assertEquals(1, results.size());
        assertEquals("Alice", results.get(0).data().name);
    }

    // ============================================================
    // Exception catch branches in read()
    // ============================================================
    @Test
    void read_consumerThrowsCsvReadException_shouldPropagate() {
        String csv = "Name,Age\nAlice,30";

        var handler = CsvReader.<Person>mapping(row ->
                new Person(row.get("Name").asString(), row.get("Age").asInt())
        ).build(toInputStream(csv));

        assertThrows(io.github.dornol.excelkit.csv.CsvReadException.class, () ->
                handler.read(r -> {
                    throw new io.github.dornol.excelkit.csv.CsvReadException("test error");
                }));
    }

    @Test
    void read_consumerThrowsReadAbortException_shouldPropagate() {
        String csv = "Name,Age\nAlice,30";

        var handler = CsvReader.<Person>mapping(row ->
                new Person(row.get("Name").asString(), row.get("Age").asInt())
        ).build(toInputStream(csv));

        assertThrows(io.github.dornol.excelkit.shared.ReadAbortException.class, () ->
                handler.read(r -> {
                    throw new io.github.dornol.excelkit.shared.ReadAbortException("abort!");
                }));
    }

    @Test
    void read_setterMode_withProgress_shouldFireCallback() {
        String csv = "Name,Age\nA,1\nB,2\nC,3\nD,4";

        AtomicLong lastProgress = new AtomicLong(0);
        new CsvReader<>(MutablePerson::new, null)
                .addColumn("Name", (p, cell) -> p.name = cell.asString())
                .addColumn("Age", (p, cell) -> p.age = cell.asInt())
                .onProgress(2, (count, cursor) -> lastProgress.set(count))
                .build(toInputStream(csv))
                .read(r -> {});

        assertEquals(4, lastProgress.get());
    }

    private InputStream toInputStream(String content) {
        return new ByteArrayInputStream(content.getBytes(StandardCharsets.UTF_8));
    }
}
