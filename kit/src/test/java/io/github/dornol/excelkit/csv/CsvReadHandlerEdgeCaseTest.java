package io.github.dornol.excelkit.csv;

import io.github.dornol.excelkit.core.ReadAbortException;
import io.github.dornol.excelkit.core.ReadResult;
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
                .column("Name", (p, cell) -> p.name = cell.asString())
                .column("Age", (p, cell) -> p.age = cell.asInt())
                .read(new ByteArrayInputStream(combined), results::add);

        assertEquals(1, results.size());
        assertTrue(results.get(0).success());
        assertEquals("Alice", results.get(0).data().name, "BOM should be stripped, Name should be 'Alice'");
        assertEquals(30, results.get(0).data().age);
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
        ).read(new ByteArrayInputStream(combined), r -> {
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
                .column("Name", (arr, cell) -> arr[0] = cell.asString())
                .column("Age", (arr, cell) -> arr[1] = cell.asString())
                .column("City", (arr, cell) -> arr[2] = cell.asString())
                .read(new ByteArrayInputStream(csv.getBytes(StandardCharsets.UTF_8)), results::add);

        assertEquals(1, results.size());
        assertEquals("Alice", results.get(0).data()[0]);
        // Row succeeded regardless of missing columns
        assertTrue(results.get(0).success());
    }

    // ============================================================
    // readAsStream without progress (progressCallback null)
    // ============================================================


    // ============================================================
    // readAsStream stream close without consuming
    // ============================================================


    // ============================================================
    // readAsStream with empty file (headerLine == null)
    // ============================================================


    // ============================================================
    // setter mode read (not mapping mode) — covers rowMapper == null branch
    // ============================================================
    @Test
    void read_setterMode_shouldWork() {
        String csv = "Name,Age\nAlice,30\nBob,25";

        List<ReadResult<MutablePerson>> results = new ArrayList<>();
        new CsvReader<>(MutablePerson::new, null)
                .column("Name", (p, cell) -> p.name = cell.asString())
                .column("Age", (p, cell) -> p.age = cell.asInt())
                .read(toInputStream(csv), results::add);

        assertEquals(2, results.size());
        assertTrue(results.get(0).success());
        assertEquals("Alice", results.get(0).data().name);
    }



    // ============================================================
    // Exception catch branches in read()
    // ============================================================




    @Test
    void read_setterMode_withProgress_shouldFireCallback() {
        String csv = "Name,Age\nA,1\nB,2\nC,3\nD,4";

        AtomicLong lastProgress = new AtomicLong(0);
        new CsvReader<>(MutablePerson::new, null)
                .column("Name", (p, cell) -> p.name = cell.asString())
                .column("Age", (p, cell) -> p.age = cell.asInt())
                .onProgress(2, (count, cursor) -> lastProgress.set(count))
                .read(toInputStream(csv), r -> {});

        assertEquals(4, lastProgress.get());
    }

    private InputStream toInputStream(String content) {
        return new ByteArrayInputStream(content.getBytes(StandardCharsets.UTF_8));
    }
}
