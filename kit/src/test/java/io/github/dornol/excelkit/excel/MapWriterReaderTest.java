package io.github.dornol.excelkit.excel;

import io.github.dornol.excelkit.csv.CsvDialect;
import io.github.dornol.excelkit.csv.CsvReadException;
import io.github.dornol.excelkit.csv.CsvReader;
import io.github.dornol.excelkit.csv.CsvWriter;
import io.github.dornol.excelkit.core.ReadResult;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.nio.charset.StandardCharsets;
import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.Set;
import java.util.concurrent.atomic.AtomicLong;
import java.util.stream.Stream;

import static org.junit.jupiter.api.Assertions.*;

class MapWriterReaderTest {

    @Test
    void excelMapWriter_shouldWriteMapData() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        var writer = ExcelWriter.forMap("Name", "Age");
        writer.write(Stream.of(
                Map.of("Name", "Alice", "Age", 30),
                Map.of("Name", "Bob", "Age", 25)
        )).writeTo(out);

        try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
            var sheet = wb.getSheetAt(0);
            assertEquals("Name", sheet.getRow(0).getCell(0).getStringCellValue());
            assertEquals("Age", sheet.getRow(0).getCell(1).getStringCellValue());
            assertEquals("Alice", sheet.getRow(1).getCell(0).getStringCellValue());
            assertEquals("Bob", sheet.getRow(2).getCell(0).getStringCellValue());
        }
    }

    @Test
    void excelMapReader_shouldReadAllColumns() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        ExcelWriter.forMap("Name", "City").write(Stream.of(
                Map.of("Name", "Alice", "City", "Seoul"),
                Map.of("Name", "Bob", "City", "Tokyo")
        )).writeTo(out);

        List<Map<String, String>> results = new ArrayList<>();
        ExcelReader.forMap()
                .read(new ByteArrayInputStream(out.toByteArray()), r -> results.add(r.data()));

        assertEquals(2, results.size());
        assertEquals("Alice", results.get(0).get("Name"));
        assertEquals("Seoul", results.get(0).get("City"));
        assertEquals("Bob", results.get(1).get("Name"));
        assertEquals("Tokyo", results.get(1).get("City"));
    }



    @Test
    void excelMapReader_withSheetIndex() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        try (ExcelWorkbook wb = ExcelWorkbook.create()) {
            wb.<Map<String, Object>>sheet("Sheet1")
                    .column("A", m -> m.get("A"))
                    .write(Stream.of(Map.of("A", "first")));
            wb.<Map<String, Object>>sheet("Sheet2")
                    .column("B", m -> m.get("B"))
                    .write(Stream.of(Map.of("B", "second")));
            wb.finish().writeTo(out);
        }

        List<Map<String, String>> results = new ArrayList<>();
        ExcelReader.forMap()
                .sheetIndex(1)
                .read(new ByteArrayInputStream(out.toByteArray()), r -> results.add(r.data()));

        assertEquals(1, results.size());
        assertEquals("second", results.get(0).get("B"));
    }

    @Test
    void csvMapWriter_shouldWriteMapData() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        var writer = CsvWriter.forMap("Name", "Age");
        writer.write(Stream.of(
                Map.of("Name", "Alice", "Age", 30)
        )).writeTo(out);

        String csv = out.toString(StandardCharsets.UTF_8);
        String[] lines = csv.split("\r?\n");
        // BOM + header + 1 data row
        assertTrue(lines.length >= 2, "Expected at least 2 lines (header + data), got " + lines.length);
        // Verify header line contains both column names in order
        String headerLine = lines[0].replace("\uFEFF", "");
        assertEquals("Name,Age", headerLine);
        // Verify data line
        assertEquals("Alice,30", lines[1]);
    }

    @Test
    void csvMapWriter_withDialect_TSV() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        CsvWriter.forMap("Name", "Age")
                .dialect(CsvDialect.TSV)
                .write(Stream.of(Map.of("Name", "Alice", "Age", 30)))
                .writeTo(out);

        String tsv = out.toString(StandardCharsets.UTF_8).replace("\uFEFF", "");
        String[] lines = tsv.split("\r?\n");
        assertEquals("Name\tAge", lines[0]);
        assertEquals("Alice\t30", lines[1]);
    }

    @Test
    void csvMapWriter_withDelimiter() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        CsvWriter.forMap("A", "B")
                .delimiter('|')
                .write(Stream.of(Map.of("A", "x", "B", "y")))
                .writeTo(out);

        String csv = out.toString(StandardCharsets.UTF_8).replace("\uFEFF", "");
        String[] lines = csv.split("\r?\n");
        assertEquals("A|B", lines[0]);
        assertEquals("x|y", lines[1]);
    }

    @Test
    void csvMapWriter_withBomDisabled() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        CsvWriter.forMap("A")
                .bom(false)
                .write(Stream.of(Map.of("A", "val")))
                .writeTo(out);

        byte[] bytes = out.toByteArray();
        assertFalse(bytes[0] == (byte) 0xEF && bytes[1] == (byte) 0xBB && bytes[2] == (byte) 0xBF,
                "Should not start with BOM");
        String csv = out.toString(StandardCharsets.UTF_8);
        assertTrue(csv.startsWith("A"), "Should start directly with header");
    }

    @Test
    void csvMapReader_shouldReadAllColumns() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        CsvWriter.forMap("Name", "City").write(Stream.of(
                Map.of("Name", "Alice", "City", "Seoul"),
                Map.of("Name", "Bob", "City", "Tokyo")
        )).writeTo(out);

        List<ReadResult<Map<String, String>>> results = new ArrayList<>();
        CsvReader.forMap()
                .read(new ByteArrayInputStream(out.toByteArray()), results::add);

        assertEquals(2, results.size());

        // Verify first row
        assertTrue(results.get(0).success());
        assertNull(results.get(0).messages());
        Map<String, String> row0 = results.get(0).data();
        assertEquals(Set.of("Name", "City"), row0.keySet());
        assertEquals("Alice", row0.get("Name"));
        assertEquals("Seoul", row0.get("City"));

        // Verify second row
        assertTrue(results.get(1).success());
        Map<String, String> row1 = results.get(1).data();
        assertEquals(Set.of("Name", "City"), row1.keySet());
        assertEquals("Bob", row1.get("Name"));
        assertEquals("Tokyo", row1.get("City"));
    }





    @Test
    void csvMapReader_withDelimiter() throws IOException {
        String tsv = "Name\tAge\nAlice\t30\nBob\t25\n";
        List<Map<String, String>> results = new ArrayList<>();
        CsvReader.forMap()
                .delimiter('\t')
                .read(new ByteArrayInputStream(tsv.getBytes()), r -> results.add(r.data()));

        assertEquals(2, results.size());
        assertEquals(Set.of("Name", "Age"), results.get(0).keySet());
        assertEquals("Alice", results.get(0).get("Name"));
        assertEquals("30", results.get(0).get("Age"));
        assertEquals("Bob", results.get(1).get("Name"));
        assertEquals("25", results.get(1).get("Age"));
    }

    @Test
    void csvMapReader_withDialect_TSV() throws IOException {
        String tsv = "Name\tAge\nAlice\t30\n";
        List<Map<String, String>> results = new ArrayList<>();
        CsvReader.forMap()
                .dialect(CsvDialect.TSV)
                .read(new ByteArrayInputStream(tsv.getBytes()), r -> results.add(r.data()));

        assertEquals(1, results.size());
        assertEquals("Alice", results.get(0).get("Name"));
        assertEquals("30", results.get(0).get("Age"));
    }

    @Test
    void csvMapReader_withHeaderRowIndex() throws IOException {
        String csv = "skip this line\nName,Age\nAlice,30\n";
        List<Map<String, String>> results = new ArrayList<>();
        CsvReader.forMap()
                .headerRowIndex(1)
                .read(new ByteArrayInputStream(csv.getBytes()), r -> results.add(r.data()));

        assertEquals(1, results.size());
        assertEquals(Set.of("Name", "Age"), results.get(0).keySet());
        assertEquals("Alice", results.get(0).get("Name"));
        assertEquals("30", results.get(0).get("Age"));
    }

    @Test
    void csvMapReader_withProgress() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        CsvWriter.forMap("Name").write(Stream.of(
                Map.of("Name", "A"),
                Map.of("Name", "B"),
                Map.of("Name", "C"),
                Map.of("Name", "D")
        )).writeTo(out);

        List<Long> progressCounts = new ArrayList<>();
        List<Map<String, String>> results = new ArrayList<>();
        CsvReader.forMap()
                .onProgress(2, (count, total) -> progressCounts.add(count))
                .read(new ByteArrayInputStream(out.toByteArray()), r -> results.add(r.data()));

        // Verify data was actually read correctly
        assertEquals(4, results.size());
        assertEquals("A", results.get(0).get("Name"));
        assertEquals("D", results.get(3).get("Name"));

        // Verify progress was called at row 2 and 4
        assertEquals(List.of(2L, 4L), progressCounts);
    }

    @Test
    void csvMapReader_roundTripWithCsvMapWriter() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        CsvWriter.forMap("Name", "Age", "Email").write(Stream.of(
                Map.of("Name", "Alice", "Age", 30, "Email", "alice@test.com"),
                Map.of("Name", "Bob", "Age", 25, "Email", "bob@test.com")
        )).writeTo(out);

        List<ReadResult<Map<String, String>>> results = new ArrayList<>();
        CsvReader.forMap()
                .read(new ByteArrayInputStream(out.toByteArray()), results::add);

        assertEquals(2, results.size());

        // First row - all fields
        assertTrue(results.get(0).success());
        Map<String, String> row0 = results.get(0).data();
        assertEquals(Set.of("Name", "Age", "Email"), row0.keySet());
        assertEquals("Alice", row0.get("Name"));
        assertEquals("30", row0.get("Age"));
        assertEquals("alice@test.com", row0.get("Email"));

        // Second row - all fields
        assertTrue(results.get(1).success());
        Map<String, String> row1 = results.get(1).data();
        assertEquals("Bob", row1.get("Name"));
        assertEquals("25", row1.get("Age"));
        assertEquals("bob@test.com", row1.get("Email"));
    }

    @Test
    void csvMapReader_emptyFile_throwsException() {
        String csv = "";
        assertThrows(CsvReadException.class, () ->
                CsvReader.forMap()
                        .read(new ByteArrayInputStream(csv.getBytes()), r -> {}));
    }

    @Test
    void csvMapReader_headerOnly_returnsNoRows() throws IOException {
        String csv = "Name,Age\n";
        List<Map<String, String>> results = new ArrayList<>();
        CsvReader.forMap()
                .read(new ByteArrayInputStream(csv.getBytes()), r -> results.add(r.data()));

        assertTrue(results.isEmpty());
    }

    @Test
    void csvMapReader_fewerDataColumnsThanHeaders() throws IOException {
        // Row has fewer columns than header. Behavior preserved from the deleted CsvMapReader:
        // the map is truncated at min(headerCount, cellCount), so trailing header keys are absent.
        String csv = "Name,Age,City\nAlice,30\n";
        List<Map<String, String>> results = new ArrayList<>();
        CsvReader.forMap()
                .read(new ByteArrayInputStream(csv.getBytes()), r -> results.add(r.data()));

        assertEquals(1, results.size());
        assertEquals("Alice", results.get(0).get("Name"));
        assertEquals("30", results.get(0).get("Age"));
        assertFalse(results.get(0).containsKey("City"));
    }

    @Test
    void csvMapReader_withBOM() throws IOException {
        byte[] bom = {(byte) 0xEF, (byte) 0xBB, (byte) 0xBF};
        byte[] content = "Name,Age\nAlice,30\n".getBytes(StandardCharsets.UTF_8);
        byte[] csvBytes = new byte[bom.length + content.length];
        System.arraycopy(bom, 0, csvBytes, 0, bom.length);
        System.arraycopy(content, 0, csvBytes, bom.length, content.length);

        List<Map<String, String>> results = new ArrayList<>();
        CsvReader.forMap()
                .read(new ByteArrayInputStream(csvBytes), r -> results.add(r.data()));

        assertEquals(1, results.size());
        // BOM should be stripped — key should be "Name", not "\uFEFFName"
        assertTrue(results.get(0).containsKey("Name"), "BOM should be stripped from header");
        assertEquals("Alice", results.get(0).get("Name"));
        assertEquals("30", results.get(0).get("Age"));
    }

    @Test
    void csvMapReader_onProgress_invalidInterval_throwsException() {
        assertThrows(IllegalArgumentException.class, () ->
                CsvReader.forMap().onProgress(0, (count, total) -> {}));
        assertThrows(IllegalArgumentException.class, () ->
                CsvReader.forMap().onProgress(-1, (count, total) -> {}));
    }

    @Test
    void csvMapReader_preservesColumnOrder() throws IOException {
        String csv = "C,A,B\n3,1,2\n";
        List<Map<String, String>> results = new ArrayList<>();
        CsvReader.forMap()
                .read(new ByteArrayInputStream(csv.getBytes()), r -> results.add(r.data()));

        assertEquals(1, results.size());
        // LinkedHashMap preserves insertion order
        List<String> keys = new ArrayList<>(results.get(0).keySet());
        assertEquals(List.of("C", "A", "B"), keys);
        assertEquals("3", results.get(0).get("C"));
        assertEquals("1", results.get(0).get("A"));
        assertEquals("2", results.get(0).get("B"));
    }

    @Test
    void excelMapWriter_missingKeyReturnsNull() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        Map<String, Object> row = new LinkedHashMap<>();
        row.put("Name", "Alice");
        // "Age" key is missing

        ExcelWriter.forMap("Name", "Age")
                .write(Stream.of(row))
                .writeTo(out);

        try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
            assertEquals("Alice", wb.getSheetAt(0).getRow(1).getCell(0).getStringCellValue());
            // Missing key should result in empty cell
            assertEquals("", wb.getSheetAt(0).getRow(1).getCell(1).getStringCellValue());
        }
    }
}
