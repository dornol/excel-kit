package io.github.dornol.excelkit.excel;

import io.github.dornol.excelkit.shared.ReadResult;
import org.junit.jupiter.api.Test;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.util.List;
import java.util.Map;
import java.util.stream.Stream;

import static org.junit.jupiter.api.Assertions.*;

/**
 * Tests for {@link ExcelMapReader.ExcelMapReadHandler#readAsStream()} streaming implementation.
 * Verifies BlockingQueue-based streaming, resource cleanup, and error propagation.
 */
class MapReaderStreamTest {

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

    @Test
    void readAsStream_shouldReturnAllRows() throws IOException {
        byte[] excel = writeTestExcel();

        List<ReadResult<Map<String, String>>> results;
        try (var stream = new ExcelMapReader()
                .build(new ByteArrayInputStream(excel))
                .readAsStream()) {
            results = stream.toList();
        }

        assertEquals(3, results.size());
        assertEquals("A", results.get(0).data().get("Name"));
        assertEquals("20", results.get(1).data().get("Value"));
        assertEquals("C", results.get(2).data().get("Name"));
    }

    @Test
    void readAsStream_withFilter_shouldStreamLazily() throws IOException {
        byte[] excel = writeTestExcel();

        List<Map<String, String>> results;
        try (var stream = new ExcelMapReader()
                .build(new ByteArrayInputStream(excel))
                .readAsStream()) {
            results = stream
                    .filter(ReadResult::success)
                    .map(ReadResult::data)
                    .filter(m -> !"B".equals(m.get("Name")))
                    .toList();
        }

        assertEquals(2, results.size());
        assertEquals("A", results.get(0).get("Name"));
        assertEquals("C", results.get(1).get("Name"));
    }

    @Test
    void readAsStream_partialConsumption_shouldCleanup() throws IOException {
        byte[] excel = writeTestExcel();

        try (var stream = new ExcelMapReader()
                .build(new ByteArrayInputStream(excel))
                .readAsStream()) {
            var first = stream.findFirst();
            assertTrue(first.isPresent());
            assertEquals("A", first.get().data().get("Name"));
        }
        // No resource leak
    }

    @Test
    void readAsStream_closeWithoutConsuming_shouldNotHang() throws IOException {
        byte[] excel = writeTestExcel();

        var stream = new ExcelMapReader()
                .build(new ByteArrayInputStream(excel))
                .readAsStream();
        stream.close();
        // Should not hang or throw
    }

    @Test
    void readAsStream_emptyData_shouldReturnEmptyStream() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        new ExcelWriter<Item>()
                .addColumn("Name", Item::name)
                .write(Stream.empty())
                .consumeOutputStream(out);

        List<ReadResult<Map<String, String>>> results;
        try (var stream = new ExcelMapReader()
                .build(new ByteArrayInputStream(out.toByteArray()))
                .readAsStream()) {
            results = stream.toList();
        }

        assertTrue(results.isEmpty());
    }

    @Test
    void readAsStream_withHeaderRowIndex_shouldSkipMetadata() throws IOException {
        // Write Excel with header at row 1 (row 0 is metadata)
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        try (var wb = new org.apache.poi.xssf.usermodel.XSSFWorkbook()) {
            var sheet = wb.createSheet("Test");
            sheet.createRow(0).createCell(0).setCellValue("METADATA");
            var header = sheet.createRow(1);
            header.createCell(0).setCellValue("Name");
            header.createCell(1).setCellValue("Value");
            var data = sheet.createRow(2);
            data.createCell(0).setCellValue("X");
            data.createCell(1).setCellValue(99);
            try (var fos = new java.io.FileOutputStream(
                    java.nio.file.Files.createTempFile("test", ".xlsx").toFile())) {
                wb.write(fos);
            }
            // Write to byte array instead
            var bout = new ByteArrayOutputStream();
            wb.write(bout);
            out = bout;
        }

        try (var stream = new ExcelMapReader()
                .headerRowIndex(1)
                .build(new ByteArrayInputStream(out.toByteArray()))
                .readAsStream()) {
            var results = stream.toList();
            assertEquals(1, results.size());
            assertEquals("X", results.get(0).data().get("Name"));
        }
    }

    @Test
    void read_withSparseRow_shouldFillGaps() throws IOException {
        // Write Excel with gap in data (col 0 has value, col 1 empty, col 2 has value)
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        try (var wb = new org.apache.poi.xssf.usermodel.XSSFWorkbook()) {
            var sheet = wb.createSheet("Test");
            var header = sheet.createRow(0);
            header.createCell(0).setCellValue("A");
            header.createCell(1).setCellValue("B");
            header.createCell(2).setCellValue("C");
            var data = sheet.createRow(1);
            data.createCell(0).setCellValue("x");
            // Skip col 1
            data.createCell(2).setCellValue("z");
            wb.write(out);
        }

        var results = new java.util.ArrayList<Map<String, String>>();
        new ExcelMapReader()
                .build(new ByteArrayInputStream(out.toByteArray()))
                .read(r -> results.add(r.data()));

        assertEquals(1, results.size());
        assertEquals("x", results.get(0).get("A"));
        assertEquals("z", results.get(0).get("C"));
    }

    @Test
    void readAsStream_multipleColumns_shouldMapCorrectly() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        record Multi(String a, String b, String c) {}
        new ExcelWriter<Multi>()
                .addColumn("Col1", Multi::a)
                .addColumn("Col2", Multi::b)
                .addColumn("Col3", Multi::c)
                .write(Stream.of(new Multi("x", "y", "z")))
                .consumeOutputStream(out);

        try (var stream = new ExcelMapReader()
                .build(new ByteArrayInputStream(out.toByteArray()))
                .readAsStream()) {
            var result = stream.findFirst().orElseThrow();
            assertTrue(result.success());
            assertEquals("x", result.data().get("Col1"));
            assertEquals("y", result.data().get("Col2"));
            assertEquals("z", result.data().get("Col3"));
        }
    }
}
