package io.github.dornol.excelkit.excel;

import io.github.dornol.excelkit.csv.CsvMapWriter;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.stream.Stream;

import static org.junit.jupiter.api.Assertions.*;

class MapWriterReaderTest {

    @Test
    void excelMapWriter_shouldWriteMapData() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        ExcelMapWriter writer = new ExcelMapWriter("Name", "Age");
        writer.write(Stream.of(
                Map.of("Name", "Alice", "Age", 30),
                Map.of("Name", "Bob", "Age", 25)
        )).consumeOutputStream(out);

        try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
            var sheet = wb.getSheetAt(0);
            assertEquals("Name", sheet.getRow(0).getCell(0).getStringCellValue());
            assertEquals("Age", sheet.getRow(0).getCell(1).getStringCellValue());
            assertEquals("Alice", sheet.getRow(1).getCell(0).getStringCellValue());
            assertEquals("Bob", sheet.getRow(2).getCell(0).getStringCellValue());
        }
    }

    @Test
    void excelMapWriter_withCustomWriter() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        ExcelMapWriter writer = new ExcelMapWriter(
                new ExcelWriter<>(ExcelColor.LIGHT_BLUE),
                "Name", "Score"
        );
        writer.write(Stream.of(
                Map.of("Name", "Alice", "Score", 95)
        )).consumeOutputStream(out);

        try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
            assertEquals("Alice", wb.getSheetAt(0).getRow(1).getCell(0).getStringCellValue());
        }
    }

    @Test
    void excelMapWriter_writerAccessor() {
        ExcelMapWriter writer = new ExcelMapWriter("A", "B");
        assertNotNull(writer.writer());
    }

    @Test
    void excelMapReader_shouldReadAllColumns() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        new ExcelMapWriter("Name", "City").write(Stream.of(
                Map.of("Name", "Alice", "City", "Seoul"),
                Map.of("Name", "Bob", "City", "Tokyo")
        )).consumeOutputStream(out);

        List<Map<String, String>> results = new ArrayList<>();
        new ExcelMapReader()
                .build(new ByteArrayInputStream(out.toByteArray()))
                .read(r -> results.add(r.data()));

        assertEquals(2, results.size());
        assertEquals("Alice", results.get(0).get("Name"));
        assertEquals("Seoul", results.get(0).get("City"));
        assertEquals("Bob", results.get(1).get("Name"));
        assertEquals("Tokyo", results.get(1).get("City"));
    }

    @Test
    void excelMapReader_readAsStream() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        new ExcelMapWriter("Name").write(Stream.of(
                Map.of("Name", "Alice"),
                Map.of("Name", "Bob")
        )).consumeOutputStream(out);

        var results = new ExcelMapReader()
                .build(new ByteArrayInputStream(out.toByteArray()))
                .readAsStream()
                .toList();

        assertEquals(2, results.size());
        assertEquals("Alice", results.get(0).data().get("Name"));
    }

    @Test
    void excelMapReader_withSheetIndex() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        try (ExcelWorkbook wb = new ExcelWorkbook()) {
            wb.<Map<String, Object>>sheet("Sheet1")
                    .column("A", m -> m.get("A"))
                    .write(Stream.of(Map.of("A", "first")));
            wb.<Map<String, Object>>sheet("Sheet2")
                    .column("B", m -> m.get("B"))
                    .write(Stream.of(Map.of("B", "second")));
            wb.finish().consumeOutputStream(out);
        }

        List<Map<String, String>> results = new ArrayList<>();
        new ExcelMapReader()
                .sheetIndex(1)
                .build(new ByteArrayInputStream(out.toByteArray()))
                .read(r -> results.add(r.data()));

        assertEquals(1, results.size());
        assertEquals("second", results.get(0).get("B"));
    }

    @Test
    void csvMapWriter_shouldWriteMapData() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        CsvMapWriter writer = new CsvMapWriter("Name", "Age");
        writer.write(Stream.of(
                Map.of("Name", "Alice", "Age", 30)
        )).consumeOutputStream(out);

        String csv = out.toString();
        assertTrue(csv.contains("Name"));
        assertTrue(csv.contains("Age"));
        assertTrue(csv.contains("Alice"));
        assertTrue(csv.contains("30"));
    }

    @Test
    void csvMapWriter_writerAccessor() {
        CsvMapWriter writer = new CsvMapWriter("A");
        assertNotNull(writer.writer());
    }

    @Test
    void excelMapWriter_missingKeyReturnsNull() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        Map<String, Object> row = new LinkedHashMap<>();
        row.put("Name", "Alice");
        // "Age" key is missing

        new ExcelMapWriter("Name", "Age")
                .write(Stream.of(row))
                .consumeOutputStream(out);

        try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
            assertEquals("Alice", wb.getSheetAt(0).getRow(1).getCell(0).getStringCellValue());
            // Missing key should result in empty cell
            assertEquals("", wb.getSheetAt(0).getRow(1).getCell(1).getStringCellValue());
        }
    }
}
