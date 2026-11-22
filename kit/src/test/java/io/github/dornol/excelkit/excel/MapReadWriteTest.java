package io.github.dornol.excelkit.excel;

import io.github.dornol.excelkit.csv.CsvWriter;
import io.github.dornol.excelkit.excel.read.ExcelReader;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.nio.charset.StandardCharsets;
import java.util.*;
import java.util.stream.Stream;

import static org.junit.jupiter.api.Assertions.*;

class MapReadWriteTest {
    @Test
    void excelMapWriter_withConfigurers() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        @SuppressWarnings("unchecked")
        var writer = ExcelWriter.forMap(
                new String[]{"Name", "Price"},
                c -> c.bold(true),
                c -> c.type(ExcelDataType.INTEGER)
        );
        writer.write(Stream.of(
                Map.of("Name", "Item", "Price", 100)
        )).writeTo(out);

        try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
            assertEquals("Item", wb.getSheetAt(0).getRow(1).getCell(0).getStringCellValue());
        }
    }

    @Test
    void excelMapWriter_emptyStream() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        ExcelWriter.forMap("A", "B")
                .write(Stream.empty())
                .writeTo(out);

        try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
            assertEquals("A", wb.getSheetAt(0).getRow(0).getCell(0).getStringCellValue());
            assertNull(wb.getSheetAt(0).getRow(1)); // no data rows
        }
    }

    @Test
    void excelMapWriter_nullValuesInMap() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        Map<String, Object> row = new HashMap<>();
        row.put("Name", null);
        row.put("Age", null);

        ExcelWriter.forMap("Name", "Age")
                .write(Stream.of(row))
                .writeTo(out);

        try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
            // null values should result in empty cells
            assertEquals("", wb.getSheetAt(0).getRow(1).getCell(0).getStringCellValue());
        }
    }

    @Test
    void excelMapWriter_extraKeysIgnored() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        ExcelWriter.forMap("Name")
                .write(Stream.of(Map.of("Name", "Alice", "Extra", "ignored")))
                .writeTo(out);

        try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
            assertEquals(1, wb.getSheetAt(0).getRow(0).getLastCellNum()); // only 1 column
        }
    }

    @Test
    void excelMapReader_emptyDataRows() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        ExcelWriter.forMap("Name", "Age")
                .write(Stream.empty())
                .writeTo(out);

        List<Map<String, String>> results = new ArrayList<>();
        ExcelReader.forMap()
                .read(new ByteArrayInputStream(out.toByteArray()), r -> results.add(r.data()));

        assertTrue(results.isEmpty());
    }

    @Test
    void excelMapReader_sparseData() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        Map<String, Object> fullRow = new LinkedHashMap<>();
        fullRow.put("A", "val1");
        fullRow.put("B", "val2");
        fullRow.put("C", "val3");

        Map<String, Object> sparseRow = new LinkedHashMap<>();
        sparseRow.put("A", "only-a");
        sparseRow.put("B", null);
        sparseRow.put("C", null);

        ExcelWriter.forMap("A", "B", "C")
                .write(Stream.of(fullRow, sparseRow))
                .writeTo(out);

        List<Map<String, String>> results = new ArrayList<>();
        ExcelReader.forMap()
                .read(new ByteArrayInputStream(out.toByteArray()), r -> results.add(r.data()));

        assertEquals(2, results.size());
        assertEquals("val1", results.get(0).get("A"));
        assertEquals("val2", results.get(0).get("B"));
        assertEquals("only-a", results.get(1).get("A"));
    }

    @Test
    void excelMapReader_headerRowIndex() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        ExcelWriter.<String>create()
                .beforeHeader(ctx -> {
                    ctx.getSheet().createRow(0).createCell(0).setCellValue("Title");
                    return 1;
                })
                .column("Name", s -> s)
                .write(Stream.of("Alice"))
                .writeTo(out);

        List<Map<String, String>> results = new ArrayList<>();
        ExcelReader.forMap()
                .headerRowIndex(1)
                .read(new ByteArrayInputStream(out.toByteArray()), r -> results.add(r.data()));

        assertEquals(1, results.size());
        assertEquals("Alice", results.get(0).get("Name"));
    }

    @Test
    void csvMapWriter_specialCharacters() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        CsvWriter.forMap("Name", "Desc")
                .write(Stream.of(Map.of("Name", "Alice, Jr.", "Desc", "She said \"hi\"")))
                .writeTo(out);

        String csv = out.toString(StandardCharsets.UTF_8);
        assertTrue(csv.contains("Alice, Jr.") || csv.contains("\"Alice, Jr.\""));
    }

    @Test
    void csvMapWriter_withDelimiter() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        CsvWriter.forMap("A", "B")
                .delimiter('\t')
                .write(Stream.of(Map.of("A", "1", "B", "2")))
                .writeTo(out);

        String csv = out.toString(StandardCharsets.UTF_8);
        assertTrue(csv.contains("A\tB"));
    }

    @Test
    void csvMapWriter_nullValues() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        Map<String, Object> row = new HashMap<>();
        row.put("Name", "Alice");
        row.put("Value", null);

        CsvWriter.forMap("Name", "Value")
                .write(Stream.of(row))
                .writeTo(out);

        assertFalse(out.toByteArray().length == 0);
    }
}
