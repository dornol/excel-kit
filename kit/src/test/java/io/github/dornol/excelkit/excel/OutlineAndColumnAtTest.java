package io.github.dornol.excelkit.excel;

import io.github.dornol.excelkit.csv.CsvReader;
import io.github.dornol.excelkit.shared.CellData;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
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
 * Tests for Feature 9 (column outline) and Feature 11 (columnAt index-based reading).
 */
class OutlineAndColumnAtTest {

    @TempDir
    Path tempDir;

    // ========================================================================
    // Feature 9: Column outline
    // ========================================================================
    @Test
    void outline_shouldGroupColumns() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        new ExcelWriter<String[]>()
                .column("Name", r -> r[0])
                .column("Detail1", r -> r[1]).outline(1)
                .column("Detail2", r -> r[2]).outline(1)
                .column("Summary", r -> r[3])
                .write(Stream.<String[]>of(new String[]{"A", "d1", "d2", "S"}))
                .consumeOutputStream(out);

        try (XSSFWorkbook wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
            Sheet sheet = wb.getSheetAt(0);
            // Detail1 (col 1) and Detail2 (col 2) should be grouped
            // POI doesn't expose outline directly, but we can check column is grouped
            // by verifying the outline level via CT (internal XML)
            var xssfSheet = wb.getSheetAt(0);
            assertTrue(xssfSheet.getColumnOutlineLevel(1) > 0, "Column 1 should have outline level > 0");
            assertTrue(xssfSheet.getColumnOutlineLevel(2) > 0, "Column 2 should have outline level > 0");
            assertEquals(0, xssfSheet.getColumnOutlineLevel(0), "Column 0 should have no outline");
            assertEquals(0, xssfSheet.getColumnOutlineLevel(3), "Column 3 should have no outline");
        }
    }

    @Test
    void outline_shouldWorkInExcelSheetWriter() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        try (ExcelWorkbook wb = new ExcelWorkbook()) {
            wb.<String[]>sheet("Test")
                    .column("A", r -> r[0])
                    .column("B", r -> r[1], c -> c.outline(1))
                    .column("C", r -> r[2], c -> c.outline(1))
                    .column("D", r -> r[3])
                    .write(Stream.<String[]>of(new String[]{"a", "b", "c", "d"}));
            wb.finish().consumeOutputStream(out);
        }

        try (XSSFWorkbook wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
            var sheet = wb.getSheetAt(0);
            assertTrue(sheet.getColumnOutlineLevel(1) > 0);
            assertTrue(sheet.getColumnOutlineLevel(2) > 0);
            assertEquals(0, sheet.getColumnOutlineLevel(0));
        }
    }

    @Test
    void outline_invalidLevel_shouldThrow() {
        assertThrows(IllegalArgumentException.class, () ->
                new ExcelWriter<String>().column("A", s -> s).outline(8));
        assertThrows(IllegalArgumentException.class, () ->
                new ExcelWriter<String>().column("A", s -> s).outline(-1));
    }

    // ========================================================================
    // Feature 11: columnAt — index-based reading
    // ========================================================================
    @Test
    void columnAt_shouldReadByExplicitIndex() throws IOException {
        Path file = tempDir.resolve("index-read.xlsx");
        createExcelFile(file, new String[]{"A", "B", "C", "D", "E"},
                new Object[][]{{"a0", "b0", "c0", "d0", "e0"}});

        List<TestRow> results = new ArrayList<>();
        try (InputStream is = Files.newInputStream(file)) {
            new ExcelReader<>(TestRow::new, null)
                    .columnAt(0, (r, cell) -> r.first = cell.asString())
                    .columnAt(2, (r, cell) -> r.second = cell.asString())
                    .columnAt(4, (r, cell) -> r.third = cell.asString())
                    .build(is)
                    .read(result -> results.add(result.data()));
        }

        assertEquals(1, results.size());
        assertEquals("a0", results.get(0).first);
        assertEquals("c0", results.get(0).second);
        assertEquals("e0", results.get(0).third);
    }

    @Test
    void columnAt_shouldWorkWithBuilderChain() throws IOException {
        Path file = tempDir.resolve("index-chain.xlsx");
        createExcelFile(file, new String[]{"Name", "Age", "City"},
                new Object[][]{{"Alice", 30, "Seoul"}});

        List<TestRow> results = new ArrayList<>();
        try (InputStream is = Files.newInputStream(file)) {
            new ExcelReader<>(TestRow::new, null)
                    .columnAt(2, (TestRow r, CellData cell) -> r.first = cell.asString())
                    .columnAt(0, (TestRow r, CellData cell) -> r.second = cell.asString())
                    .build(is)
                    .read(result -> results.add(result.data()));
        }

        assertEquals(1, results.size());
        assertEquals("Seoul", results.get(0).first);  // col 2
        assertEquals("Alice", results.get(0).second);  // col 0
    }

    @Test
    void columnAt_csv_shouldReadByExplicitIndex() {
        String csv = "A,B,C,D,E\na0,b0,c0,d0,e0\n";
        InputStream is = new ByteArrayInputStream(csv.getBytes(StandardCharsets.UTF_8));

        List<TestRow> results = new ArrayList<>();
        new CsvReader<>(TestRow::new, null)
                .columnAt(0, (r, cell) -> r.first = cell.asString())
                .columnAt(4, (r, cell) -> r.second = cell.asString())
                .build(is)
                .read(result -> results.add(result.data()));

        assertEquals(1, results.size());
        assertEquals("a0", results.get(0).first);
        assertEquals("e0", results.get(0).second);
    }

    @Test
    void columnAt_canMixWithNameBased() throws IOException {
        Path file = tempDir.resolve("mixed.xlsx");
        createExcelFile(file, new String[]{"Name", "Age", "City"},
                new Object[][]{{"Alice", 30, "Seoul"}});

        List<TestRow> results = new ArrayList<>();
        try (InputStream is = Files.newInputStream(file)) {
            new ExcelReader<>(TestRow::new, null)
                    .column("City", (TestRow r, CellData cell) -> r.first = cell.asString())
                    .columnAt(1, (r, cell) -> r.second = String.valueOf(cell.asInt()))
                    .build(is)
                    .read(result -> results.add(result.data()));
        }

        assertEquals(1, results.size());
        assertEquals("Seoul", results.get(0).first);   // by name
        assertEquals("30", results.get(0).second);       // by index
    }

    private void createExcelFile(Path filePath, String[] headers, Object[][] data) throws IOException {
        try (Workbook workbook = new XSSFWorkbook()) {
            Sheet sheet = workbook.createSheet("Test");
            Row headerRow = sheet.createRow(0);
            for (int i = 0; i < headers.length; i++) {
                headerRow.createCell(i).setCellValue(headers[i]);
            }
            for (int r = 0; r < data.length; r++) {
                Row row = sheet.createRow(r + 1);
                for (int c = 0; c < data[r].length; c++) {
                    Object val = data[r][c];
                    if (val instanceof String s) row.createCell(c).setCellValue(s);
                    else if (val instanceof Number n) row.createCell(c).setCellValue(n.doubleValue());
                }
            }
            try (FileOutputStream fos = new FileOutputStream(filePath.toFile())) {
                workbook.write(fos);
            }
        }
    }

    public static class TestRow {
        String first;
        String second;
        String third;
    }
}
