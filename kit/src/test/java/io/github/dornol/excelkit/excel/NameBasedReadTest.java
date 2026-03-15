package io.github.dornol.excelkit.excel;

import io.github.dornol.excelkit.shared.CellData;
import io.github.dornol.excelkit.shared.ReadResult;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.List;

import static org.junit.jupiter.api.Assertions.*;

/**
 * Tests for header name-based column mapping in ExcelReader.
 */
class NameBasedReadTest {

    @TempDir
    Path tempDir;

    @Test
    void readByName_shouldMatchByHeaderName() throws IOException {
        Path file = tempDir.resolve("name-based.xlsx");
        createExcelFile(file, new String[]{"Name", "Age", "City"},
                new Object[][]{
                        {"Alice", 30, "Seoul"},
                        {"Bob", 25, "Busan"}
                });

        List<TestPerson> results = new ArrayList<>();
        try (InputStream is = Files.newInputStream(file)) {
            new ExcelReader<>(TestPerson::new, null)
                    .column("Name", (p, cell) -> p.name = cell.asString())
                    .column("Age", (p, cell) -> p.age = cell.asInt())
                    .column("City", (p, cell) -> p.city = cell.asString())
                    .build(is)
                    .read(r -> results.add(r.data()));
        }

        assertEquals(2, results.size());
        assertEquals("Alice", results.get(0).name);
        assertEquals(30, results.get(0).age);
        assertEquals("Seoul", results.get(0).city);
    }

    @Test
    void readByName_shouldWorkWithDifferentColumnOrder() throws IOException {
        // Excel has columns in order: City, Age, Name (reversed)
        Path file = tempDir.resolve("reversed.xlsx");
        createExcelFile(file, new String[]{"City", "Age", "Name"},
                new Object[][]{
                        {"Seoul", 30, "Alice"},
                        {"Busan", 25, "Bob"}
                });

        List<TestPerson> results = new ArrayList<>();
        try (InputStream is = Files.newInputStream(file)) {
            // Reader defines columns in: Name, Age, City order
            new ExcelReader<>(TestPerson::new, null)
                    .column("Name", (p, cell) -> p.name = cell.asString())
                    .column("Age", (p, cell) -> p.age = cell.asInt())
                    .column("City", (p, cell) -> p.city = cell.asString())
                    .build(is)
                    .read(r -> results.add(r.data()));
        }

        assertEquals(2, results.size());
        assertEquals("Alice", results.get(0).name);
        assertEquals(30, results.get(0).age);
        assertEquals("Seoul", results.get(0).city);
        assertEquals("Bob", results.get(1).name);
        assertEquals(25, results.get(1).age);
        assertEquals("Busan", results.get(1).city);
    }

    @Test
    void readByName_shouldReadSubsetOfColumns() throws IOException {
        Path file = tempDir.resolve("subset.xlsx");
        createExcelFile(file, new String[]{"Name", "Age", "City", "Email"},
                new Object[][]{
                        {"Alice", 30, "Seoul", "alice@test.com"},
                });

        List<TestPerson> results = new ArrayList<>();
        try (InputStream is = Files.newInputStream(file)) {
            // Only read Name and City, skip Age and Email
            new ExcelReader<>(TestPerson::new, null)
                    .column("Name", (p, cell) -> p.name = cell.asString())
                    .column("City", (p, cell) -> p.city = cell.asString())
                    .build(is)
                    .read(r -> results.add(r.data()));
        }

        assertEquals(1, results.size());
        assertEquals("Alice", results.get(0).name);
        assertEquals("Seoul", results.get(0).city);
        assertNull(results.get(0).age);
    }

    @Test
    void readByName_shouldThrowWhenHeaderNotFound() throws IOException {
        Path file = tempDir.resolve("missing-header.xlsx");
        createExcelFile(file, new String[]{"Name", "Age"},
                new Object[][]{{"Alice", 30}});

        try (InputStream is = Files.newInputStream(file)) {
            ExcelReadHandler<TestPerson> handler = new ExcelReader<>(TestPerson::new, null)
                    .column("Name", (p, cell) -> p.name = cell.asString())
                    .column("NonExistent", (p, cell) -> p.city = cell.asString())
                    .build(is);

            assertThrows(ExcelReadException.class, () -> handler.read(r -> {}));
        }
    }

    @Test
    void readByName_shouldWorkWithAddColumnMethod() throws IOException {
        Path file = tempDir.resolve("add-column.xlsx");
        createExcelFile(file, new String[]{"City", "Name"},
                new Object[][]{{"Seoul", "Alice"}});

        List<TestPerson> results = new ArrayList<>();
        try (InputStream is = Files.newInputStream(file)) {
            new ExcelReader<>(TestPerson::new, null)
                    .addColumn("Name", (p, cell) -> p.name = cell.asString())
                    .addColumn("City", (p, cell) -> p.city = cell.asString())
                    .build(is)
                    .read(r -> results.add(r.data()));
        }

        assertEquals(1, results.size());
        assertEquals("Alice", results.get(0).name);
        assertEquals("Seoul", results.get(0).city);
    }

    @Test
    void readByName_shouldWorkWithBuilderChaining() throws IOException {
        Path file = tempDir.resolve("builder-chain.xlsx");
        createExcelFile(file, new String[]{"Age", "Name"},
                new Object[][]{{30, "Alice"}});

        List<TestPerson> results = new ArrayList<>();
        try (InputStream is = Files.newInputStream(file)) {
            new ExcelReader<>(TestPerson::new, null)
                    .column("Name", (TestPerson p, CellData cell) -> p.name = cell.asString())
                    .column("Age", (TestPerson p, CellData cell) -> p.age = cell.asInt())
                    .build(is)
                    .read(r -> results.add(r.data()));
        }

        assertEquals(1, results.size());
        assertEquals("Alice", results.get(0).name);
        assertEquals(30, results.get(0).age);
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
                    if (val instanceof String s) {
                        row.createCell(c).setCellValue(s);
                    } else if (val instanceof Number n) {
                        row.createCell(c).setCellValue(n.doubleValue());
                    }
                }
            }
            try (FileOutputStream fos = new FileOutputStream(filePath.toFile())) {
                workbook.write(fos);
            }
        }
    }

    public static class TestPerson {
        String name;
        Integer age;
        String city;
    }
}
