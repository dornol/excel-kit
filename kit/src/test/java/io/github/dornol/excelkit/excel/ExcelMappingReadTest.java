package io.github.dornol.excelkit.excel;

import io.github.dornol.excelkit.shared.ReadAbortException;
import io.github.dornol.excelkit.shared.ReadResult;
import jakarta.validation.Validation;
import jakarta.validation.Validator;
import jakarta.validation.constraints.Max;
import jakarta.validation.constraints.Min;
import jakarta.validation.constraints.NotBlank;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.hibernate.validator.messageinterpolation.ParameterMessageInterpolator;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.List;
import java.util.concurrent.atomic.AtomicLong;
import java.util.stream.Stream;

import static org.junit.jupiter.api.Assertions.*;

/**
 * Tests for mapping-based (immutable object) Excel reading.
 */
class ExcelMappingReadTest {

    @TempDir
    Path tempDir;

    record PersonRecord(String name, Integer age, String city) {}

    // --- Basic functionality ---

    @Test
    void mapping_shouldCreateImmutableObjectsByHeaderName() throws IOException {
        Path file = createExcelFile("mapping-basic.xlsx",
                new String[]{"Name", "Age", "City"},
                new Object[][]{
                        {"Alice", 30, "Seoul"},
                        {"Bob", 25, "Busan"}
                });

        List<PersonRecord> results = new ArrayList<>();
        try (InputStream is = Files.newInputStream(file)) {
            ExcelReader.<PersonRecord>mapping(row -> new PersonRecord(
                    row.get("Name").asString(),
                    row.get("Age").asInt(),
                    row.get("City").asString()
            )).build(is).read(r -> {
                assertTrue(r.success());
                results.add(r.data());
            });
        }

        assertEquals(2, results.size());
        assertEquals("Alice", results.get(0).name());
        assertEquals(30, results.get(0).age());
        assertEquals("Seoul", results.get(0).city());
        assertEquals("Bob", results.get(1).name());
        assertEquals(25, results.get(1).age());
        assertEquals("Busan", results.get(1).city());
    }

    @Test
    void mapping_shouldWorkWithDifferentColumnOrder() throws IOException {
        Path file = createExcelFile("mapping-reversed.xlsx",
                new String[]{"City", "Age", "Name"},
                new Object[][]{
                        {"Seoul", 30, "Alice"},
                        {"Busan", 25, "Bob"}
                });

        List<PersonRecord> results = new ArrayList<>();
        try (InputStream is = Files.newInputStream(file)) {
            ExcelReader.<PersonRecord>mapping(row -> new PersonRecord(
                    row.get("Name").asString(),
                    row.get("Age").asInt(),
                    row.get("City").asString()
            )).build(is).read(r -> results.add(r.data()));
        }

        assertEquals(2, results.size());
        assertEquals("Alice", results.get(0).name());
        assertEquals(30, results.get(0).age());
        assertEquals("Seoul", results.get(0).city());
    }

    @Test
    void mapping_shouldReadSubsetOfColumns() throws IOException {
        Path file = createExcelFile("mapping-subset.xlsx",
                new String[]{"Name", "Age", "City", "Email"},
                new Object[][]{
                        {"Alice", 30, "Seoul", "alice@test.com"}
                });

        List<PersonRecord> results = new ArrayList<>();
        try (InputStream is = Files.newInputStream(file)) {
            ExcelReader.<PersonRecord>mapping(row -> new PersonRecord(
                    row.get("Name").asString(),
                    null,
                    row.get("City").asString()
            )).build(is).read(r -> results.add(r.data()));
        }

        assertEquals(1, results.size());
        assertEquals("Alice", results.get(0).name());
        assertNull(results.get(0).age());
        assertEquals("Seoul", results.get(0).city());
    }

    // --- Error handling ---

    @Test
    void mapping_shouldThrowOnMissingHeader() throws IOException {
        Path file = createExcelFile("mapping-missing.xlsx",
                new String[]{"Name", "Age"},
                new Object[][]{{"Alice", 30}});

        List<ReadResult<PersonRecord>> results = new ArrayList<>();
        try (InputStream is = Files.newInputStream(file)) {
            ExcelReader.<PersonRecord>mapping(row -> new PersonRecord(
                    row.get("Name").asString(),
                    row.get("Age").asInt(),
                    row.get("NonExistent").asString()
            )).build(is).read(results::add);
        }

        assertEquals(1, results.size());
        assertFalse(results.get(0).success());
        assertNull(results.get(0).data());
        assertNotNull(results.get(0).messages());
        assertTrue(results.get(0).messages().get(0).contains("NonExistent"));
    }

    @Test
    void mapping_shouldHandleConversionError() throws IOException {
        Path file = createExcelFile("mapping-conv-error.xlsx",
                new String[]{"Name", "Age"},
                new Object[][]{{"Alice", "not-a-number"}});

        List<ReadResult<PersonRecord>> results = new ArrayList<>();
        try (InputStream is = Files.newInputStream(file)) {
            ExcelReader.<PersonRecord>mapping(row -> new PersonRecord(
                    row.get("Name").asString(),
                    row.get("Age").asInt(),  // will fail: "not-a-number"
                    null
            )).build(is).read(results::add);
        }

        assertEquals(1, results.size());
        assertFalse(results.get(0).success());
        assertNull(results.get(0).data());
    }

    @Test
    void mapping_shouldContinueAfterRowError() throws IOException {
        Path file = createExcelFile("mapping-continue.xlsx",
                new String[]{"Name", "Age"},
                new Object[][]{
                        {"Alice", 30},
                        {"Bob", "bad"},    // will fail
                        {"Charlie", 35}
                });

        List<ReadResult<PersonRecord>> results = new ArrayList<>();
        try (InputStream is = Files.newInputStream(file)) {
            ExcelReader.<PersonRecord>mapping(row -> new PersonRecord(
                    row.get("Name").asString(),
                    row.get("Age").asInt(),
                    null
            )).build(is).read(results::add);
        }

        assertEquals(3, results.size());
        assertTrue(results.get(0).success());
        assertEquals("Alice", results.get(0).data().name());
        assertFalse(results.get(1).success());  // Bob's row failed
        assertTrue(results.get(2).success());
        assertEquals("Charlie", results.get(2).data().name());
    }

    // --- Read modes ---

    @Test
    void mapping_shouldWorkWithReadStrict() throws IOException {
        Path file = createExcelFile("mapping-strict.xlsx",
                new String[]{"Name", "Age", "City"},
                new Object[][]{{"Alice", 30, "Seoul"}});

        List<PersonRecord> results = new ArrayList<>();
        try (InputStream is = Files.newInputStream(file)) {
            ExcelReader.<PersonRecord>mapping(row -> new PersonRecord(
                    row.get("Name").asString(),
                    row.get("Age").asInt(),
                    row.get("City").asString()
            )).build(is).readStrict(results::add);
        }

        assertEquals(1, results.size());
        assertEquals("Alice", results.get(0).name());
    }

    @Test
    void mapping_readStrict_shouldThrowOnError() throws IOException {
        Path file = createExcelFile("mapping-strict-fail.xlsx",
                new String[]{"Name", "Age"},
                new Object[][]{
                        {"Alice", 30},
                        {"Bob", "bad"}
                });

        try (InputStream is = Files.newInputStream(file)) {
            var handler = ExcelReader.<PersonRecord>mapping(row -> new PersonRecord(
                    row.get("Name").asString(),
                    row.get("Age").asInt(),
                    null
            )).build(is);

            assertThrows(ReadAbortException.class, () -> handler.readStrict(r -> {}));
        }
    }

    @Test
    void mapping_shouldWorkWithReadAsStream() throws IOException {
        Path file = createExcelFile("mapping-stream.xlsx",
                new String[]{"Name", "Age", "City"},
                new Object[][]{
                        {"Alice", 30, "Seoul"},
                        {"Bob", 25, "Busan"}
                });

        List<PersonRecord> results;
        try (InputStream is = Files.newInputStream(file)) {
            try (var stream = ExcelReader.<PersonRecord>mapping(row -> new PersonRecord(
                    row.get("Name").asString(),
                    row.get("Age").asInt(),
                    row.get("City").asString()
            )).build(is).readAsStream()) {
                results = stream
                        .filter(ReadResult::success)
                        .map(ReadResult::data)
                        .toList();
            }
        }

        assertEquals(2, results.size());
        assertEquals("Alice", results.get(0).name());
        assertEquals("Bob", results.get(1).name());
    }

    // --- Configuration options ---

    @Test
    void mapping_shouldWorkWithSheetIndex() throws IOException {
        Path file = tempDir.resolve("mapping-multi-sheet.xlsx");
        try (Workbook workbook = new XSSFWorkbook()) {
            Sheet sheet0 = workbook.createSheet("Sheet0");
            Row h0 = sheet0.createRow(0);
            h0.createCell(0).setCellValue("X");
            sheet0.createRow(1).createCell(0).setCellValue("ignore");

            Sheet sheet1 = workbook.createSheet("Sheet1");
            Row h1 = sheet1.createRow(0);
            h1.createCell(0).setCellValue("Name");
            h1.createCell(1).setCellValue("Age");
            Row d1 = sheet1.createRow(1);
            d1.createCell(0).setCellValue("Charlie");
            d1.createCell(1).setCellValue(35);

            try (FileOutputStream fos = new FileOutputStream(file.toFile())) {
                workbook.write(fos);
            }
        }

        List<PersonRecord> results = new ArrayList<>();
        try (InputStream is = Files.newInputStream(file)) {
            ExcelReader.<PersonRecord>mapping(row -> new PersonRecord(
                    row.get("Name").asString(),
                    row.get("Age").asInt(),
                    null
            )).sheetIndex(1)
              .build(is)
              .read(r -> results.add(r.data()));
        }

        assertEquals(1, results.size());
        assertEquals("Charlie", results.get(0).name());
        assertEquals(35, results.get(0).age());
    }

    @Test
    void mapping_shouldWorkWithHeaderRowIndex() throws IOException {
        Path file = tempDir.resolve("mapping-header-offset.xlsx");
        try (Workbook workbook = new XSSFWorkbook()) {
            Sheet sheet = workbook.createSheet("Test");
            // Row 0: metadata
            sheet.createRow(0).createCell(0).setCellValue("METADATA ROW");
            // Row 1: also metadata
            sheet.createRow(1).createCell(0).setCellValue("SKIP THIS");
            // Row 2: actual header
            Row headerRow = sheet.createRow(2);
            headerRow.createCell(0).setCellValue("Name");
            headerRow.createCell(1).setCellValue("Age");
            // Row 3: data
            Row dataRow = sheet.createRow(3);
            dataRow.createCell(0).setCellValue("Alice");
            dataRow.createCell(1).setCellValue(30);

            try (FileOutputStream fos = new FileOutputStream(file.toFile())) {
                workbook.write(fos);
            }
        }

        List<PersonRecord> results = new ArrayList<>();
        try (InputStream is = Files.newInputStream(file)) {
            ExcelReader.<PersonRecord>mapping(row -> new PersonRecord(
                    row.get("Name").asString(),
                    row.get("Age").asInt(),
                    null
            )).headerRowIndex(2)
              .build(is)
              .read(r -> results.add(r.data()));
        }

        assertEquals(1, results.size());
        assertEquals("Alice", results.get(0).name());
        assertEquals(30, results.get(0).age());
    }

    @Test
    void mapping_shouldWorkWithProgressCallback() throws IOException {
        Path file = createExcelFile("mapping-progress.xlsx",
                new String[]{"Name", "Age"},
                new Object[][]{
                        {"A", 1}, {"B", 2}, {"C", 3}, {"D", 4}, {"E", 5}
                });

        AtomicLong lastProgress = new AtomicLong(0);
        try (InputStream is = Files.newInputStream(file)) {
            ExcelReader.<PersonRecord>mapping(row -> new PersonRecord(
                    row.get("Name").asString(),
                    row.get("Age").asInt(),
                    null
            )).onProgress(2, (count, cursor) -> lastProgress.set(count))
              .build(is)
              .read(r -> {});
        }

        // 5 rows, interval 2: fires at 2 and 4
        assertEquals(4, lastProgress.get());
    }

    // --- RowData access patterns ---

    @Test
    void mapping_shouldSupportRowDataIndexAccess() throws IOException {
        Path file = createExcelFile("mapping-index.xlsx",
                new String[]{"Name", "Age", "City"},
                new Object[][]{{"Alice", 30, "Seoul"}});

        List<PersonRecord> results = new ArrayList<>();
        try (InputStream is = Files.newInputStream(file)) {
            ExcelReader.<PersonRecord>mapping(row -> new PersonRecord(
                    row.get(0).asString(),
                    row.get(1).asInt(),
                    row.get(2).asString()
            )).build(is).read(r -> results.add(r.data()));
        }

        assertEquals(1, results.size());
        assertEquals("Alice", results.get(0).name());
        assertEquals(30, results.get(0).age());
        assertEquals("Seoul", results.get(0).city());
    }

    @Test
    void mapping_rowDataHasShouldWork() throws IOException {
        Path file = createExcelFile("mapping-has.xlsx",
                new String[]{"Name", "Age"},
                new Object[][]{{"Alice", 30}});

        try (InputStream is = Files.newInputStream(file)) {
            ExcelReader.<PersonRecord>mapping(row -> {
                assertTrue(row.has("Name"));
                assertTrue(row.has("Age"));
                assertFalse(row.has("NonExistent"));
                assertEquals(2, row.headerNames().size());
                assertEquals(List.of("Name", "Age"), row.headerNames());
                return new PersonRecord(row.get("Name").asString(), row.get("Age").asInt(), null);
            }).build(is).read(r -> {});
        }
    }

    @Test
    void mapping_shouldHandleMissingCellsGracefully() throws IOException {
        Path file = tempDir.resolve("mapping-sparse.xlsx");
        try (Workbook workbook = new XSSFWorkbook()) {
            Sheet sheet = workbook.createSheet("Test");
            Row headerRow = sheet.createRow(0);
            headerRow.createCell(0).setCellValue("Name");
            headerRow.createCell(1).setCellValue("Age");
            headerRow.createCell(2).setCellValue("City");

            // Only set Name, skip Age and City
            Row dataRow = sheet.createRow(1);
            dataRow.createCell(0).setCellValue("Alice");

            try (FileOutputStream fos = new FileOutputStream(file.toFile())) {
                workbook.write(fos);
            }
        }

        List<PersonRecord> results = new ArrayList<>();
        try (InputStream is = Files.newInputStream(file)) {
            ExcelReader.<PersonRecord>mapping(row -> new PersonRecord(
                    row.get("Name").asString(),
                    row.get("Age").isEmpty() ? null : row.get("Age").asInt(),
                    row.get("City").asString()
            )).build(is).read(r -> results.add(r.data()));
        }

        assertEquals(1, results.size());
        assertEquals("Alice", results.get(0).name());
        assertNull(results.get(0).age());
        assertEquals("", results.get(0).city());
    }

    // --- Bean Validation ---

    @Test
    void mapping_shouldWorkWithBeanValidation() throws IOException {
        Path file = createExcelFile("mapping-validation.xlsx",
                new String[]{"Name", "Age"},
                new Object[][]{
                        {"Alice", 30},     // valid
                        {"", 25},          // invalid: blank name
                        {"Charlie", 150}   // invalid: age > 100
                });

        Validator validator = Validation.byDefaultProvider()
                .configure()
                .messageInterpolator(new ParameterMessageInterpolator())
                .buildValidatorFactory()
                .getValidator();

        List<ReadResult<ValidatedPerson>> results = new ArrayList<>();
        try (InputStream is = Files.newInputStream(file)) {
            ExcelReader.<ValidatedPerson>mapping(row -> {
                ValidatedPerson p = new ValidatedPerson();
                p.name = row.get("Name").asString();
                p.age = row.get("Age").asInt();
                return p;
            }, validator).build(is).read(results::add);
        }

        assertEquals(3, results.size());
        assertTrue(results.get(0).success());
        assertEquals("Alice", results.get(0).data().name);
        assertFalse(results.get(1).success());  // blank name
        assertFalse(results.get(2).success());  // age > 100
    }

    // --- Round-trip: write then read back with mapping ---

    @Test
    void roundTrip_writeWithExcelWriter_readWithMapping() throws IOException {
        // Write
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        new ExcelWriter<PersonRecord>()
                .addColumn("Name", PersonRecord::name)
                .addColumn("Age", p -> p.age())
                .addColumn("City", PersonRecord::city)
                .write(Stream.of(
                        new PersonRecord("Alice", 30, "Seoul"),
                        new PersonRecord("Bob", 25, "Busan")))
                .consumeOutputStream(out);

        // Read back with mapping
        List<PersonRecord> results = new ArrayList<>();
        try (InputStream is = new ByteArrayInputStream(out.toByteArray())) {
            ExcelReader.<PersonRecord>mapping(row -> new PersonRecord(
                    row.get("Name").asString(),
                    row.get("Age").asInt(),
                    row.get("City").asString()
            )).build(is).read(r -> {
                assertTrue(r.success());
                results.add(r.data());
            });
        }

        assertEquals(2, results.size());
        assertEquals("Alice", results.get(0).name());
        assertEquals(30, results.get(0).age());
        assertEquals("Seoul", results.get(0).city());
        assertEquals("Bob", results.get(1).name());
    }

    // --- Edge cases ---

    @Test
    void mapping_shouldHandleEmptyFile() throws IOException {
        Path file = tempDir.resolve("mapping-empty.xlsx");
        try (Workbook workbook = new XSSFWorkbook()) {
            Sheet sheet = workbook.createSheet("Test");
            Row headerRow = sheet.createRow(0);
            headerRow.createCell(0).setCellValue("Name");
            // No data rows

            try (FileOutputStream fos = new FileOutputStream(file.toFile())) {
                workbook.write(fos);
            }
        }

        List<ReadResult<PersonRecord>> results = new ArrayList<>();
        try (InputStream is = Files.newInputStream(file)) {
            ExcelReader.<PersonRecord>mapping(row -> new PersonRecord(
                    row.get("Name").asString(), null, null
            )).build(is).read(results::add);
        }

        assertTrue(results.isEmpty());
    }

    @Test
    void mapping_shouldHandleManyColumns() throws IOException {
        int colCount = 50;
        String[] headers = new String[colCount];
        Object[] values = new Object[colCount];
        for (int i = 0; i < colCount; i++) {
            headers[i] = "Col" + i;
            values[i] = "val" + i;
        }
        Path file = createExcelFile("mapping-many-cols.xlsx", headers, new Object[][]{values});

        try (InputStream is = Files.newInputStream(file)) {
            ExcelReader.<String>mapping(row -> {
                assertEquals(colCount, row.headerNames().size());
                assertEquals("val0", row.get("Col0").asString());
                assertEquals("val49", row.get("Col49").asString());
                return row.get("Col0").asString();
            }).build(is).read(r -> {
                assertTrue(r.success());
                assertEquals("val0", r.data());
            });
        }
    }

    // --- Helper ---

    public static class ValidatedPerson {
        @NotBlank
        String name;
        @Min(1) @Max(100)
        int age;
    }

    private Path createExcelFile(String filename, String[] headers, Object[][] data) throws IOException {
        Path file = tempDir.resolve(filename);
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
            try (FileOutputStream fos = new FileOutputStream(file.toFile())) {
                workbook.write(fos);
            }
        }
        return file;
    }
}
