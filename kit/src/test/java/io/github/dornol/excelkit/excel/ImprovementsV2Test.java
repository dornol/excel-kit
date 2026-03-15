package io.github.dornol.excelkit.excel;

import io.github.dornol.excelkit.csv.CsvWriter;
import io.github.dornol.excelkit.csv.CsvWriteException;
import io.github.dornol.excelkit.shared.CellData;
import io.github.dornol.excelkit.shared.ReadAbortException;
import io.github.dornol.excelkit.shared.ReadResult;
import jakarta.validation.Validation;
import jakarta.validation.Validator;
import jakarta.validation.constraints.NotBlank;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.hibernate.validator.messageinterpolation.ParameterMessageInterpolator;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import java.io.*;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.List;
import java.util.stream.Stream;

import static org.junit.jupiter.api.Assertions.*;

/**
 * Tests for v0.5.1 improvements:
 * - readStrict row numbers
 * - CSV duplicate column validation
 * - autoWidthSampleRows
 * - CellData.asEnum()
 * - Excel read progress callback
 */
class ImprovementsV2Test {

    @TempDir
    Path tempDir;

    // ========================================================================
    // readStrict - row number in error message
    // ========================================================================
    @Test
    void readStrict_errorShouldIncludeRowNumber() throws IOException {
        Path file = tempDir.resolve("invalid.xlsx");
        try (Workbook wb = new XSSFWorkbook()) {
            Sheet sheet = wb.createSheet("Test");
            sheet.createRow(0).createCell(0).setCellValue("Name");
            sheet.createRow(1).createCell(0).setCellValue("Alice");
            sheet.createRow(2).createCell(0).setCellValue(""); // invalid: blank
            try (FileOutputStream fos = new FileOutputStream(file.toFile())) {
                wb.write(fos);
            }
        }

        Validator validator = Validation.byDefaultProvider().configure()
                .messageInterpolator(new ParameterMessageInterpolator())
                .buildValidatorFactory().getValidator();

        try (InputStream is = Files.newInputStream(file)) {
            ExcelReadHandler<ValidatedPerson> handler = new ExcelReader<>(ValidatedPerson::new, validator)
                    .column((p, cell) -> p.name = cell.asString())
                    .build(is);

            ReadAbortException ex = assertThrows(ReadAbortException.class,
                    () -> handler.readStrict(p -> {}));
            assertTrue(ex.getMessage().contains("Row 2"), "Error should contain 'Row 2' but was: " + ex.getMessage());
        }
    }

    @Test
    void readStrict_successShouldNotThrow() throws IOException {
        Path file = tempDir.resolve("valid.xlsx");
        try (Workbook wb = new XSSFWorkbook()) {
            Sheet sheet = wb.createSheet("Test");
            sheet.createRow(0).createCell(0).setCellValue("Name");
            sheet.createRow(1).createCell(0).setCellValue("Alice");
            sheet.createRow(2).createCell(0).setCellValue("Bob");
            try (FileOutputStream fos = new FileOutputStream(file.toFile())) {
                wb.write(fos);
            }
        }

        List<String> names = new ArrayList<>();
        try (InputStream is = Files.newInputStream(file)) {
            new ExcelReader<>(ValidatedPerson::new, null)
                    .column((p, cell) -> p.name = cell.asString())
                    .build(is)
                    .readStrict(p -> names.add(p.name));
        }

        assertEquals(List.of("Alice", "Bob"), names);
    }

    // ========================================================================
    // CSV duplicate column validation
    // ========================================================================
    @Test
    void csvDuplicateColumn_shouldThrow() {
        var writer = new CsvWriter<String>()
                .column("Name", s -> s)
                .column("Name", s -> s);
        assertThrows(CsvWriteException.class, () -> writer.write(Stream.of("test")));
    }

    @Test
    void csvUniqueColumns_shouldNotThrow() {
        assertDoesNotThrow(() ->
                new CsvWriter<String>()
                        .column("Name", s -> s)
                        .column("Age", s -> s)
                        .write(Stream.of("test")));
    }

    // ========================================================================
    // autoWidthSampleRows
    // ========================================================================
    @Test
    void autoWidthSampleRows_shouldBeConfigurable() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        new ExcelWriter<String>()
                .autoWidthSampleRows(5)
                .column("Name", s -> s)
                .write(Stream.of("short", "a very long column value that should affect width"))
                .consumeOutputStream(out);

        assertTrue(out.toByteArray().length > 0);
    }

    @Test
    void autoWidthSampleRows_zero_shouldDisable() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        new ExcelWriter<String>()
                .autoWidthSampleRows(0)
                .column("Name", s -> s)
                .write(Stream.of("test"))
                .consumeOutputStream(out);

        assertTrue(out.toByteArray().length > 0);
    }

    @Test
    void autoWidthSampleRows_negative_shouldThrow() {
        assertThrows(IllegalArgumentException.class, () ->
                new ExcelWriter<String>().autoWidthSampleRows(-1));
    }

    @Test
    void autoWidthSampleRows_inExcelSheetWriter() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        try (ExcelWorkbook wb = new ExcelWorkbook()) {
            wb.<String>sheet("Test")
                    .autoWidthSampleRows(10)
                    .column("Name", s -> s)
                    .write(Stream.of("test"));
            wb.finish().consumeOutputStream(out);
        }
        assertTrue(out.toByteArray().length > 0);
    }

    // ========================================================================
    // CellData.asEnum()
    // ========================================================================
    @Test
    void asEnum_shouldParseByName() {
        CellData cell = new CellData(0, "ACTIVE");
        assertEquals(Status.ACTIVE, cell.asEnum(Status.class));
    }

    @Test
    void asEnum_shouldBeCaseInsensitive() {
        assertEquals(Status.ACTIVE, new CellData(0, "active").asEnum(Status.class));
        assertEquals(Status.INACTIVE, new CellData(0, "Inactive").asEnum(Status.class));
        assertEquals(Status.PENDING, new CellData(0, "PENDING").asEnum(Status.class));
    }

    @Test
    void asEnum_blank_shouldReturnNull() {
        assertNull(new CellData(0, "").asEnum(Status.class));
        assertNull(new CellData(0, "  ").asEnum(Status.class));
    }

    @Test
    void asEnum_invalidValue_shouldThrow() {
        assertThrows(IllegalArgumentException.class, () ->
                new CellData(0, "UNKNOWN").asEnum(Status.class));
    }

    @Test
    void asEnum_withTrimming() {
        assertEquals(Status.ACTIVE, new CellData(0, " ACTIVE ").asEnum(Status.class));
    }

    // ========================================================================
    // Excel read progress callback
    // ========================================================================
    @Test
    void readProgress_shouldFireAtCorrectIntervals() throws IOException {
        Path file = tempDir.resolve("progress.xlsx");
        createExcelFileWithRows(file, 10);

        List<Long> progressCounts = new ArrayList<>();
        try (InputStream is = Files.newInputStream(file)) {
            new ExcelReader<>(SimplePerson::new, null)
                    .column((p, cell) -> p.name = cell.asString())
                    .column((p, cell) -> p.age = cell.asInt())
                    .onProgress(3, (count, cursor) -> progressCounts.add(count))
                    .build(is)
                    .read(r -> {});
        }

        assertEquals(List.of(3L, 6L, 9L), progressCounts);
    }

    @Test
    void readProgress_shouldNotFireWhenIntervalNotReached() throws IOException {
        Path file = tempDir.resolve("progress-small.xlsx");
        createExcelFileWithRows(file, 2);

        List<Long> progressCounts = new ArrayList<>();
        try (InputStream is = Files.newInputStream(file)) {
            new ExcelReader<>(SimplePerson::new, null)
                    .column((p, cell) -> p.name = cell.asString())
                    .column((p, cell) -> p.age = cell.asInt())
                    .onProgress(100, (count, cursor) -> progressCounts.add(count))
                    .build(is)
                    .read(r -> {});
        }

        assertTrue(progressCounts.isEmpty());
    }

    @Test
    void readProgress_invalidInterval_shouldThrow() {
        assertThrows(IllegalArgumentException.class, () ->
                new ExcelReader<>(SimplePerson::new, null)
                        .onProgress(0, (c, cur) -> {}));
    }

    @Test
    void readProgress_viaBuilderChain() throws IOException {
        Path file = tempDir.resolve("progress-chain.xlsx");
        createExcelFileWithRows(file, 6);

        List<Long> counts = new ArrayList<>();
        try (InputStream is = Files.newInputStream(file)) {
            new ExcelReader<>(SimplePerson::new, null)
                    .column((SimplePerson p, CellData cell) -> p.name = cell.asString())
                    .column((SimplePerson p, CellData cell) -> p.age = cell.asInt())
                    .onProgress(2, (count, cursor) -> counts.add(count))
                    .build(is)
                    .read(r -> {});
        }

        assertEquals(List.of(2L, 4L, 6L), counts);
    }

    // ========================================================================
    // Helpers
    // ========================================================================
    private void createExcelFileWithRows(Path filePath, int rowCount) throws IOException {
        try (Workbook workbook = new XSSFWorkbook()) {
            Sheet sheet = workbook.createSheet("Test");
            Row header = sheet.createRow(0);
            header.createCell(0).setCellValue("Name");
            header.createCell(1).setCellValue("Age");
            for (int i = 0; i < rowCount; i++) {
                Row row = sheet.createRow(i + 1);
                row.createCell(0).setCellValue("Person" + i);
                row.createCell(1).setCellValue(20 + i);
            }
            try (FileOutputStream fos = new FileOutputStream(filePath.toFile())) {
                workbook.write(fos);
            }
        }
    }

    public enum Status { ACTIVE, INACTIVE, PENDING }

    public static class ValidatedPerson {
        @NotBlank(message = "Name must not be blank")
        String name;
    }

    public static class SimplePerson {
        String name;
        Integer age;
    }
}
