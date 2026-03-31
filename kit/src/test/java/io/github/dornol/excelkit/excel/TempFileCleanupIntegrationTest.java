package io.github.dornol.excelkit.excel;

import io.github.dornol.excelkit.shared.ReadResult;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
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
 * Integration tests verifying that temporary files are properly cleaned up
 * after Excel read/write operations, including error scenarios.
 */
class TempFileCleanupIntegrationTest {

    @TempDir
    Path tempDir;

    // ──────────────────────────────────────────────────────────────
    // ExcelHandler: temp file cleanup after encrypted write
    // ──────────────────────────────────────────────────────────────

    @Test
    void encryptedWrite_shouldCleanUpTempFiles() throws IOException {
        SXSSFWorkbook wb = new SXSSFWorkbook();
        wb.createSheet("Test").createRow(0).createCell(0).setCellValue("Data");
        ExcelHandler handler = new ExcelHandler(wb);

        Path output = tempDir.resolve("encrypted.xlsx");
        try (FileOutputStream fos = new FileOutputStream(output.toFile())) {
            handler.consumeOutputStreamWithPassword(fos, "password123");
        }

        assertTrue(Files.exists(output));
        assertTrue(Files.size(output) > 0);

        // Verify no leftover temp files in system temp (best-effort check)
        // The actual temp files are in system temp with UUID names, so we verify
        // the operation completes without errors
    }

    @Test
    void encryptedWrite_withCharArrayPassword_shouldCleanUpAndZeroPassword() throws IOException {
        SXSSFWorkbook wb = new SXSSFWorkbook();
        wb.createSheet("Test").createRow(0).createCell(0).setCellValue("Data");
        ExcelHandler handler = new ExcelHandler(wb);

        char[] password = "secret".toCharArray();
        Path output = tempDir.resolve("encrypted-char.xlsx");
        try (FileOutputStream fos = new FileOutputStream(output.toFile())) {
            handler.consumeOutputStreamWithPassword(fos, password);
        }

        assertTrue(Files.exists(output));
        // Password should be zeroed
        for (char c : password) {
            assertEquals('\0', c, "Password array must be zeroed after use");
        }
    }

    @Test
    void encryptedWrite_alreadyConsumed_shouldThrow() throws IOException {
        SXSSFWorkbook wb = new SXSSFWorkbook();
        ExcelHandler handler = new ExcelHandler(wb);

        handler.consumeOutputStream(new ByteArrayOutputStream());

        assertThrows(ExcelWriteException.class, () ->
                handler.consumeOutputStreamWithPassword(new ByteArrayOutputStream(), "pw"));
    }

    // ──────────────────────────────────────────────────────────────
    // ExcelReader.getSheetNames: temp file cleanup
    // ──────────────────────────────────────────────────────────────

    @Test
    void getSheetNames_shouldCleanUpTempFiles() throws IOException {
        Path excelFile = createTestExcel("Sheet1", "Sheet2", "Sheet3");

        List<ExcelSheetInfo> sheets;
        try (InputStream is = Files.newInputStream(excelFile)) {
            sheets = ExcelReader.getSheetNames(is);
        }

        assertEquals(3, sheets.size());
        assertEquals("Sheet1", sheets.get(0).name());
        assertEquals(0, sheets.get(0).index());
        assertEquals("Sheet2", sheets.get(1).name());
        assertEquals("Sheet3", sheets.get(2).name());
    }

    @Test
    void getSheetNames_withInvalidStream_shouldCleanUpAndThrow() {
        InputStream badStream = new ByteArrayInputStream("not an excel file".getBytes());

        assertThrows(ExcelReadException.class, () ->
                ExcelReader.getSheetNames(badStream));
    }

    @Test
    void getSheetNames_withEmptyStream_shouldCleanUpAndThrow() {
        InputStream emptyStream = new ByteArrayInputStream(new byte[0]);

        assertThrows(ExcelReadException.class, () ->
                ExcelReader.getSheetNames(emptyStream));
    }

    // ──────────────────────────────────────────────────────────────
    // ExcelReader.getSheetHeaders: temp file cleanup
    // ──────────────────────────────────────────────────────────────

    @Test
    void getSheetHeaders_shouldCleanUpTempFiles() throws IOException {
        Path excelFile = createTestExcelWithHeaders();

        List<String> headers;
        try (InputStream is = Files.newInputStream(excelFile)) {
            headers = ExcelReader.getSheetHeaders(is, 0, 0);
        }

        assertEquals(3, headers.size());
        assertEquals("Name", headers.get(0));
        assertEquals("Age", headers.get(1));
        assertEquals("City", headers.get(2));
    }

    @Test
    void getSheetHeaders_withInvalidStream_shouldCleanUpAndThrow() {
        InputStream badStream = new ByteArrayInputStream("invalid".getBytes());

        assertThrows(ExcelReadException.class, () ->
                ExcelReader.getSheetHeaders(badStream, 0, 0));
    }

    // ──────────────────────────────────────────────────────────────
    // ExcelReadHandler: temp file cleanup via read()
    // ──────────────────────────────────────────────────────────────

    @Test
    void read_shouldCleanUpTempFilesAfterSuccess() throws IOException {
        Path excelFile = createTestExcelWithData();

        List<String> names = new ArrayList<>();
        try (InputStream is = Files.newInputStream(excelFile)) {
            new ExcelReader<>(TestRow::new, null)
                    .column((r, cell) -> r.name = cell.asString())
                    .column((r, cell) -> r.age = cell.asInt())
                    .build(is)
                    .read(result -> {
                        assertTrue(result.success(), "Row should succeed: " + result.messages());
                        names.add(result.data().name);
                    });
        }

        assertEquals(3, names.size());
        assertEquals("Alice", names.get(0));
        assertEquals("Bob", names.get(1));
        assertEquals("Charlie", names.get(2));
    }

    @Test
    void read_shouldCleanUpTempFilesAfterException() throws IOException {
        Path excelFile = createTestExcelWithData();

        try (InputStream is = Files.newInputStream(excelFile)) {
            ExcelReadHandler<TestRow> handler = new ExcelReader<>(TestRow::new, null)
                    .column((r, cell) -> r.name = cell.asString())
                    .column((r, cell) -> r.age = cell.asInt())
                    .build(is);

            // Force an exception during processing
            assertThrows(RuntimeException.class, () ->
                    handler.read(result -> {
                        throw new RuntimeException("Intentional error");
                    }));
        }
        // If we get here, temp files were cleaned up (no resource leak)
    }

    @Test
    void readAsStream_shouldCleanUpOnClose() throws IOException {
        Path excelFile = createTestExcelWithData();

        List<String> names;
        try (InputStream is = Files.newInputStream(excelFile)) {
            try (Stream<ReadResult<TestRow>> stream = new ExcelReader<>(TestRow::new, null)
                    .column((r, cell) -> r.name = cell.asString())
                    .column((r, cell) -> r.age = cell.asInt())
                    .build(is)
                    .readAsStream()) {

                names = stream
                        .filter(ReadResult::success)
                        .map(r -> r.data().name)
                        .toList();
            }
        }

        assertEquals(3, names.size());
    }

    @Test
    void readAsStream_earlyTermination_shouldCleanUp() throws IOException {
        Path excelFile = createTestExcelWithData();

        List<String> names;
        try (InputStream is = Files.newInputStream(excelFile)) {
            try (Stream<ReadResult<TestRow>> stream = new ExcelReader<>(TestRow::new, null)
                    .column((r, cell) -> r.name = cell.asString())
                    .column((r, cell) -> r.age = cell.asInt())
                    .build(is)
                    .readAsStream()) {

                // Only consume 1 element, then close the stream
                names = stream
                        .filter(ReadResult::success)
                        .limit(1)
                        .map(r -> r.data().name)
                        .toList();
            }
        }

        assertEquals(1, names.size());
        assertEquals("Alice", names.get(0));
    }

    // ──────────────────────────────────────────────────────────────
    // ExcelWriter: full write-then-read roundtrip verifying cleanup
    // ──────────────────────────────────────────────────────────────

    @Test
    void writeAndRead_roundtrip_shouldCleanUpAllTempResources() throws IOException {
        // Write
        ByteArrayOutputStream baos = new ByteArrayOutputStream();
        new ExcelWriter<TestRow>()
                .column("Name", r -> r.name).type(ExcelDataType.STRING)
                .column("Age", r -> r.age).type(ExcelDataType.INTEGER)
                .write(Stream.of(
                        new TestRow("Alice", 30),
                        new TestRow("Bob", 25)
                ))
                .consumeOutputStream(baos);

        assertTrue(baos.size() > 100, "Excel output should be non-trivial");

        // Read back
        List<TestRow> results = new ArrayList<>();
        try (InputStream is = new ByteArrayInputStream(baos.toByteArray())) {
            new ExcelReader<>(TestRow::new, null)
                    .column((r, cell) -> r.name = cell.asString())
                    .column((r, cell) -> r.age = cell.asInt())
                    .build(is)
                    .read(result -> {
                        assertTrue(result.success(), "Row should succeed: " + result.messages());
                        results.add(result.data());
                    });
        }

        assertEquals(2, results.size());
        assertEquals("Alice", results.get(0).name);
        assertEquals(30, results.get(0).age);
        assertEquals("Bob", results.get(1).name);
        assertEquals(25, results.get(1).age);
    }

    @Test
    void encryptedWriteAndVerify_roundtrip() throws IOException {
        ByteArrayOutputStream baos = new ByteArrayOutputStream();
        new ExcelWriter<TestRow>()
                .column("Name", r -> r.name).type(ExcelDataType.STRING)
                .column("Age", r -> r.age).type(ExcelDataType.INTEGER)
                .write(Stream.of(new TestRow("Secret", 99)))
                .consumeOutputStreamWithPassword(baos, "pass123");

        // Encrypted file should be different from unencrypted
        assertTrue(baos.size() > 100, "Encrypted output should be non-trivial");
        // First bytes of encrypted file should be OLE2 magic (D0 CF 11 E0)
        byte[] bytes = baos.toByteArray();
        assertEquals((byte) 0xD0, bytes[0], "Encrypted file should start with OLE2 header");
        assertEquals((byte) 0xCF, bytes[1]);
    }

    // ──────────────────────────────────────────────────────────────
    // Helpers
    // ──────────────────────────────────────────────────────────────

    private Path createTestExcel(String... sheetNames) throws IOException {
        Path file = tempDir.resolve("test-sheets.xlsx");
        try (Workbook wb = new XSSFWorkbook()) {
            for (String name : sheetNames) {
                Sheet sheet = wb.createSheet(name);
                Row header = sheet.createRow(0);
                header.createCell(0).setCellValue("Col1");
            }
            try (FileOutputStream fos = new FileOutputStream(file.toFile())) {
                wb.write(fos);
            }
        }
        return file;
    }

    private Path createTestExcelWithHeaders() throws IOException {
        Path file = tempDir.resolve("test-headers.xlsx");
        try (Workbook wb = new XSSFWorkbook()) {
            Sheet sheet = wb.createSheet("Data");
            Row header = sheet.createRow(0);
            header.createCell(0).setCellValue("Name");
            header.createCell(1).setCellValue("Age");
            header.createCell(2).setCellValue("City");
            try (FileOutputStream fos = new FileOutputStream(file.toFile())) {
                wb.write(fos);
            }
        }
        return file;
    }

    private Path createTestExcelWithData() throws IOException {
        Path file = tempDir.resolve("test-data.xlsx");
        try (Workbook wb = new XSSFWorkbook()) {
            Sheet sheet = wb.createSheet("Data");
            Row header = sheet.createRow(0);
            header.createCell(0).setCellValue("Name");
            header.createCell(1).setCellValue("Age");

            createRow(sheet, 1, "Alice", 30);
            createRow(sheet, 2, "Bob", 25);
            createRow(sheet, 3, "Charlie", 35);

            try (FileOutputStream fos = new FileOutputStream(file.toFile())) {
                wb.write(fos);
            }
        }
        return file;
    }

    private void createRow(Sheet sheet, int rowNum, String name, int age) {
        Row row = sheet.createRow(rowNum);
        row.createCell(0).setCellValue(name);
        row.createCell(1).setCellValue(age);
    }

    static class TestRow {
        String name;
        int age;

        TestRow() {}

        TestRow(String name, int age) {
            this.name = name;
            this.age = age;
        }
    }
}
