package io.github.dornol.excelkit.excel;

import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.junit.jupiter.api.AfterEach;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import java.io.ByteArrayOutputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;

import static org.junit.jupiter.api.Assertions.*;

/**
 * Tests for {@link ExcelHandler} class.
 */
class ExcelHandlerTest {

    private SXSSFWorkbook workbook;
    private ExcelHandler handler;
    private ByteArrayOutputStream outputStream;

    @TempDir
    Path tempDir;

    @BeforeEach
    void setUp() {
        workbook = new SXSSFWorkbook();
        handler = new ExcelHandler(workbook);
        outputStream = new ByteArrayOutputStream();
    }

    @AfterEach
    void tearDown() throws IOException {
        outputStream.close();
        // Note: workbook is closed by ExcelHandler after consumption
    }

    @Test
    void constructor_shouldCreateInstanceWithValidParameters() {
        // Assert
        assertNotNull(handler, "Handler should be created with valid workbook");
    }

    @Test
    void write_shouldWriteWorkbookToOutputStream() throws IOException {
        // Arrange
        createSampleWorkbookContent();

        // Act
        handler.writeTo(outputStream);

        // Assert
        byte[] excelBytes = outputStream.toByteArray();
        assertTrue(excelBytes.length > 0, "Excel file should have content");
    }

    @Test
    void write_shouldThrowExceptionWhenAlreadyConsumed() throws IOException {
        // Arrange
        handler.writeTo(outputStream);
        ByteArrayOutputStream secondOutputStream = new ByteArrayOutputStream();

        // Act & Assert
        assertThrows(ExcelWriteException.class, () -> {
            handler.writeTo(secondOutputStream);
        }, "write should throw ExcelWriteException when already consumed");

        secondOutputStream.close();
    }

    @Test
    void writeToWithPassword_shouldWriteEncryptedWorkbook() throws IOException {
        // Arrange
        createSampleWorkbookContent();
        Path excelFile = tempDir.resolve("encrypted.xlsx");

        // Act
        try (FileOutputStream fos = new FileOutputStream(excelFile.toFile())) {
            handler.writeTo(fos, "test123");
        }

        // Assert
        assertTrue(Files.exists(excelFile), "Encrypted Excel file should be created");
        assertTrue(Files.size(excelFile) > 0, "Encrypted Excel file should have content");
    }

    @Test
    void writeToWithPassword_shouldThrowExceptionWhenAlreadyConsumed() throws IOException {
        // Arrange
        handler.writeTo(outputStream);
        ByteArrayOutputStream secondOutputStream = new ByteArrayOutputStream();

        // Act & Assert
        assertThrows(ExcelWriteException.class, () -> {
            handler.writeTo(secondOutputStream, "test123");
        }, "writeToWithPassword should throw ExcelWriteException when already consumed");

        secondOutputStream.close();
    }

    @Test
    void writeToWithPassword_shouldThrowExceptionWithNullPassword() {
        // Act & Assert
        assertThrows(IllegalArgumentException.class, () -> {
            handler.writeTo(outputStream, (String) null);
        }, "writeToWithPassword should throw IllegalArgumentException with null password");
    }

    @Test
    void writeToWithPassword_shouldThrowExceptionWithNullCharArrayPassword() {
        // Act & Assert
        assertThrows(IllegalArgumentException.class, () -> {
            handler.writeTo(outputStream, (char[]) null);
        }, "writeToWithPassword should throw IllegalArgumentException with null char[] password");
    }

    @Test
    void writeToWithPassword_shouldThrowExceptionWithEmptyCharArrayPassword() {
        // Act & Assert
        assertThrows(IllegalArgumentException.class, () -> {
            handler.writeTo(outputStream, new char[0]);
        }, "writeToWithPassword should throw IllegalArgumentException with empty char[] password");
    }

    @Test
    void writeToWithPassword_charArray_shouldWriteEncryptedWorkbook() throws IOException {
        // Arrange
        createSampleWorkbookContent();
        Path excelFile = tempDir.resolve("encrypted-char.xlsx");
        char[] password = "test123".toCharArray();

        // Act
        try (FileOutputStream fos = new FileOutputStream(excelFile.toFile())) {
            handler.writeTo(fos, password);
        }

        // Assert
        assertTrue(Files.exists(excelFile), "Encrypted Excel file should be created");
        assertTrue(Files.size(excelFile) > 0, "Encrypted Excel file should have content");
        // Password should be zeroed out after use
        for (char c : password) {
            assertEquals('\0', c, "Password char array should be zeroed after use");
        }
    }

    @Test
    void write_withPassword_shouldProduceOLE2Format() throws IOException {
        // Arrange
        SXSSFWorkbook pwWorkbook = new SXSSFWorkbook();
        pwWorkbook.createSheet("Test").createRow(0).createCell(0).setCellValue("Test");
        ExcelHandler pwHandler = new ExcelHandler(pwWorkbook, "test123");
        ByteArrayOutputStream out = new ByteArrayOutputStream();

        // Act
        pwHandler.writeTo(out);

        // Assert - verify OLE2 magic bytes (encrypted), not ZIP (unencrypted)
        byte[] bytes = out.toByteArray();
        assertEquals((byte) 0xD0, bytes[0], "Should be OLE2 encrypted format");
        assertEquals((byte) 0xCF, bytes[1], "Should be OLE2 encrypted format");
    }

    @Test
    void write_withoutPassword_shouldProduceZipFormat() throws IOException {
        // Arrange
        createSampleWorkbookContent();

        // Act
        handler.writeTo(outputStream);

        // Assert - verify ZIP magic bytes (unencrypted OOXML)
        byte[] bytes = outputStream.toByteArray();
        assertEquals((byte) 0x50, bytes[0], "Should be ZIP/OOXML format");
        assertEquals((byte) 0x4B, bytes[1], "Should be ZIP/OOXML format");
    }

    @Test
    void writeToWithPassword_charArray_blankPassword_shouldThrow() {
        // Arrange
        char[] blankPassword = {' ', ' ', ' '};

        // Act & Assert
        assertThrows(IllegalArgumentException.class, () -> {
            handler.writeTo(outputStream, blankPassword);
        }, "writeToWithPassword should throw for blank char[] password");
    }

    @Test
    void writeToPath_withPassword_roundTripViaExcelReader() throws IOException {
        // End-to-end: write encrypted file to disk, then ExcelReader with matching password
        // should decrypt and read the data back verbatim. Proves the file isn't just OLE2-shaped
        // but actually contains the data we wrote.
        SXSSFWorkbook wb = new SXSSFWorkbook();
        var sheet = wb.createSheet("Data");
        var r0 = sheet.createRow(0);
        r0.createCell(0).setCellValue("Name");
        r0.createCell(1).setCellValue("Age");
        var r1 = sheet.createRow(1);
        r1.createCell(0).setCellValue("Alice");
        r1.createCell(1).setCellValue(30);
        var r2 = sheet.createRow(2);
        r2.createCell(0).setCellValue("Bob");
        r2.createCell(1).setCellValue(25);

        ExcelHandler h = new ExcelHandler(wb);
        Path excelFile = tempDir.resolve("round-trip.xlsx");

        h.writeTo(excelFile, "roundTripPw");

        // Read back with the password
        java.util.List<String> names = new java.util.ArrayList<>();
        java.util.List<Integer> ages = new java.util.ArrayList<>();
        ExcelReader.mapping(row -> {
                    names.add(row.get("Name").asString());
                    ages.add(row.get("Age").asInt());
                    return row;
                })
                .password("roundTripPw")
                .read(Files.newInputStream(excelFile), r -> {});

        assertEquals(java.util.List.of("Alice", "Bob"), names);
        assertEquals(java.util.List.of(30, 25), ages);
    }

    @Test
    void writeToPath_withPassword_wrongPasswordRejected() throws IOException {
        createSampleWorkbookContent();
        Path excelFile = tempDir.resolve("wrong-pw.xlsx");
        handler.writeTo(excelFile, "correct");

        var ex = assertThrows(ExcelReadException.class, () ->
                ExcelReader.mapping(row -> row)
                        .password("wrong")
                        .read(Files.newInputStream(excelFile), r -> {}));
        assertTrue(ex.getMessage().toLowerCase().contains("invalid password"),
                "Expected 'invalid password' in message: " + ex.getMessage());
    }

    @Test
    void writeToPathWithPassword_shouldWriteEncryptedFile() throws IOException {
        // Arrange
        createSampleWorkbookContent();
        Path excelFile = tempDir.resolve("path-encrypted.xlsx");

        // Act
        handler.writeTo(excelFile, "test123");

        // Assert
        assertTrue(Files.exists(excelFile));
        assertTrue(Files.size(excelFile) > 0);
        // Verify OLE2 format (encrypted)
        byte[] bytes = Files.readAllBytes(excelFile);
        assertEquals((byte) 0xD0, bytes[0]);
        assertEquals((byte) 0xCF, bytes[1]);
    }

    @Test
    void writeToPathWithCharArrayPassword_shouldWriteEncryptedAndZero() throws IOException {
        // Arrange
        createSampleWorkbookContent();
        Path excelFile = tempDir.resolve("path-encrypted-char.xlsx");
        char[] password = "test123".toCharArray();

        // Act
        handler.writeTo(excelFile, password);

        // Assert
        assertTrue(Files.exists(excelFile));
        assertTrue(Files.size(excelFile) > 0);
        for (char c : password) {
            assertEquals('\0', c, "password char[] must be zeroed");
        }
    }

    @Test
    void writeToPathWithPassword_nullOrBlank_throws() {
        Path target = tempDir.resolve("x.xlsx");
        assertThrows(IllegalArgumentException.class, () -> handler.writeTo(target, (String) null));
        assertThrows(IllegalArgumentException.class, () -> handler.writeTo(target, ""));
        assertThrows(IllegalArgumentException.class, () -> handler.writeTo(target, "   "));
        assertThrows(IllegalArgumentException.class, () -> handler.writeTo(target, (char[]) null));
        assertThrows(IllegalArgumentException.class, () -> handler.writeTo(target, new char[0]));
        assertThrows(IllegalArgumentException.class, () -> handler.writeTo(target, new char[]{' ', ' '}));
    }

    @Test
    void writeToPathWithPassword_afterConsumed_throws() throws IOException {
        handler.writeTo(outputStream);
        Path target = tempDir.resolve("after.xlsx");
        assertThrows(ExcelWriteException.class, () -> handler.writeTo(target, "test123"));
    }

    @Test
    void writeToPathWithPassword_whenHandlerPasswordSet_throwsIllegalState() {
        SXSSFWorkbook wb2 = new SXSSFWorkbook();
        wb2.createSheet("S").createRow(0).createCell(0).setCellValue("x");
        ExcelHandler h = new ExcelHandler(wb2, "pw1");
        Path target = tempDir.resolve("conflict.xlsx");
        assertThrows(IllegalStateException.class, () -> h.writeTo(target, "pw2"));
    }

    @Test
    void write_shouldCloseWorkbookAfterWriting() throws IOException {
        // Arrange
        SXSSFWorkbook testWorkbook = new SXSSFWorkbook();
        ExcelHandler testHandler = new ExcelHandler(testWorkbook);
        
        // Act
        testHandler.writeTo(outputStream);
        
        // Assert - attempting to use the workbook after it's closed should throw an exception
        assertThrows(IOException.class, () -> {
            testWorkbook.write(outputStream); // retry after already closed
        });
    }

    /**
     * Helper method to create sample content in the workbook for testing.
     */
    private void createSampleWorkbookContent() {
        workbook.createSheet("Test Sheet").createRow(0).createCell(0).setCellValue("Test Data");
    }
}