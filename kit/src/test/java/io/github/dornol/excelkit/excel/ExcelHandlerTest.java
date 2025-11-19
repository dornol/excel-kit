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
    void consumeOutputStream_shouldWriteWorkbookToOutputStream() throws IOException {
        // Arrange
        createSampleWorkbookContent();

        // Act
        handler.consumeOutputStream(outputStream);

        // Assert
        byte[] excelBytes = outputStream.toByteArray();
        assertTrue(excelBytes.length > 0, "Excel file should have content");
    }

    @Test
    void consumeOutputStream_shouldThrowExceptionWhenAlreadyConsumed() throws IOException {
        // Arrange
        handler.consumeOutputStream(outputStream);
        ByteArrayOutputStream secondOutputStream = new ByteArrayOutputStream();

        // Act & Assert
        assertThrows(IllegalStateException.class, () -> {
            handler.consumeOutputStream(secondOutputStream);
        }, "consumeOutputStream should throw IllegalStateException when already consumed");

        secondOutputStream.close();
    }

    @Test
    void consumeOutputStreamWithPassword_shouldWriteEncryptedWorkbook() throws IOException {
        // Arrange
        createSampleWorkbookContent();
        Path excelFile = tempDir.resolve("encrypted.xlsx");

        // Act
        try (FileOutputStream fos = new FileOutputStream(excelFile.toFile())) {
            handler.consumeOutputStreamWithPassword(fos, "test123");
        }

        // Assert
        assertTrue(Files.exists(excelFile), "Encrypted Excel file should be created");
        assertTrue(Files.size(excelFile) > 0, "Encrypted Excel file should have content");
    }

    @Test
    void consumeOutputStreamWithPassword_shouldThrowExceptionWhenAlreadyConsumed() throws IOException {
        // Arrange
        handler.consumeOutputStream(outputStream);
        ByteArrayOutputStream secondOutputStream = new ByteArrayOutputStream();

        // Act & Assert
        assertThrows(IllegalStateException.class, () -> {
            handler.consumeOutputStreamWithPassword(secondOutputStream, "test123");
        }, "consumeOutputStreamWithPassword should throw IllegalStateException when already consumed");

        secondOutputStream.close();
    }

    @Test
    void consumeOutputStreamWithPassword_shouldThrowExceptionWithNullPassword() {
        // Act & Assert
        assertThrows(IllegalArgumentException.class, () -> {
            handler.consumeOutputStreamWithPassword(outputStream, null);
        }, "consumeOutputStreamWithPassword should throw IllegalArgumentException with null password");
    }

    @Test
    void consumeOutputStream_shouldCloseWorkbookAfterWriting() throws IOException {
        // Arrange
        SXSSFWorkbook testWorkbook = new SXSSFWorkbook();
        ExcelHandler testHandler = new ExcelHandler(testWorkbook);
        
        // Act
        testHandler.consumeOutputStream(outputStream);
        
        // Assert - attempting to use the workbook after it's closed should throw an exception
        assertThrows(IOException.class, () -> {
            testWorkbook.write(outputStream); // 이미 close 후 재시도
        });
    }

    /**
     * Helper method to create sample content in the workbook for testing.
     */
    private void createSampleWorkbookContent() {
        workbook.createSheet("Test Sheet").createRow(0).createCell(0).setCellValue("Test Data");
    }
}